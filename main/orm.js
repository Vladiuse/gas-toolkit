const DEFAULT_HEADER_ROW = 1;

/**
 * Column numbers of a sheet, looked up by the text in its header row.
 *
 * The header is read once on the first call and served from memory afterwards,
 * so reuse the same instance instead of creating one per lookup.
 */
class SheetHeaders {
    /**
     * @param {Sheet} sheet - Sheet to read the header from.
     * @param {number} [headerRow=1] - Row holding the column titles.
     */
    constructor(sheet, headerRow = DEFAULT_HEADER_ROW) {
        this._sheet = sheet;
        this._headerRow = headerRow;
        this._columnByTitle = new Map();
        this._titleByColumn = new Map();
        this._titleDuplicates = [];
        this._isHeaderLoaded = false;
    }

    /**
     * Column number of the title, counted from 1 as the sheet counts them.
     *
     * @param {string} title - Header text, matched after trimming spaces.
     * @returns {number} The column number.
     * @throws {Error} When the sheet has no such header.
     */
    get(title) {
        const column = this.find(title);

        if (column === null) {
            throw new Error(`SheetHeaders: на листе "${this._sheet.getName()}" нет колонки "${title}"`);
        }

        return column;
    }

    /**
     * Column number of the title, or null when the sheet has no such header.
     *
     * @param {string} title - Header text, matched after trimming spaces.
     * @returns {number|null}
     */
    find(title) {
        const column = this.columnByTitle.get(this._normalizeTitle(title));

        return column === undefined ? null : column;
    }

    /**
     * Whether the sheet has a column with this title.
     *
     * @param {string} title - Header text, matched after trimming spaces.
     * @returns {boolean}
     */
    has(title) {
        return this.columnByTitle.has(this._normalizeTitle(title));
    }

    /**
     * Titles the sheet has no column for, in the order they were asked about.
     *
     * @param {string[]} titles - Header texts, matched after trimming spaces.
     * @returns {string[]} Empty when every title is present.
     */
    findMissing(titles) {
        return titles.filter((title) => !this.has(title));
    }

    /**
     * Check that the sheet has a column for every title.
     *
     * @param {string[]} titles - Header texts, matched after trimming spaces.
     * @throws {Error} When at least one title is missing, listing all of them.
     */
    require(titles) {
        const missing = this.findMissing(titles);

        if (missing.length > 0) {
            const sheetName = this._sheet.getName();
            const missingList = missing.map((title) => `"${title}"`).join(", ");

            throw new Error(`SheetHeaders: на листе "${sheetName}" нет колонок: ${missingList}`);
        }
    }

    /**
     * Column number for every header in a set, keyed the same way the set is.
     *
     * Takes short names mapped to header texts and gives back the same short
     * names mapped to column numbers:
     *
     *     const TITLES = {
     *         partnerId: "Partner ID",
     *         playerId: "Player ID",
     *     };
     *
     *     const columns = headers.getColumns(TITLES);
     *     // { partnerId: 1, playerId: 4 }
     *
     *     columns.playerId;  // 4, the column "Player ID" sits in
     *
     * @param {Object<string, string>} titleByKey - Short name to header text.
     * @returns {Object<string, number>} Short name to column number.
     * @throws {Error} When at least one header is missing, listing all of them.
     */
    getColumns(titleByKey) {
        this.require(Object.values(titleByKey));

        const columnByKey = {};

        Object.entries(titleByKey).forEach(([key, title]) => {
            columnByKey[key] = this.get(title);
        });

        return columnByKey;
    }

    /**
     * Same as getColumns, but each column is looked up when it is read.
     *
     * A sheet older than the code may be missing a column the code knows
     * about. getColumns refuses such a sheet outright, while this one serves
     * every column that is there and only fails on the one that is not:
     *
     *     const columns = headers.getLazyColumns(TITLES);
     *
     *     columns.playerId;  // 4, read from the sheet right here
     *     columns.hold;      // throws when the sheet has no "Холд" column
     *
     * @param {Object<string, string>} titleByKey - Short name to header text.
     * @returns {Object<string, number>} Short name to column number.
     * @throws {Error} On reading a key the set does not hold, or a header the
     *     sheet does not have.
     */
    getLazyColumns(titleByKey) {
        return new Proxy(titleByKey, {
            get: (target, key) => {
                const title = target[key];

                if (title === undefined) {
                    throw new Error(`SheetHeaders: неизвестный ключ колонки "${String(key)}"`);
                }

                return this.get(title);
            },
        });
    }

    /**
     * A row laid out the way the sheet expects it, values put under their own
     * headers.
     *
     * The row runs from the first column to the last one the header names, so
     * that writing it starts at column 1 and covers the whole header:
     *
     *     // шапка: Partner ID | Player ID | Оплата | Холд
     *     headers.toRow({ "Player ID": "p-1", "Оплата": 100 });
     *     // [null, "p-1", 100, null]
     *
     * Columns the set says nothing about are left null, which writes an empty
     * cell. That suits a row being added, where there is nothing to keep, and
     * not a row being changed in place, where null erases what is there.
     *
     * A header the sheet does not have is an error, since it usually means a
     * typo or a sheet older than the code. Moving rows between two sheets is
     * the case where it does not: there the sets of columns differ on purpose,
     * and skipStrangers passes such a value over instead.
     *
     *     // лист без колонок "Player country" и "FTD date"
     *     headers.toRow({ "Холд": "да", "Player country": "UA", "FTD date": "" });
     *     // Error: ... некуда записать: "Player country", "FTD date"
     *
     *     headers.toRow({ "Холд": "да", "Player country": "UA" }, true);
     *     // [..., "да"]
     *
     * @param {Object<string, *>} valueByTitle - Header text to the value to
     *     put under it.
     * @param {boolean} [skipStrangers=false] - Pass over a value whose header
     *     the sheet has no column for, instead of refusing the whole row.
     * @returns {*[]} As wide as the header, ready for setValues or appendRow.
     * @throws {Error} When the sheet has no column for one of the headers and
     *     skipStrangers is off.
     */
    toRow(valueByTitle, skipStrangers = false) {
        const row = new Array(this.lastColumn).fill(null);

        if (!skipStrangers) {
            const missing = this.findMissing(Object.keys(valueByTitle));

            if (missing.length > 0) {
                const sheetName = this._sheet.getName();
                const missingList = missing.map((title) => `"${title}"`).join(", ");

                throw new Error(`SheetHeaders: на листе "${sheetName}" некуда записать: ${missingList}`);
            }
        }

        Object.entries(valueByTitle).forEach(([title, value]) => {
            const column = this.find(title);

            if (column !== null) {
                row[column - 1] = value;
            }
        });

        return row;
    }

    /**
     * Several rows laid out the same way toRow lays out one.
     *
     * Every row is as wide as the header, whatever each set of values holds,
     * so the result goes into a single setValues:
     *
     *     headers.toRows([
     *         { "Player ID": "p-1", "Оплата": 100 },
     *         { "Player ID": "p-2" },
     *     ]);
     *     // [[null, "p-1", 100, null],
     *     //  [null, "p-2", null, null]]
     *
     * @param {Object<string, *>[]} valueByTitleList - One set of values per
     *     row, in the order the rows are to be written.
     * @param {boolean} [skipStrangers=false] - Pass over a value whose header
     *     the sheet has no column for, instead of refusing the whole row.
     * @returns {*[][]} Empty when the list is, ready for setValues otherwise.
     * @throws {Error} When the sheet has no column for one of the headers and
     *     skipStrangers is off.
     */
    toRows(valueByTitleList, skipStrangers = false) {
        return valueByTitleList.map((valueByTitle) => this.toRow(valueByTitle, skipStrangers));
    }

    /**
     * A row read off the sheet, named by the headers it sits under.
     *
     * The other way round of toRow, for a row that getValues just brought
     * back:
     *
     *     headers.fromRow([42, "p-1", 100, ""]);
     *     // { "Partner ID": 42, "Player ID": "p-1", "Оплата": 100, "Холд": "" }
     *
     * Every header gets a key, whatever the row holds, so a caller can read
     * one without checking it is there first. A row shorter than the header,
     * which is what getValues gives for a range that stops early, leaves the
     * headers past its end undefined. Columns with no header are skipped, and
     * of two columns sharing a header the leftmost one wins, as it does
     * everywhere else here.
     *
     * @param {*[]} values - Cell values, left to right from column 1.
     * @returns {Object<string, *>} Header text to the value under it.
     */
    fromRow(values) {
        const valueByTitle = {};

        this.columnByTitle.forEach((column, title) => {
            valueByTitle[title] = values[column - 1];
        });

        return valueByTitle;
    }

    /**
     * Several rows read off the sheet, each named the way fromRow names one.
     *
     *     headers.fromRows([
     *         [42, "p-1", 100, ""],
     *         [43, "p-2", 0, "да"],
     *     ]);
     *     // [{ "Partner ID": 42, "Player ID": "p-1", "Оплата": 100, "Холд": "" },
     *     //  { "Partner ID": 43, "Player ID": "p-2", "Оплата": 0, "Холд": "да" }]
     *
     * @param {*[][]} rows - Rows as getValues returns them.
     * @returns {Object<string, *>[]} One set of named values per row.
     */
    fromRows(rows) {
        return rows.map((values) => this.fromRow(values));
    }

    /**
     * Header text to its column number, read from the sheet on first use.
     *
     * @returns {Map<string, number>}
     */
    get columnByTitle() {
        if (!this._isHeaderLoaded) {
            this._loadHeader();
        }

        return this._columnByTitle;
    }

    /**
     * Column number to its header text, read from the sheet on first use.
     *
     * The other way round of columnByTitle, with one difference: when the same
     * title sits in two columns, columnByTitle keeps the leftmost one, while
     * this map knows the title of both.
     *
     * @returns {Map<number, string>}
     */
    get titleByColumn() {
        if (!this._isHeaderLoaded) {
            this._loadHeader();
        }

        return this._titleByColumn;
    }

    /**
     * Header text of the column, as it is stored after trimming spaces.
     *
     * @param {number} column - Column number, counted from 1.
     * @returns {string|null} Null when the column has no header.
     */
    getTitle(column) {
        const title = this.titleByColumn.get(column);

        return title === undefined ? null : title;
    }

    /**
     * Number of the rightmost column that has a header.
     *
     * @returns {number} Zero when the header row is empty.
     */
    get lastColumn() {
        const columns = Array.from(this.titleByColumn.keys());

        return columns.length === 0 ? 0 : columns[columns.length - 1];
    }

    /**
     * Read the header row, keeping only the cells that hold a title.
     *
     * The row is asked for across the whole width of the sheet, so that the
     * one call brings back every header there is; the empty cells past the
     * last of them are dropped here rather than measured beforehand.
     *
     * @returns {{title: string, column: number}[]} Left to right.
     */
    _readHeaderCells() {
        const width = this._sheet.getMaxColumns();
        const values = this._sheet.getRange(this._headerRow, 1, 1, width).getValues()[0];
        const cells = [];

        for (let i = 0; i < values.length; i++) {
            const title = this._normalizeTitle(values[i]);

            if (title) {
                cells.push({ title: title, column: i + 1 });
            }
        }

        return cells;
    }

    /**
     * Read the header row into the maps and mark it loaded.
     *
     * A title found in two columns belongs to the leftmost of them, which is
     * the one a lookup by title answers with; the rest are only recorded, in
     * duplicates, for whoever cares to look.
     */
    _loadHeader() {
        const cells = this._readHeaderCells();

        this._columnByTitle = new Map();
        this._titleByColumn = new Map();
        this._titleDuplicates = [];

        cells.forEach((cell) => {
            if (this._columnByTitle.has(cell.title)) {
                this._titleDuplicates.push(cell);
            } else {
                this._columnByTitle.set(cell.title, cell.column);
            }

            this._titleByColumn.set(cell.column, cell.title);
        });

        this._isHeaderLoaded = true;
    }

    /**
     * Bring a title to the form it is stored and looked up by.
     *
     * Headers are often typed with a line break inside the cell, which no
     * reasonable lookup string would repeat: every run of whitespace, line
     * breaks included, becomes a single space.
     */
    _normalizeTitle(title) {
        return String(title).replace(/\s+/g, " ").trim();
    }
}

/**
 * One row of a sheet, its cells named by short attribute names.
 *
 * A record belongs to the manager that read it: the manager holds the titles
 * saying which header each attribute stands for, and the record asks it for
 * them whenever it has to name its values the way the sheet names them. So a
 * record is not built by hand, it comes from a manager:
 *
 *     const player = players.getAll().records[0];
 *
 *     player.partnerId;   // 42
 *     player.row;         // 5, the row it was read from
 *     player.toValues();  // { "Partner ID": 42, ... }
 *
 * Every attribute the titles name is set, whatever the values hold: one with
 * no value becomes undefined rather than being left out.
 */
class Model {
    /**
     * @param {Manager} manager - Manager the record was read by, and whose
     *     titles name its attributes.
     * @param {Object<string, *>} [valueByTitle] - Header text to its value.
     * @param {number} [row=0] - Sheet row the record was read from, 0 when it
     *     was not read from one.
     */
    constructor(manager, valueByTitle = {}, row = 0) {
        this._manager = manager;
        this.row = row;

        Object.entries(manager.titles).forEach(([name, title]) => {
            this[name] = valueByTitle[title];
        });
    }

    /**
     * Manager the record was read by.
     *
     * @returns {Manager}
     */
    get manager() {
        return this._manager;
    }

    /**
     * The record as values under their own headers, the way a sheet takes them.
     *
     * @returns {Object<string, *>} Header text to its value.
     */
    toValues() {
        const valueByTitle = {};

        Object.entries(this._manager.titles).forEach(([name, title]) => {
            valueByTitle[title] = this[name];
        });

        return valueByTitle;
    }

    /**
     * Write the record back to its row.
     *
     * @returns {Model} The record itself.
     */
    save() {
        // TODO: написать. Решить, пишутся ли все атрибуты подряд или только
        // изменённые — для второго запись должна помнить, какой её прочитали.
        throw new Error(`${this.constructor.name}.save: не реализовано`);
    }

    /**
     * Remove the row the record was read from.
     *
     * @returns {void}
     */
    delete() {
        // TODO: написать через Manager.delete. Решить, чем становится запись
        // после удаления: строки под ней уже съехали вверх, так что row у
        // всех прочитанных записей врёт.
        throw new Error(`${this.constructor.name}.delete: не реализовано`);
    }

    toString() {
        return `${this.constructor.name}(row=${this.row})`;
    }
}

/**
 * How a sheet is laid out: what its records are made of and where they sit.
 *
 * Written down once, next to the titles it names, and handed to every manager
 * over that sheet:
 *
 *     const PLAYERS_META = new SheetMeta({
 *         model: Player,
 *         titles: PLAYERS_COLUMN_TITLES,
 *         markerAttribute: "partnerId",
 *     });
 *
 *     const players = new Manager(sheet, PLAYERS_META);
 *
 * The settings are checked here, when the meta is made, rather than later
 * when a manager first touches the sheet: a typo in markerAttribute is then
 * an error on loading the file, not a puzzle during a run.
 */
// eslint-disable-next-line no-unused-vars
class SheetMeta {
    /**
     * @param {Object} settings
     * @param {Object<string, string>} settings.titles - Attribute name to the
     *     header text its column sits under. This is what the records are
     *     made of: each name becomes an attribute on every record read.
     * @param {string} settings.markerAttribute - Attribute telling a row with
     *     a record on it from an empty one: a row whose cell in that column
     *     is empty holds no record.
     * @param {typeof Model} [settings.model=Model] - Class the records are
     *     made of.
     * @param {number} [settings.headerRow=1] - Row holding the column titles.
     * @param {number} [settings.firstDataRow] - First row a record can sit
     *     on. Left out, it is the row right under the header, which is where
     *     records begin unless a sheet keeps something else between the
     *     header and them.
     * @throws {Error} When titles or markerAttribute are missing, or the
     *     marker names no title.
     */
    constructor(settings) {
        const given = settings || {};

        this.model = given.model || Model;
        this.titles = given.titles;
        this.markerAttribute = given.markerAttribute;
        this.headerRow = given.headerRow === undefined ? DEFAULT_HEADER_ROW : given.headerRow;
        this.firstDataRow = given.firstDataRow === undefined ? this.headeааrRow + 1 : given.firstDataRow;

        this._requireSettings();
    }

    /**
     * Refuse a meta that could not read a sheet, at the place it is written.
     *
     * @throws {Error} When titles or markerAttribute are missing, or the
     *     marker names no title.
     */
    _requireSettings() {
        if (!this.titles || Object.keys(this.titles).length === 0) {
            throw new Error("SheetMeta: не заданы titles");
        }

        if (!this.markerAttribute) {
            throw new Error("SheetMeta: не задан markerAttribute");
        }

        if (!(this.markerAttribute in this.titles)) {
            throw new Error(`SheetMeta: markerAttribute "${this.markerAttribute}" не найден среди titles`);
        }
    }

    toString() {
        return `SheetMeta(model=${this.model.name}, headerRow=${this.headerRow})`;
    }
}

/**
 * The records of one sheet, read and written as whole rows.
 *
 * A manager is handed a sheet and its SheetMeta, and does the reading and
 * writing itself:
 *
 *     const players = new Manager(sheet, PLAYERS_META);
 *
 *     players.getAll();                    // RecordSet of Player
 *     players.updateRecord(5, { hold: "да" });
 *
 * A sheet worked with in several places is better off as a subclass holding
 * its own meta, so that no caller has to carry it:
 *
 *     class Players extends Manager {
 *         constructor(sheet) {
 *             super(sheet, PLAYERS_META);
 *         }
 *     }
 *
 *     const players = new Players(sheet);
 *
 * Rows are counted from the sheet, not kept in memory: the last one is looked
 * up whenever it is asked for, so a manager does not go stale when the sheet
 * is written to behind its back. The header, on the other hand, is read once
 * and held, as SheetHeaders holds it.
 */
// eslint-disable-next-line no-unused-vars
class Manager {
    /**
     * @param {Sheet} sheet - Sheet holding the records.
     * @param {SheetMeta} meta - How the sheet is laid out.
     * @throws {Error} When the meta is missing.
     */
    constructor(sheet, meta) {
        if (!meta) {
            throw new Error(`${this.constructor.name}: не передана meta`);
        }

        this._sheet = sheet;
        this._meta = meta;
        this._headers = new SheetHeaders(sheet, meta.headerRow);
    }

    /********** Лист и его шапка **********/

    /**
     * @returns {Sheet}
     */
    get sheet() {
        return this._sheet;
    }

    /**
     * @returns {SheetHeaders}
     */
    get headers() {
        return this._headers;
    }

    /**
     * Attribute name to the header text its column sits under.
     *
     * @returns {Object<string, string>}
     */
    get titles() {
        return this._meta.titles;
    }

    /**
     * Column number per attribute, each looked up when it is read.
     *
     * @returns {Object<string, number>} Attribute name to column number.
     */
    get cols() {
        return this._headers.getLazyColumns(this.titles);
    }

    /**
     * First row a record can sit on: the one a subclass named, or the row
     * right under the header when it named none.
     *
     * @returns {number}
     */
    get firstDataRow() {
        return this._meta.firstDataRow;
    }

    /********** Границы данных **********/

    /**
     * Last row holding a record, found by the marker attribute.
     *
     * @returns {number} Zero when the sheet holds no records.
     */
    get lastDataRow() {
        const markerColumn = this.cols[this._meta.markerAttribute];
        const lastRow = this._sheet.getLastRow();
        const firstDataRow = this.firstDataRow;

        if (lastRow < firstDataRow) {
            return 0;
        }

        const rowCount = lastRow - firstDataRow + 1;
        const markers = this._sheet.getRange(firstDataRow, markerColumn, rowCount, 1).getValues();

        for (let i = markers.length - 1; i >= 0; i--) {
            const cellValue = String(markers[i][0]).trim();
            // if find first cell with value in column from bottom
            if (cellValue) {
                return firstDataRow + i;
            }
        }

        return 0;
    }

    /**
     * @returns {number} Zero when the sheet holds no records.
     */
    get count() {
        const lastDataRow = this.lastDataRow;

        return lastDataRow === 0 ? 0 : lastDataRow - this.firstDataRow + 1;
    }

    /********** Имена атрибутов **********/

    /**
     * Whether the model has an attribute under this name.
     *
     * @param {string} name - Attribute name.
     * @returns {boolean}
     */
    hasAttribute(name) {
        return name in this.titles;
    }

    /**
     * Names the model has no attribute for, in the order they were asked about.
     *
     * @param {string[]} names - Attribute names.
     * @returns {string[]} Empty when the model has every one of them.
     */
    findUnknownAttributes(names) {
        return names.filter((name) => !this.hasAttribute(name));
    }

    /**
     * Check that the model has an attribute for every name.
     *
     * @param {string[]} names - Attribute names.
     * @throws {Error} When at least one name is unknown, listing all of them.
     */
    requireAttributes(names) {
        const unknown = this.findUnknownAttributes(names);

        if (unknown.length > 0) {
            const unknownList = unknown.map((name) => `"${name}"`).join(", ");

            throw new Error(`${this.constructor.name}: неизвестные атрибуты: ${unknownList}`);
        }
    }

    /**
     * Every record of the sheet, in one read.
     *
     * Each one carries the row it was read from, so the set can write them
     * back:
     *
     *     const records = manager.getAll();
     *
     *     records.update({ payment: 100 });
     *
     * @returns {RecordSet} Empty set when the sheet holds no records.
     */
    getAll() {
        return new RecordSet(this, this._readRecords());
    }

    /**
     * The one record of the sheet whose attributes equal the given ones.
     *
     *     players.get({ playerId: "p-1" });
     *
     * The sheet is read whole, as getAll reads it, and the record is picked
     * out of what came back: one read, whatever the search asks for.
     *
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value it has to equal.
     * @returns {Model} The only record that matches.
     */
    get(valueByAttribute) {
        return this.getAll().get(valueByAttribute);
    }

    /**
     * The records of the sheet whose attributes equal the given ones.
     *
     *     players.filter({ campaignId: 7 }).update({ hold: "да" });
     *
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value it has to equal.
     * @returns {RecordSet} Empty set when none matches.
     */
    filter(valueByAttribute) {
        return this.getAll().filter(valueByAttribute);
    }

    /**
     * The records of the sheet that filter would leave out.
     *
     *     players.exclude({ hold: "" });  // те, у кого холд стоит
     *
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value that gets a record dropped.
     * @returns {RecordSet} Empty set when every record matches.
     */
    exclude(valueByAttribute) {
        return this.getAll().exclude(valueByAttribute);
    }

    /**
     * Remove the rows holding the given records.
     *
     * @param {RecordSet} records - Records to remove.
     * @returns {number} How many rows were removed.
     */
    delete(records) {
        // TODO: написать. Строки удаляются снизу вверх, иначе номера тех, что
        // ниже, съезжают на каждом удалении. Подряд идущие строки стоит
        // сносить одним deleteRows, чтобы не звать лист на каждую.
        throw new Error(`${this.constructor.name}.delete: не реализовано, records=${records}`);
    }

    /**
     * Every value of one attribute, top to bottom, in one read.
     *
     * @param {string} name - Attribute name.
     * @returns {*[]} Empty when the sheet holds no records.
     * @throws {Error} When the model has no such attribute.
     */
    getValues(name) {
        this.requireAttributes([name]);

        return this._getColValues(this.cols[name]);
    }

    /********** Запись **********/

    /**
     * Add records below the ones already there, all in one write.
     *
     * Values are keyed by header text, the way Model.toValues gives them:
     *
     *     manager.add([
     *         { "Partner ID": 42, "Player ID": "p-1" },
     *         { "Partner ID": 42, "Player ID": "p-2" },
     *     ]);
     *
     * @param {Object<string, *>[]} valueByTitleList - Values per record, in
     *     the order they are to be written.
     * @param {boolean} [skipStrangers=false] - Pass over a value whose header
     *     the sheet has no column for, instead of refusing the whole row.
     * @returns {Range|null} The range written, null when there was nothing to
     *     write.
     * @throws {Error} When the sheet has no column for one of the headers and
     *     skipStrangers is off.
     */
    add(valueByTitleList, skipStrangers = false) {
        const lastDataRow = this.lastDataRow;
        const firstRow = lastDataRow === 0 ? this.firstDataRow : lastDataRow + 1;

        return this._setRows(firstRow, valueByTitleList, skipStrangers);
    }

    /**
     * Change several attributes of one row, each one written on its own.
     *
     * Attributes are named the way the model names them, so a caller never
     * deals with column numbers:
     *
     *     manager.updateRecord(5, { hold: "да", comment: "по срезу" });
     *
     * Unlike add, the row is not rewritten whole: attributes left out keep
     * what they hold. The cells wanted are rarely next to each other, so they
     * go one setValue apiece rather than one setValues over a range.
     *
     * @param {number} row - Sheet row number of the record.
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value to put in its column.
     * @throws {Error} When a name is unknown, or the sheet has no column for
     *     the header it maps to.
     */
    updateRecord(row, valueByAttribute) {
        this.requireAttributes(Object.keys(valueByAttribute));

        const cols = this.cols;

        Object.entries(valueByAttribute).forEach(([name, value]) => {
            this._sheet.getRange(row, cols[name]).setValue(value);
        });
    }

    /**
     * Write whole columns at once, each from the first data row down.
     *
     * Every attribute named is written over the sheet in one call, so this is
     * how a run over all the records ends:
     *
     *     manager.update({
     *         playerCountry: countries,
     *         signUpDate: signUpDates,
     *     });
     *
     * Values are taken as they come, one per row from the first data row: a
     * list shorter than the sheet leaves the rows past its end alone, and a
     * longer one writes below the records.
     *
     * @param {Object<string, *[]>} valuesByAttribute - Attribute name and its
     *     values, top to bottom.
     * @throws {Error} When the model has no such attribute.
     */
    update(valuesByAttribute) {
        this.requireAttributes(Object.keys(valuesByAttribute));

        const cols = this.cols;

        Object.entries(valuesByAttribute).forEach(([name, values]) => {
            this._setColValues(cols[name], values);
        });
    }

    /********** Внутреннее **********/

    /**
     * Read every row holding a record and make a record of each.
     *
     * @returns {Model[]} Empty when the sheet holds no records.
     */
    _readRecords() {
        const rowCount = this.count;

        if (rowCount === 0) {
            return [];
        }

        const ModelClass = this._meta.model;
        const firstDataRow = this.firstDataRow;
        const range = this._sheet.getRange(firstDataRow, 1, rowCount, this._headers.lastColumn);

        return this._headers.fromRows(range.getValues()).map((valueByTitle, index) => {
            return new ModelClass(this, valueByTitle, firstDataRow + index);
        });
    }

    _getColValues(colNum) {
        const rowCount = this.count;

        if (rowCount === 0) {
            return [];
        }

        return this._sheet.getRange(this.firstDataRow, colNum, rowCount, 1).getValues().flat();
    }

    _setColValues(colNum, values) {
        if (values.length === 0) {
            return;
        }

        const rows = values.map((value) => [value]);
        this._sheet.getRange(this.firstDataRow, colNum, rows.length, 1).setValues(rows);
    }

    /**
     * Write records one after another, starting at the given row.
     *
     * Rows are overwritten whole rather than pushed down: a column the values
     * say nothing about is cleared, and whatever sat below firstRow is lost.
     *
     * @param {number} firstRow - Row to start at.
     * @param {Object<string, *>[]} valueByTitleList - Values per record.
     * @param {boolean} [skipStrangers=false] - Pass over a value whose header
     *     the sheet has no column for.
     * @returns {Range|null} The range written, null when there was nothing to
     *     write.
     */
    _setRows(firstRow, valueByTitleList, skipStrangers = false) {
        if (valueByTitleList.length === 0) {
            return null;
        }

        const rows = this._headers.toRows(valueByTitleList, skipStrangers);
        const range = this._sheet.getRange(firstRow, 1, rows.length, rows[0].length);

        range.setValues(rows);

        return range;
    }
}

/**
 * Records already read off a sheet, kept together so they can be written back
 * in one go.
 *
 * A set is what a manager gives for a question about its records, and it holds
 * the answer rather than the question: the sheet was read when the set was
 * made, and nothing here reads it again.
 *
 *     const held = players.filter((player) => isChecked_(player.hold));
 *
 *     held.count;                        // 12
 *     held.update({ holdMonth: "май" }); // writes those 12 rows
 *
 * Every set is its own: two of them made from the same manager hold two
 * arrays, and writing through one says nothing about the other.
 */
class RecordSet {
    /**
     * @param {Manager} manager - Manager the records were read by, and the
     *     one that writes them back.
     * @param {Model[]} records - Records the set holds, in the order they sit
     *     on the sheet.
     */
    constructor(manager, records) {
        this._manager = manager;
        this._records = records;
    }

    /**
     * Manager the records belong to.
     *
     * @returns {Manager}
     */
    get manager() {
        return this._manager;
    }

    /**
     * Records the set holds.
     *
     * @returns {Model[]}
     */
    get records() {
        return this._records;
    }

    /**
     * @returns {number}
     */
    get count() {
        return this._records.length;
    }

    /**
     * Whether the set holds anything, for asking without counting:
     *
     *     if (!players.getAll().filter({ hold: "да" }).exists()) { ... }
     *
     * @returns {boolean}
     */
    exists() {
        return this._records.length > 0;
    }

    /**
     * The records whose attributes are all equal to the given ones.
     *
     * Values are compared as they came off the sheet, without converting one
     * side to the other:
     *
     *     players.getAll().filter({ campaignId: 7 });
     *     players.getAll().filter({ campaignId: 7, hold: "" });
     *
     * A set comes back, not an array, so the search goes on where it left off:
     *
     *     players.getAll().filter({ campaignId: 7 }).update({ hold: "да" });
     *
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value it has to equal.
     * @returns {RecordSet} A new set over the same manager, empty when none
     *     matches.
     */
    filter(valueByAttribute) {
        return this._select(valueByAttribute, { keepMatching: true });
    }

    /**
     * The records that filter would leave out — the other side of the same
     * search.
     *
     *     players.getAll().exclude({ hold: "" });  // те, у кого холд стоит
     *
     * A record matching on every attribute given is dropped; one differing in
     * any of them is kept.
     *
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value that gets a record dropped.
     * @returns {RecordSet} A new set over the same manager, empty when every
     *     record matches.
     */
    exclude(valueByAttribute) {
        return this._select(valueByAttribute, { keepMatching: false });
    }

    /**
     * The one record whose attributes are all equal to the given ones.
     *
     *     players.getAll().get({ playerId: "p-1" });
     *     players.getAll().get({ partnerId: 42, campaignId: 7 });
     *
     * The search is the one filter does; what get adds is that it insists on
     * a single record. Finding none is an error, and so is finding several: a
     * search meant for one record that matches two says the sheet holds a
     * duplicate, which is worth hearing about where it happened.
     *
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value it has to equal.
     * @returns {Model} The only record that matches.
     * @throws {Error} When no record matches, or when several do.
     */
    get(valueByAttribute) {
        const found = this.filter(valueByAttribute).records;
        const asked = Object.entries(valueByAttribute)
            .map(([name, value]) => `${name}=${value}`)
            .join(", ");

        if (found.length === 0) {
            throw new Error(`${this.constructor.name}: запись не найдена: ${asked}`);
        }

        if (found.length > 1) {
            const rows = found.map((record) => record.row).join(", ");

            throw new Error(`${this.constructor.name}: найдено записей ${found.length}: ${asked}, строки: ${rows}`);
        }

        return found[0];
    }

    /**
     * The first record of the set, for when several matches are fine and the
     * one on top is wanted.
     *
     * @returns {Model|null} Null when the set is empty.
     */
    first() {
        // TODO: написать. Решить заодно, что значит "первый": сейчас записи
        // лежат в порядке строк листа, но после orderBy это будет не так.
        throw new Error(`${this.constructor.name}.first: не реализовано`);
    }

    /**
     * The set sorted by one attribute.
     *
     * @param {string} name - Attribute name to sort by.
     * @returns {RecordSet} A new set over the same manager.
     */
    orderBy(name) {
        // TODO: написать. Продумать сравнение: на листе в одной колонке
        // попадаются и числа, и строки, и даты, и пустые ячейки.
        throw new Error(`${this.constructor.name}.orderBy: не реализовано, name=${name}`);
    }

    /**
     * Remove the rows the set holds.
     *
     * @returns {number} How many rows were removed.
     */
    delete() {
        // TODO: написать через Manager.delete. Решить, что делать с самим
        // набором после удаления: записи в нём останутся с номерами строк,
        // которых на листе уже нет.
        throw new Error(`${this.constructor.name}.delete: не реализовано`);
    }

    /**
     * Give every record of the set the same attributes.
     *
     * The records are changed too, not only the sheet, so a set read before
     * the write still says what the sheet now holds.
     *
     * Each record is written on its own, which is one call to the sheet per
     * record per attribute. That is fine for a handful of rows and slow for a
     * few hundred; a whole column at once is what Manager.setValues is for.
     *
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value to put in its column.
     * @returns {number} How many records were written.
     */
    update(valueByAttribute) {
        this._records.forEach((record) => {
            this._manager.updateRecord(record.row, valueByAttribute);
            this._setAttributes(record, valueByAttribute);
        });

        return this._records.length;
    }

    /**
     * The records a search keeps, or the ones it leaves out.
     *
     * @param {Object<string, *>} valueByAttribute - Attribute name and the
     *     value it is compared to.
     * @param {{keepMatching: boolean}} options - True keeps the records that
     *     match every attribute, false keeps the rest.
     * @returns {RecordSet} A new set over the same manager.
     * @throws {Error} When no attribute is given, which would ask nothing of
     *     the records and quietly take either all of them or none.
     */
    _select(valueByAttribute, { keepMatching }) {
        const wanted = Object.entries(valueByAttribute);

        if (wanted.length === 0) {
            throw new Error(`${this.constructor.name}: не заданы атрибуты для поиска`);
        }

        this._manager.requireAttributes(Object.keys(valueByAttribute));

        const found = this._records.filter((record) => {
            const isMatch = wanted.every(([name, value]) => record[name] === value);

            return isMatch === keepMatching;
        });

        return new RecordSet(this._manager, found);
    }

    /**
     * Put the values on the record itself, so it says what its row now holds.
     *
     * @param {Model} record - Record to change.
     * @param {Object<string, *>} valueByAttribute - Attribute name and value.
     */
    _setAttributes(record, valueByAttribute) {
        Object.entries(valueByAttribute).forEach(([name, value]) => {
            record[name] = value;
        });
    }

    toString() {
        return `${this.constructor.name}(count=${this._records.length})`;
    }
}
