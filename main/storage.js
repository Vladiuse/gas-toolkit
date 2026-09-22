const STORAGE_SHEET = "_STORAGE";
const LOCK_TIMEOUT_MS = 30000;
const STORAGE_SHEET_HEADERS = ["Key", "Value"];
const STORAGE_HEADER_COLOR = "#D9EDF7";

const STORAGE_KEY_COLUMN = 1;
const STORAGE_VALUE_COLUMN = 2;
const STORAGE_FIRST_DATA_ROW = 2;

/**
 * Key value storage kept in a spreadsheet sheet, in the spirit of localStorage.
 *
 * Keys and values are plain strings, exactly like localStorage.
 * The sheet is read once per instance and then served from memory, so reuse the
 * same instance instead of creating one per call.
 */
// eslint-disable-next-line no-unused-vars
class SheetStorage {
    /**
     * @param {string} [sheetName] - Sheet holding the keys and values.
     */
    constructor(sheetName = STORAGE_SHEET) {
        this._sheetName = sheetName;
    }

    /**
     * Read the value stored under the key.
     *
     * @param {string} key
     * @param {string} [fallback=null] - Returned when the key is missing.
     * @returns {string} The stored value, or the fallback.
     * @throws {Error} When the key is missing, the way localStorage does.
     */
    get(key, fallback = null) {
        this._requireKey(arguments.length, "get");

        const row = this._loadKeyRows().get(this._normalizeKey(key));

        if (row === undefined) {
            return fallback;
        }
        const value = this._getSheet().getRange(row, STORAGE_VALUE_COLUMN).getValue()
        return String(value);
    }

    /**
     * Store the value under the key, replacing whatever was there.
     *
     * The sheet is read and written under a script lock, so two triggers
     * running at the same time cannot both append the same key.
     *
     * @param {string} key
     * @param {string} value
     * @throws {Error} When the value is missing or the lock cannot be taken.
     */
    set(key, value) {
        this._requireKey(arguments.length, "set");

        if (arguments.length < 2) {
            throw new Error("SheetStorage.set: нужно передать значение");
        }

        const storedKey = this._normalizeKey(key);
        const storedValue = String(value);

        const lock = LockService.getScriptLock();
        if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
            throw new Error("SheetStorage.set: не удалось получить блокировку");
        }

        try {
            const sheet = this._getSheet();
            const row = this._loadKeyRows().get(storedKey);

            if (row !== undefined) {
                sheet.getRange(row, STORAGE_VALUE_COLUMN).setValue(storedValue);
            } else {
                sheet.appendRow([storedKey, storedValue]);
            }

            SpreadsheetApp.flush();
        } finally {
            lock.releaseLock();
        }
    }

    /**
     * Tell whether the key is present, even when its value is empty.
     *
     * @param {string} key
     * @returns {boolean}
     * @throws {Error} When the key is missing, the way localStorage does.
     */
    has(key) {
        this._requireKey(arguments.length, "has");

        return this._loadKeyRows().has(this._normalizeKey(key));
    }

    /**
     * Delete the key and its row, the way localStorage.removeItem does.
     *
     * The sheet is read and the row deleted under a script lock: the row
     * number comes from a read, and a stale one would delete someone else.
     *
     * @param {string} key
     * @returns {boolean} True when a row was deleted, false when the key
     *     was not there.
     * @throws {Error} When the key is missing or the lock cannot be taken.
     */
    remove(key) {
        this._requireKey(arguments.length, "remove");

        const storedKey = this._normalizeKey(key);

        const lock = LockService.getScriptLock();
        if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
            throw new Error("SheetStorage.remove: не удалось получить блокировку");
        }

        try {
            const row = this._loadKeyRows().get(storedKey);
            if (row === undefined) {
                return false;
            }

            this._getSheet().deleteRow(row);
            SpreadsheetApp.flush();
            return true;
        } finally {
            lock.releaseLock();
        }
    }

    /**
     * Drop every stored key, the way localStorage.clear does.
     * The header row stays, so the sheet keeps its look.
     *
     * @returns {number} How many rows were deleted.
     * @throws {Error} When the lock cannot be taken.
     */
    clear() {
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
            throw new Error("SheetStorage.clear: не удалось получить блокировку");
        }

        try {
            const rowCount = this.rowCount;
            if (rowCount === 0) {
                return 0;
            }

            this._getSheet().deleteRows(STORAGE_FIRST_DATA_ROW, rowCount);
            SpreadsheetApp.flush();

            return rowCount;
        } finally {
            lock.releaseLock();
        }
    }

    /**
     * How many keys the sheet holds, not counting the header row.
     *
     * @returns {number}
     */
    get rowCount() {
        const lastRow = this._getSheet().getLastRow();

        if (lastRow < STORAGE_FIRST_DATA_ROW) {
            return 0;
        }

        return lastRow - STORAGE_FIRST_DATA_ROW + 1;
    }

    /**
     * Refuse a call made without a key, the way localStorage does.
     *
     * TODO: move to helpers.js, any class checking its arguments needs this.
     *
     * @param {number} argumentCount - How many arguments the caller passed.
     * @param {string} methodName - Named in the error message.
     */
    _requireKey(argumentCount, methodName) {
        if (argumentCount < 1) {
            throw new Error("SheetStorage." + methodName + ": нужно передать ключ");
        }
    }

    /**
     * Bring a key to the form it is stored and looked up by.
     *
     * Unlike localStorage, surrounding spaces are dropped: this sheet is also
     * edited by hand, where a trailing space is invisible and would silently
     * turn one key into two.
     *
     * @param {*} key - Anything printable, converted to a trimmed string.
     * @returns {string}
     */
    _normalizeKey(key) {
        return String(key).trim();
    }

    /**
     * Read the key column and map every key to its row.
     *
     * @returns {Map<string, number>}
     */
    _loadKeyRows() {
        const keyRows = new Map();

        const rowCount = this.rowCount;
        if (rowCount === 0) {
            return keyRows;
        }

        const sheet = this._getSheet();
        const keyColumn = sheet.getRange(STORAGE_FIRST_DATA_ROW, STORAGE_KEY_COLUMN, rowCount, 1).getValues();

        for (let i = 0; i < keyColumn.length; i++) {
            const key = this._normalizeKey(keyColumn[i][0]);
            const row = STORAGE_FIRST_DATA_ROW + i;

            if (key) {
                keyRows.set(key, row);
            }
        }

        return keyRows;
    }

    /**
     * Return the storage sheet, creating it with headers on first use.
     */
    _getSheet() {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = spreadsheet.getSheetByName(this._sheetName);

        if (!sheet) {
            sheet = spreadsheet.insertSheet(this._sheetName);
            sheet.appendRow(STORAGE_SHEET_HEADERS);
            sheet
                .getRange(1, 1, 1, STORAGE_SHEET_HEADERS.length)
                .setBackground(STORAGE_HEADER_COLOR)
                .setFontWeight("bold");
        }

        return sheet;
    }
}
