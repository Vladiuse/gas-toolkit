/* global isString, isNumber, isFunction, isSheet */

const DEFAULT_FIRST_DATA_ROW = 2;

/**
 * Tell whether a cell value stands for a ticked checkbox.
 */
function isCheckboxTicked(value) {
    return value === true || String(value).toUpperCase() === "TRUE";
}

/**
 * Ways the edited value may have changed, one per valueChange filter option.
 */
const VALUE_CHANGE_CHECKS = {
    any: (context) => context.value !== context.oldValue,
    filled: (context) => Boolean(context.value) && !context.oldValue,
    cleared: (context) => !context.value && Boolean(context.oldValue),
    checked: (context) => isCheckboxTicked(context.value),
    unchecked: (context) => !isCheckboxTicked(context.value) && isCheckboxTicked(context.oldValue),
};
// TODO
// editedByUser — skip edits made by the script itself. Our own code writes into cells (appendComment_, stamping dates), and that can fire the trigger again.
// ignoreSheets — a blacklist of sheets. Handy for service ones: _LOG, SlackLog, MonthlyAlertLog — edits there must never start anything.
// singleCellOnly — ignore a range paste. Pasting 50 rows currently gives the handler a single event holding a range, while the code expects one cell (e.range.getRow() returns the first).

/**
 * Routes a single onEdit event to every registered handler that matches it.
 *
 * Each handler runs in its own try/catch, so a broken one neither stops the
 * others nor surfaces as a failed trigger run.
 */
class EditRouter {
    /**
     * @param {SheetLogger} logger - Used to report handler failures.
     */
    constructor(logger) {
        this.logger = logger;
        this.handlers = [];
    }

    /**
     * Tell whether the handler set this filter at all.
     */
    _isFilterSet(value) {
        return value !== undefined && value !== null;
    }

    /**
     * Tell whether the edit landed above the data, on a header row or the like.
     */
    _isAboveData(context, firstDataRow) {
        if (isNumber(firstDataRow)) {
            return context.row < firstDataRow;
        }

        throw new Error("The firstDataRow filter must be a row number");
    }

    /**
     * Match the edited sheet against the filter, which is either its name,
     * the sheet itself, or a predicate taking the name.
     */
    _matchesSheet(filter, context) {
        if (!this._isFilterSet(filter)) {
            return true;
        }

        if (isString(filter)) {
            return filter === context.sheetName;
        }

        if (isSheet(filter)) {
            return filter.getSheetId() === context.sheet.getSheetId();
        }

        if (isFunction(filter)) {
            return Boolean(filter(context.sheetName));
        }

        throw new Error("The sheet filter must be a sheet name, a sheet or a function");
    }

    /**
     * Match the edited column against the filter, which is either the column
     * number or a function taking the sheet and returning that number.
     */
    _matchesColumn(filter, context) {
        if (!this._isFilterSet(filter)) {
            return true;
        }

        if (isFunction(filter)) {
            return filter(context.sheet) === context.column;
        }

        if (isNumber(filter)) {
            return filter === context.column;
        }

        throw new Error("The column filter must be a column number or a function");
    }

    /**
     * Match how the edited value changed against the filter.
     * The filter itself is validated in register().
     */
    _matchesValueChange(filter, context) {
        if (!this._isFilterSet(filter)) {
            return true;
        }

        return VALUE_CHANGE_CHECKS[filter](context);
    }

    /**
     * Run the check written by the handler itself, for conditions the other
     * filters cannot express. A handler without such a check passes.
     */
    _matchesCustomFilter(filter, context) {
        if (!this._isFilterSet(filter)) {
            return true;
        }

        if (isFunction(filter)) {
            return Boolean(filter(context));
        }

        throw new Error("The customFilter filter must be a function");
    }

    /**
     * Check the context against every filter of the handler.
     */
    _matches(handler, context) {
        if (!this._matchesSheet(handler.sheet, context)) {
            return false;
        }

        if (this._isAboveData(context, handler.firstDataRow)) {
            return false;
        }

        if (!this._matchesColumn(handler.column, context)) {
            return false;
        }

        if (!this._matchesValueChange(handler.valueChange, context)) {
            return false;
        }

        if (!this._matchesCustomFilter(handler.customFilter, context)) {
            return false;
        }

        return true;
    }

    /**
     * Collect what every handler needs, so each one does not read it again.
     */
    _buildContext(event) {
        const range = event.range;
        const sheet = range.getSheet();

        return {
            event: event,
            range: range,
            sheet: sheet,
            sheetName: sheet.getName(),
            row: range.getRow(),
            column: range.getColumn(),
            value: event.value,
            oldValue: event.oldValue,
        };
    }

    /**
     * Register one handler with the filters that decide when it runs.
     * Name and run are required, every filter is optional and an omitted one
     * accepts anything.
     *
     * @param {Object} handler
     * @param {string} handler.name - Shown in the log when the handler fails.
     * @param {Function} handler.run - Called as run(context).
     * @param {string|Sheet|Function} [handler.sheet] - The sheet to watch:
     *     its name, the sheet itself, or a predicate taking the name and
     *     returning a boolean.
     * @param {number} [handler.firstDataRow=2] - Number of the first row that
     *     holds data. Edits above it are ignored, so the default skips a
     *     single header row.
     * @param {number|Function} [handler.column] - Column number, or a function
     *     taking the sheet and returning the column number to watch.
     * @param {string} [handler.valueChange] - How the value must have changed:
     *     any, filled, cleared, checked or unchecked.
     * @param {Function} [handler.customFilter] - Check written by the handler
     *     itself, taking the context and returning a boolean. For conditions
     *     the fields above cannot express.
     * @returns {EditRouter} The router itself, so calls can be chained.
     * @throws {Error} When a required field is missing or a value is unknown.
     */
    register(handler) {
        const {
            name,
            run,
            sheet = null,
            firstDataRow = DEFAULT_FIRST_DATA_ROW,
            column = null,
            valueChange = null,
            customFilter = null,
        } = handler;

        if (!name) {
            throw new Error("A handler must have a name");
        }

        if (!isFunction(run)) {
            throw new Error("Handler \"" + name + "\": the run field must be a function");
        }

        if (valueChange !== null && !VALUE_CHANGE_CHECKS[valueChange]) {
            throw new Error(
                "Handler \"" + name + "\": the valueChange field must be one of: " + Object.keys(VALUE_CHANGE_CHECKS).join(", "),
            );
        }

        this.handlers.push({ name, run, sheet, firstDataRow, column, valueChange, customFilter });
        return this;
    }

    /**
     * Run every matching handler for the given onEdit event.
     *
     * @param {Object} event - The event object passed by the trigger.
     */
    handle(event) {
        if (!event || !event.range) {
            return;
        }

        const context = this._buildContext(event);

        for (const handler of this.handlers) {
            try {
                if (this._matches(handler, context)) {
                    handler.run(context);
                }
            } catch (error) {
                this.logger.error("Handler \"" + handler.name + "\": " + error);
            }
        }
    }
}
