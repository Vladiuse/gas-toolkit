/* global truncate */

const LEVEL_INFO = "INFO";
const LEVEL_WARNING = "WARNING";
const LEVEL_ERROR = "ERROR";

const LOG_SHEET = "_LOG";
const LOG_SHEET_HEADERS = ["Date", "Level", "Message"];

const HEADER_COLOR = "#D9EDF7";

const LEVEL_COLORS = {
    WARNING: "#FFF3CD",
    WARN: "#FFF3CD",
    ERROR: "#F8D7DA",
};

const TOAST_TEXT_LIMIT = 100;

const CONSOLE_METHODS = {
    INFO: console.info,
    WARNING: console.warn,
    ERROR: console.error,
};

const LEVEL_ALIASES = {
    INFO: LEVEL_INFO,
    WARN: LEVEL_WARNING,
    WARNING: LEVEL_WARNING,
    ERROR: LEVEL_ERROR,
};

/**
 * Bring a level written in any case to its canonical name.
 * Unknown values fall back to INFO.
 */
function normalizeLevel(level) {
    const name = String(level).trim().toUpperCase();
    return LEVEL_ALIASES[name] || LEVEL_INFO;
}

/**
 * Append one row to the log sheet, creating the sheet on first use.
 * Rows are colored by level, INFO is left plain.
 *
 * @param {*} message - Anything printable, converted to a string.
 * @param {string} [level=INFO] - Any case, WARN is accepted for WARNING;
 *     an unknown value falls back to INFO.
 */
function logToSheet(message, level = LEVEL_INFO) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    let sheet = spreadsheet.getSheetByName(LOG_SHEET);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(LOG_SHEET);
        sheet.appendRow(LOG_SHEET_HEADERS);
        sheet.getRange(1, 1, 1, LOG_SHEET_HEADERS.length).setBackground(HEADER_COLOR).setFontWeight("bold");
    }

    const normalizedLevel = normalizeLevel(level);
    sheet.appendRow([new Date(), normalizedLevel, String(message)]);

    const color = LEVEL_COLORS[normalizedLevel];
    if (color) {
        sheet.getRange(sheet.getLastRow(), 1, 1, LOG_SHEET_HEADERS.length).setBackground(color);
    }
}

// eslint-disable-next-line no-unused-vars
class SheetLogger {
    _showToast(level, text) {
        try {
            SpreadsheetApp.getActive().toast(truncate(text, TOAST_TEXT_LIMIT), level);
        } catch (error) {
            console.error("Failed to show the toast: " + error);
        }
    }

    /**
     * Write the message to the execution log and to the log sheet.
     * Never throws: logging must not break the calling code.
     *
     * @param {string} level - INFO, WARNING or ERROR.
     * @param {*} message - Anything printable, converted to a string.
     * @param {Object} [options] - Extra output channels.
     * @param {boolean} [options.toast=false] - Show the message as a toast
     *     in the spreadsheet on top of the usual logging.
     */
    _log(level, message, options = {}) {
        const text = String(message);
        const writeToConsole = CONSOLE_METHODS[level] || console.info;
        writeToConsole(text);

        try {
            logToSheet(text, level);
        } catch (error) {
            console.error("Failed to write to the log sheet: " + error);
        }

        if (options.toast) {
            this._showToast(level, text);
        }
    }

    info(message, options) {
        this._log(LEVEL_INFO, message, options);
    }

    warning(message, options) {
        this._log(LEVEL_WARNING, message, options);
    }
    
    warn(message, options) {
        this.warning(message, options);
    }

    error(message, options) {
        this._log(LEVEL_ERROR, message, options);
    }
}
