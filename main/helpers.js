/**
 * Small helpers shared across the project.
 */

const DEFAULT_MODAL_TITLE = "Сообщение";
const DEFAULT_CONFIRM_TITLE = "Подтверждение";

// eslint-disable-next-line no-unused-vars
function isString(value) {
    return typeof value === "string";
}

// eslint-disable-next-line no-unused-vars
function isNumber(value) {
    return typeof value === "number";
}

// eslint-disable-next-line no-unused-vars
function isFunction(value) {
    return typeof value === "function";
}

/**
 * Tell a Sheet apart from a plain value by the method only it has.
 * Apps Script exposes no Sheet constructor, so instanceof is not an option.
 */
// eslint-disable-next-line no-unused-vars
function isSheet(value) {
    return typeof value === "object" && value !== null && typeof value.getSheetId === "function";
}

/**
 * Tell a usable Date apart from a plain value, an invalid date included.
 */
// eslint-disable-next-line no-unused-vars
function isDate(value) {
    return value instanceof Date && !isNaN(value.getTime());
}

/**
 * Cut the text to the given length, marking the cut with an ellipsis.
 */
// eslint-disable-next-line no-unused-vars
function truncate(text, limit) {
    return text.length > limit ? text.slice(0, limit - 1) + "…" : text;
}

/**
 * Show a short message in the corner of the active spreadsheet.
 *
 * Called with one argument it is the message itself; with two the title comes
 * first, the way a heading is written above what it heads:
 *
 *     toast("Готово");
 *     toast("Перенос холдов", "Август -> Сентябрь");
 *
 * @param {string} titleOrText - Title when a text follows, the text itself
 *     otherwise.
 * @param {string} [text] - Message body.
 */
// eslint-disable-next-line no-unused-vars
function toast(titleOrText, text) {
    if (text === undefined) {
        SpreadsheetApp.getActive().toast(titleOrText);
        return;
    }

    SpreadsheetApp.getActive().toast(text, titleOrText);
}

/**
 * Show a message the user has to close.
 *
 * @param {string} text - Message body.
 * @param {string} [title="Сообщение"] - Line above the text.
 */
function modal(text, title = DEFAULT_MODAL_TITLE) {
    const ui = SpreadsheetApp.getUi();

    ui.alert(title, text, ui.ButtonSet.OK);
}

// eslint-disable-next-line no-unused-vars
function modalError(text){
    return modal(text, "Ошибка")
}

/**
 * Ask the user a yes-or-no question and wait for the answer.
 *
 * Closing the dialog with the cross counts as a no, so the answer is true only
 * when "Да" was actually pressed:
 *
 *     if (!modalConfirm("Удалить выбранные строки?")) {
 *         return;
 *     }
 *
 * @param {string} text - Question body.
 * @param {string} [title="Подтверждение"] - Line above the text.
 * @returns {boolean} Whether the user agreed.
 */
// eslint-disable-next-line no-unused-vars
function modalConfirm(text, title = DEFAULT_CONFIRM_TITLE) {
    const ui = SpreadsheetApp.getUi();
    const answer = ui.alert(title, text, ui.ButtonSet.YES_NO);

    return answer === ui.Button.YES;
}