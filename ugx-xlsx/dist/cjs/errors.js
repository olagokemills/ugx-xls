"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SpreadsheetError = void 0;
/**
 * Structured error thrown by ugx-xlsx operations.
 *
 * The `code` property lets callers distinguish error categories without
 * parsing the human-readable message string.
 *
 * @example
 * try {
 *   await jsonToWorkbookBuffer(rows);
 * } catch (err) {
 *   if (err instanceof SpreadsheetError && err.code === 'LIMIT_ROWS') {
 *     console.error('Too many rows — paginate your export.');
 *   }
 * }
 */
class SpreadsheetError extends Error {
    /**
     * @param message - Human-readable description of the error.
     * @param code    - Machine-readable category: `INVALID_INPUT`, `LIMIT_ROWS`, or `LIMIT_COLS`.
     */
    constructor(message, code) {
        super(message);
        this.name = "SpreadsheetError";
        this.code = code;
        Object.setPrototypeOf(this, SpreadsheetError.prototype);
    }
}
exports.SpreadsheetError = SpreadsheetError;
