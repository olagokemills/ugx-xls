"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SpreadsheetError = void 0;
class SpreadsheetError extends Error {
    constructor(message, code) {
        super(message);
        this.name = "SpreadsheetError";
        this.code = code;
        Object.setPrototypeOf(this, SpreadsheetError.prototype);
    }
}
exports.SpreadsheetError = SpreadsheetError;
