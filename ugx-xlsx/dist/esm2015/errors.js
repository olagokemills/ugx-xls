export class SpreadsheetError extends Error {
    constructor(message, code) {
        super(message);
        this.name = "SpreadsheetError";
        this.code = code;
        Object.setPrototypeOf(this, SpreadsheetError.prototype);
    }
}
