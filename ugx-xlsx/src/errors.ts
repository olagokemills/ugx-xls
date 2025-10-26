export class SpreadsheetError extends Error {
  code?: string;
  constructor(message: string, code?: string) {
    super(message);
    this.name = "SpreadsheetError";
    this.code = code;
    Object.setPrototypeOf(this, SpreadsheetError.prototype);
  }
}
