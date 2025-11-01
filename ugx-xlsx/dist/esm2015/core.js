var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import Excel from "exceljs";
import { sanitizeHeader, enforceLimits } from "./sanitizer";
import { SpreadsheetError } from "./errors";
export const DEFAULT_WRITE_OPTIONS = {
    sheetName: "Sheet1",
    sanitizeHeaders: true,
    maxRows: 100000,
    maxCols: 1000,
    maxCellChars: 10000,
};
/** JSON → XLSX buffer (Node/Browser compatible buffer) */
export function jsonToWorkbookBuffer(rows, opts) {
    return __awaiter(this, void 0, void 0, function* () {
        const options = Object.assign(Object.assign({}, DEFAULT_WRITE_OPTIONS), (opts || {}));
        if (!Array.isArray(rows)) {
            throw new SpreadsheetError("rows must be an array", "INVALID_INPUT");
        }
        if (rows.length > options.maxRows) {
            throw new SpreadsheetError("Row limit exceeded", "LIMIT_ROWS");
        }
        const wb = new Excel.Workbook();
        const ws = wb.addWorksheet(options.sheetName);
        if (rows.length === 0) {
            ws.addRow([]);
            return wb.xlsx.writeBuffer();
        }
        const headers = Object.keys(rows[0])
            .slice(0, options.maxCols)
            .map((h) => (options.sanitizeHeaders ? sanitizeHeader(h) : h));
        if (headers.length > options.maxCols) {
            throw new SpreadsheetError("Column limit exceeded", "LIMIT_COLS");
        }
        ws.addRow(headers);
        for (let i = 0; i < rows.length; i++) {
            if (i + 1 > options.maxRows) {
                throw new SpreadsheetError("Row limit exceeded during write", "LIMIT_ROWS");
            }
            const r = rows[i];
            const values = [];
            for (let c = 0; c < headers.length; c++) {
                if (c + 1 > options.maxCols) {
                    throw new SpreadsheetError("Column limit exceeded during write", "LIMIT_COLS");
                }
                const key = headers[c];
                let v = r[key];
                if (v === undefined || v === null) {
                    values.push("");
                    continue;
                }
                if (typeof v === "object") {
                    try {
                        v = JSON.stringify(v);
                    }
                    catch (_a) {
                        v = String(v);
                    }
                }
                else {
                    v = String(v);
                }
                values.push(enforceLimits(v, options.maxCellChars));
            }
            ws.addRow(values);
        }
        return wb.xlsx.writeBuffer();
    });
}
