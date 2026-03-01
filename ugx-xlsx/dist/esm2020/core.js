import { sanitizeHeader, enforceLimits } from "./sanitizer";
import { SpreadsheetError } from "./errors";
import writeXlsxFile from "write-excel-file";
export const DEFAULT_WRITE_OPTIONS = {
    sheetName: "Sheet1",
    sanitizeHeaders: true,
    maxRows: 100000,
    maxCols: 1000,
    maxCellChars: 10000,
};
function isNodeRuntime() {
    const g = globalThis;
    return !!g.process?.versions?.node;
}
function getNodeWriter() {
    const g = globalThis;
    if (typeof g.require === "function") {
        const loaded = g.require("write-excel-file/node");
        if (typeof loaded === "function")
            return loaded;
        return loaded.default;
    }
    return undefined;
}
/** JSON → XLSX buffer (Node/Browser compatible buffer) */
export async function jsonToWorkbookBuffer(rows, opts) {
    const options = { ...DEFAULT_WRITE_OPTIONS, ...(opts || {}) };
    if (!Array.isArray(rows)) {
        throw new SpreadsheetError("rows must be an array", "INVALID_INPUT");
    }
    if (rows.length > options.maxRows) {
        throw new SpreadsheetError("Row limit exceeded", "LIMIT_ROWS");
    }
    const headers = (rows[0] ? Object.keys(rows[0]) : [])
        .slice(0, options.maxCols)
        .map((h) => (options.sanitizeHeaders ? sanitizeHeader(h) : h));
    if (headers.length > options.maxCols) {
        throw new SpreadsheetError("Column limit exceeded", "LIMIT_COLS");
    }
    const data = [];
    if (rows.length === 0) {
        data.push([]);
    }
    else {
        data.push(headers.map((h) => ({ value: h })));
    }
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
                values.push({ value: "" });
                continue;
            }
            if (typeof v === "object") {
                try {
                    v = JSON.stringify(v);
                }
                catch {
                    v = String(v);
                }
            }
            else {
                v = String(v);
            }
            values.push({ value: enforceLimits(v, options.maxCellChars) });
        }
        data.push(values);
    }
    if (isNodeRuntime()) {
        const nodeWriter = getNodeWriter();
        if (nodeWriter) {
            return nodeWriter(data, {
                buffer: true,
                sheet: options.sheetName,
            });
        }
    }
    const out = await writeXlsxFile(data, { sheet: options.sheetName });
    return out.arrayBuffer();
}
