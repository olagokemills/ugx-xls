import { sanitizeHeader, enforceLimits } from "./sanitizer";
import { SpreadsheetError } from "./errors";
import writeXlsxFile from "write-excel-file";
/** Merged defaults applied when individual {@link WriteOptions} fields are omitted. */
export const DEFAULT_WRITE_OPTIONS = {
    sheetName: "Sheet1",
    sanitizeHeaders: true,
    maxRows: 100000,
    maxCols: 1000,
    maxCellChars: 10000,
};
/** Returns `true` when running inside Node.js (not a browser). */
function isNodeRuntime() {
    const g = globalThis;
    return !!g.process?.versions?.node;
}
/**
 * Dynamically loads `write-excel-file/node` via CJS `require`.
 * Returns `undefined` when `require` is not available (browser environments).
 */
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
/**
 * Converts an array of plain objects into a raw XLSX workbook buffer.
 *
 * Works in both **Node.js** (returns `Uint8Array`) and **browsers**
 * (returns `ArrayBuffer`). Object keys become column headers; values are
 * coerced to strings. `null` / `undefined` cells are written as empty strings.
 * Objects are JSON-serialised; all other primitives use `String()`.
 *
 * @param rows - Array of plain objects to serialise. Column order is derived
 *               from the keys of the first row.
 * @param opts - Optional overrides for sheet name, safety limits, and sanitisation.
 * @returns A raw XLSX buffer suitable for saving to disk or streaming to a client.
 *
 * @throws {SpreadsheetError} `INVALID_INPUT` — `rows` is not an array.
 * @throws {SpreadsheetError} `LIMIT_ROWS`    — row count exceeds `maxRows`.
 * @throws {SpreadsheetError} `LIMIT_COLS`    — column count exceeds `maxCols`.
 *
 * @example
 * // Node.js — write buffer to disk
 * const buf = await jsonToWorkbookBuffer(
 *   [{ name: 'Alice', score: 42 }, { name: 'Bob', score: 98 }],
 *   { sheetName: 'Results' }
 * );
 * fs.writeFileSync('results.xlsx', Buffer.from(buf));
 *
 * @example
 * // Browser — wrap in a Blob for download
 * const buf = await jsonToWorkbookBuffer(rows);
 * const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
 * saveAs(blob, 'export.xlsx');
 */
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
