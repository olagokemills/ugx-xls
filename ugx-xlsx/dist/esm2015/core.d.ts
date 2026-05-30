/** A single data row: any plain object whose values will be serialised to strings. */
export type WorksheetRow = Record<string, any>;
/** Options for {@link jsonToWorkbookBuffer} and {@link exportJsonToExcel}. */
export interface WriteOptions {
    /** Name of the worksheet tab. @default "Sheet1" */
    sheetName?: string;
    /** Strip ASCII control chars from header keys before writing. @default true */
    sanitizeHeaders?: boolean;
    /** Maximum number of data rows allowed (guard against runaway exports). @default 100_000 */
    maxRows?: number;
    /** Maximum number of columns (keys per row). @default 1_000 */
    maxCols?: number;
    /** Maximum characters per cell; longer values are hard-truncated. @default 10_000 */
    maxCellChars?: number;
}
/** Merged defaults applied when individual {@link WriteOptions} fields are omitted. */
export declare const DEFAULT_WRITE_OPTIONS: Required<WriteOptions>;
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
export declare function jsonToWorkbookBuffer(rows: WorksheetRow[], opts?: WriteOptions): Promise<ArrayBuffer | Uint8Array>;
