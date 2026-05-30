/** A single CSV row: any plain object whose values will be serialised to strings. */
export type CsvRow = Record<string, any>;
/** Options for {@link jsonToCsvString} and {@link exportJsonToCsv}. */
export interface CsvOptions {
    /** Strip ASCII control chars from header keys before writing. @default true */
    sanitizeHeaders?: boolean;
    /** Maximum number of data rows allowed (guard against runaway exports). @default 100_000 */
    maxRows?: number;
    /** Maximum number of columns. @default 1_000 */
    maxCols?: number;
    /** Maximum characters per cell; longer values are hard-truncated. @default 10_000 */
    maxCellChars?: number;
    /**
     * Explicit column order. When provided, only these keys are written and in
     * this order. Omitting keys from a row writes an empty field for that column.
     * Defaults to the key order of the first row.
     */
    headers?: string[];
    /**
     * Prepend a UTF-8 BOM (`﻿`) to the output.
     * Required for Excel on Windows to open the file with correct encoding.
     * @default false
     */
    includeBom?: boolean;
}
/** Merged defaults applied when individual {@link CsvOptions} fields are omitted. */
export declare const DEFAULT_CSV_OPTIONS: Required<Omit<CsvOptions, 'headers'>>;
/**
 * Converts an array of plain objects to an RFC 4180-compliant CSV string.
 *
 * Column headers are derived from the keys of the first row (or from
 * `opts.headers` when provided). Values are coerced to strings; objects are
 * JSON-serialised; `null` / `undefined` cells become empty fields.
 *
 * @param rows - Array of plain objects to serialise.
 * @param opts - Optional overrides for headers, limits, BOM, and sanitisation.
 * @returns RFC 4180 CSV string with CRLF line endings. Returns an empty string
 *          (or a bare BOM if `includeBom` is true) when `rows` is empty.
 *
 * @throws {SpreadsheetError} `INVALID_INPUT` — `rows` is not an array.
 * @throws {SpreadsheetError} `LIMIT_COLS`    — column count exceeds `maxCols`.
 * @throws {SpreadsheetError} `LIMIT_ROWS`    — row count exceeds `maxRows`.
 *
 * @example
 * const csv = jsonToCsvString(
 *   [{ name: 'Alice', score: 42 }, { name: 'Bob', score: 98 }],
 *   { includeBom: true }
 * );
 * fs.writeFileSync('results.csv', csv, 'utf8');
 */
export declare function jsonToCsvString(rows: CsvRow[], opts?: CsvOptions): string;
/**
 * Converts an array of objects to a CSV file and triggers a browser download.
 *
 * Uses no external dependencies for the download step — a temporary object URL
 * is created and immediately revoked after the click. If `data` is empty a
 * warning is logged and the function returns early without error.
 *
 * @param data     - Array of plain objects to export. Must be non-empty.
 * @param fileName - Download filename without extension (`.csv` is appended
 *                   automatically if missing). Defaults to `"export"`.
 * @param opts     - Pass `{ includeBom: true }` to add a UTF-8 BOM, which is
 *                   required for Excel on Windows to detect the encoding.
 * @returns Resolves once the file download has been triggered.
 *
 * @example
 * await exportJsonToCsv(
 *   [{ date: '2025-01-01', amount: 500 }],
 *   `transactions-${new Date().toISOString().slice(0, 10)}`,
 *   { includeBom: true }
 * );
 */
export declare function exportJsonToCsv<T extends Record<string, any>>(data: T[], fileName?: string, opts?: {
    includeBom?: boolean;
}): Promise<void>;
