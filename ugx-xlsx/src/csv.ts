import { sanitizeHeader, enforceLimits } from './sanitizer';
import { SpreadsheetError } from './errors';

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
export const DEFAULT_CSV_OPTIONS: Required<Omit<CsvOptions, 'headers'>> = {
  sanitizeHeaders: true,
  maxRows: 100_000,
  maxCols: 1000,
  maxCellChars: 10_000,
  includeBom: false,
};

/** Normalises any mix of CR, LF, or CRLF line endings to CRLF as required by RFC 4180. */
function normalizeNewlinesToCRLF(s: string): string {
  return s.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n/g, '\r\n');
}

/**
 * RFC 4180 field escaping: wraps the value in double-quotes if it contains a
 * comma, double-quote, or line ending, and doubles any inner double-quotes.
 */
function csvEscape(value: string): string {
  if (value === undefined || value === null) return '';
  let s = String(value);
  s = normalizeNewlinesToCRLF(s);
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

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
export function jsonToCsvString(rows: CsvRow[], opts?: CsvOptions): string {
  const options = { ...DEFAULT_CSV_OPTIONS, ...(opts || {}) };

  if (!Array.isArray(rows))
    throw new SpreadsheetError('rows must be an array', 'INVALID_INPUT');
  if (rows.length === 0) return options.includeBom ? '﻿' : '';

  let headers = options.headers?.length
    ? [...options.headers]
    : Object.keys(rows[0]);
  if (options.sanitizeHeaders) headers = headers.map(sanitizeHeader);

  if (headers.length > options.maxCols)
    throw new SpreadsheetError('Column limit exceeded', 'LIMIT_COLS');
  if (rows.length > options.maxRows)
    throw new SpreadsheetError('Row limit exceeded', 'LIMIT_ROWS');

  const lines: string[] = [];
  if (options.includeBom) lines.push('﻿');

  lines.push(headers.map((h) => csvEscape(h)).join(','));

  for (let i = 0; i < rows.length; i++) {
    if (i + 1 > options.maxRows)
      throw new SpreadsheetError(
        'Row limit exceeded during write',
        'LIMIT_ROWS'
      );
    const r = rows[i];
    const fields: string[] = [];
    for (let c = 0; c < headers.length; c++) {
      if (c + 1 > options.maxCols)
        throw new SpreadsheetError(
          'Column limit exceeded during write',
          'LIMIT_COLS'
        );
      const h = headers[c];
      let v = r[h];
      if (v === undefined || v === null) v = '';
      if (typeof v === 'object') {
        try {
          v = JSON.stringify(v);
        } catch {
          v = String(v);
        }
      } else {
        v = String(v);
      }
      v = enforceLimits(v, options.maxCellChars);
      fields.push(csvEscape(v));
    }
    lines.push(fields.join(','));
  }

  return lines.join('\r\n');
}

/**
 * Minimal, dependency-free browser file download using a temporary object URL.
 * Falls back to `msSaveOrOpenBlob` for legacy IE/Edge.
 */
function downloadBlob(blob: Blob, filename: string) {
  const navAny = window.navigator as any;
  if (navAny && typeof navAny.msSaveOrOpenBlob === 'function') {
    navAny.msSaveOrOpenBlob(blob, filename);
    return;
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

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
export async function exportJsonToCsv<T extends Record<string, any>>(
  data: T[],
  fileName = 'export',
  opts?: { includeBom?: boolean }
): Promise<void> {
  if (!Array.isArray(data) || data.length === 0) {
    console.warn('No data to export.');
    return;
  }

  const csv = jsonToCsvString(data, opts);
  const parts = (opts?.includeBom ? ['﻿', csv] : [csv]) as BlobPart[];
  const blob = new Blob(parts, { type: 'text/csv;charset=utf-8' });
  const safeName = fileName.endsWith('.csv') ? fileName : `${fileName}.csv`;
  downloadBlob(blob, safeName);
}
