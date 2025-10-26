// src/csv.ts
import { sanitizeHeader, enforceLimits } from './sanitizer';
import { SpreadsheetError } from './errors';
export const DEFAULT_CSV_OPTIONS = {
    sanitizeHeaders: true,
    maxRows: 100000,
    maxCols: 1000,
    maxCellChars: 10000,
    includeBom: false,
};
/** Normalize any newline to CRLF (RFC 4180) */
function normalizeNewlinesToCRLF(s) {
    return s.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n/g, '\r\n');
}
/** RFC-4180 escape: quote if contains comma, quote, or CR/LF. Double inner quotes. */
function csvEscape(value) {
    if (value === undefined || value === null)
        return '';
    let s = String(value);
    s = normalizeNewlinesToCRLF(s);
    if (/[",\r\n]/.test(s))
        return `"${s.replace(/"/g, '""')}"`;
    return s;
}
/** Convert JSON rows -> CSV string with safety limits */
export function jsonToCsvString(rows, opts) {
    const options = { ...DEFAULT_CSV_OPTIONS, ...(opts || {}) };
    if (!Array.isArray(rows))
        throw new SpreadsheetError('rows must be an array', 'INVALID_INPUT');
    if (rows.length === 0)
        return options.includeBom ? '\uFEFF' : '';
    let headers = options.headers?.length
        ? [...options.headers]
        : Object.keys(rows[0]);
    if (options.sanitizeHeaders)
        headers = headers.map(sanitizeHeader);
    if (headers.length > options.maxCols)
        throw new SpreadsheetError('Column limit exceeded', 'LIMIT_COLS');
    if (rows.length > options.maxRows)
        throw new SpreadsheetError('Row limit exceeded', 'LIMIT_ROWS');
    const lines = [];
    if (options.includeBom)
        lines.push('\uFEFF');
    lines.push(headers.map((h) => csvEscape(h)).join(','));
    for (let i = 0; i < rows.length; i++) {
        if (i + 1 > options.maxRows)
            throw new SpreadsheetError('Row limit exceeded during write', 'LIMIT_ROWS');
        const r = rows[i];
        const fields = [];
        for (let c = 0; c < headers.length; c++) {
            if (c + 1 > options.maxCols)
                throw new SpreadsheetError('Column limit exceeded during write', 'LIMIT_COLS');
            const h = headers[c];
            let v = r[h];
            if (v === undefined || v === null)
                v = '';
            if (typeof v === 'object') {
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
            v = enforceLimits(v, options.maxCellChars);
            fields.push(csvEscape(v));
        }
        lines.push(fields.join(','));
    }
    return lines.join('\r\n');
}
/** Minimal, dependency-free browser saver */
function downloadBlob(blob, filename) {
    // Legacy IE/Edge (very old)
    const navAny = window.navigator;
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
/** Browser helper: JSON -> CSV download */
export async function exportJsonToCsv(data, fileName = 'export', opts) {
    if (!Array.isArray(data) || data.length === 0) {
        console.warn('No data to export.');
        return;
    }
    const csv = jsonToCsvString(data, opts);
    const parts = (opts?.includeBom ? ['\uFEFF', csv] : [csv]);
    const blob = new Blob(parts, { type: 'text/csv;charset=utf-8' });
    const safeName = fileName.endsWith('.csv') ? fileName : `${fileName}.csv`;
    downloadBlob(blob, safeName);
}
