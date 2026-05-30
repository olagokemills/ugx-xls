import { WriteOptions, WorksheetRow } from "./core";
/**
 * Converts an array of objects to an XLSX file and triggers a browser download.
 *
 * This is a thin browser wrapper around {@link jsonToWorkbookBuffer}. It has no
 * effect outside browser contexts (file-saver requires `window`). If `rows` is
 * empty a warning is logged and the function returns early without error.
 *
 * @param rows     - Array of plain objects to export. Must be non-empty.
 * @param fileName - Download filename without extension (`.xlsx` is appended
 *                   automatically if missing). Defaults to `"export"`.
 * @param opts     - Optional {@link WriteOptions} (sheet name, limits, sanitisation).
 * @returns Resolves once the file download has been triggered.
 *
 * @example
 * await exportJsonToExcel(
 *   [{ product: 'Widget', units: 120, price: 9.99 }],
 *   `inventory-${new Date().toISOString().slice(0, 10)}`,
 *   { sheetName: 'Stock' }
 * );
 */
export declare function exportJsonToExcel(rows: WorksheetRow[], fileName?: string, opts?: WriteOptions): Promise<void>;
