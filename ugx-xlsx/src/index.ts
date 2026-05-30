/**
 * @module ugx-xlsx
 *
 * Safe, TypeScript-first Excel & CSV export for Angular and modern web apps.
 *
 * **Browser API** (triggers a file download)
 * - {@link exportJsonToExcel} — JSON → XLSX file download
 * - {@link exportJsonToCsv}   — JSON → CSV file download
 *
 * **Universal API** (Node.js + Browser)
 * - {@link jsonToWorkbookBuffer} — JSON → raw XLSX buffer
 * - {@link jsonToCsvString}      — JSON → RFC 4180 CSV string
 *
 * **Types & Errors**
 * - {@link WriteOptions} / {@link CsvOptions} — export configuration
 * - {@link SpreadsheetError} — structured error with machine-readable `code`
 */
export type { WorksheetRow, WriteOptions } from "./core";
export { jsonToWorkbookBuffer } from "./core";
export { SpreadsheetError } from "./errors";
export { exportJsonToExcel } from "./browser-helper";
export { jsonToCsvString, exportJsonToCsv, type CsvOptions } from "./csv";
