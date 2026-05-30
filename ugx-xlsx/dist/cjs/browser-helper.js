"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.exportJsonToExcel = exportJsonToExcel;
const core_1 = require("./core");
const file_saver_1 = require("file-saver");
/**
 * Normalises a Node.js `Buffer` / `Uint8Array` into a standalone `ArrayBuffer`.
 *
 * The `.buffer` property of typed arrays may reference a shared backing store,
 * so we `slice` to produce an isolated copy before wrapping it in a `Blob`.
 */
function toArrayBuffer(data) {
    if (data instanceof ArrayBuffer)
        return data;
    const { buffer, byteOffset, byteLength } = data;
    // slice to a standalone ArrayBuffer (not SharedArrayBuffer)
    return buffer.slice(byteOffset, byteOffset + byteLength);
}
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
function exportJsonToExcel(rows_1) {
    return __awaiter(this, arguments, void 0, function* (rows, fileName = "export", opts) {
        if (!Array.isArray(rows) || rows.length === 0) {
            console.warn("No data to export.");
            return;
        }
        const raw = yield (0, core_1.jsonToWorkbookBuffer)(rows, opts);
        // raw is ArrayBuffer in browsers, but TS sees union with Buffer; normalize:
        const ab = raw instanceof ArrayBuffer
            ? raw
            : toArrayBuffer(raw);
        const blob = new Blob([ab], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const safeName = fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`;
        (0, file_saver_1.saveAs)(blob, safeName);
    });
}
