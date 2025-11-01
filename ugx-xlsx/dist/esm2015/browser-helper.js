var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import { jsonToWorkbookBuffer } from "./core";
import { saveAs } from "file-saver";
// Normalize (ArrayBuffer | Buffer/Uint8Array) -> ArrayBuffer
function toArrayBuffer(data) {
    if (data instanceof ArrayBuffer)
        return data;
    const { buffer, byteOffset, byteLength } = data;
    // slice to a standalone ArrayBuffer (not SharedArrayBuffer)
    return buffer.slice(byteOffset, byteOffset + byteLength);
}
export function exportJsonToExcel(rows_1) {
    return __awaiter(this, arguments, void 0, function* (rows, fileName = "export", opts) {
        if (!Array.isArray(rows) || rows.length === 0) {
            console.warn("No data to export.");
            return;
        }
        const raw = yield jsonToWorkbookBuffer(rows, opts);
        // raw is ArrayBuffer in browsers, but TS sees union with Buffer; normalize:
        const ab = raw instanceof ArrayBuffer
            ? raw
            : toArrayBuffer(raw);
        const blob = new Blob([ab], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const safeName = fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`;
        saveAs(blob, safeName);
    });
}
