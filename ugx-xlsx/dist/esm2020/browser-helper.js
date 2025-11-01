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
export async function exportJsonToExcel(rows, fileName = "export", opts) {
    if (!Array.isArray(rows) || rows.length === 0) {
        console.warn("No data to export.");
        return;
    }
    const raw = await jsonToWorkbookBuffer(rows, opts);
    // raw is ArrayBuffer in browsers, but TS sees union with Buffer; normalize:
    const ab = raw instanceof ArrayBuffer
        ? raw
        : toArrayBuffer(raw);
    const blob = new Blob([ab], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const safeName = fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`;
    saveAs(blob, safeName);
}
