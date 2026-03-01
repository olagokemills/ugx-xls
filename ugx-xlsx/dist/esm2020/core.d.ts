export type WorksheetRow = Record<string, any>;
export interface WriteOptions {
    sheetName?: string;
    sanitizeHeaders?: boolean;
    maxRows?: number;
    maxCols?: number;
    maxCellChars?: number;
}
export declare const DEFAULT_WRITE_OPTIONS: Required<WriteOptions>;
/** JSON → XLSX buffer (Node/Browser compatible buffer) */
export declare function jsonToWorkbookBuffer(rows: WorksheetRow[], opts?: WriteOptions): Promise<ArrayBuffer | Uint8Array>;
