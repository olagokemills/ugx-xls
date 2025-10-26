import { WriteOptions, WorksheetRow } from "./core";
export declare function exportJsonToExcel(rows: WorksheetRow[], fileName?: string, opts?: WriteOptions): Promise<void>;
