export type CsvRow = Record<string, any>;
export interface CsvOptions {
    sanitizeHeaders?: boolean;
    maxRows?: number;
    maxCols?: number;
    maxCellChars?: number;
    headers?: string[];
    includeBom?: boolean;
}
export declare const DEFAULT_CSV_OPTIONS: Required<Omit<CsvOptions, 'headers'>>;
/** Convert JSON rows -> CSV string with safety limits */
export declare function jsonToCsvString(rows: CsvRow[], opts?: CsvOptions): string;
/** Browser helper: JSON -> CSV download */
export declare function exportJsonToCsv<T extends Record<string, any>>(data: T[], fileName?: string, opts?: {
    includeBom?: boolean;
}): Promise<void>;
