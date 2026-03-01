import { sanitizeHeader, enforceLimits } from "./sanitizer";
import { SpreadsheetError } from "./errors";
import writeXlsxFile from "write-excel-file";
import type { Cell, SheetData } from "write-excel-file";

export type WorksheetRow = Record<string, any>;

export interface WriteOptions {
  sheetName?: string;
  sanitizeHeaders?: boolean;
  maxRows?: number; // default 100_000
  maxCols?: number; // default 1000
  maxCellChars?: number; // default 10_000
}

export const DEFAULT_WRITE_OPTIONS: Required<WriteOptions> = {
  sheetName: "Sheet1",
  sanitizeHeaders: true,
  maxRows: 100_000,
  maxCols: 1000,
  maxCellChars: 10_000,
};

type CellValue = string | number | boolean | Date;
type NodeWriteXlsxFn = (
  data: SheetData,
  options: { buffer: true; sheet?: string }
) => Promise<Uint8Array>;

function isNodeRuntime(): boolean {
  const g = globalThis as {
    process?: { versions?: { node?: string } };
  };
  return !!g.process?.versions?.node;
}

function getNodeWriter(): NodeWriteXlsxFn | undefined {
  const g = globalThis as { require?: (name: string) => unknown };
  if (typeof g.require === "function") {
    const loaded = g.require("write-excel-file/node") as
      | { default?: NodeWriteXlsxFn }
      | NodeWriteXlsxFn;
    if (typeof loaded === "function") return loaded;
    return loaded.default;
  }
  return undefined;
}

/** JSON → XLSX buffer (Node/Browser compatible buffer) */
export async function jsonToWorkbookBuffer(
  rows: WorksheetRow[],
  opts?: WriteOptions
): Promise<ArrayBuffer | Uint8Array> {
  const options = { ...DEFAULT_WRITE_OPTIONS, ...(opts || {}) };

  if (!Array.isArray(rows)) {
    throw new SpreadsheetError("rows must be an array", "INVALID_INPUT");
  }
  if (rows.length > options.maxRows) {
    throw new SpreadsheetError("Row limit exceeded", "LIMIT_ROWS");
  }

  const headers = (rows[0] ? Object.keys(rows[0]) : [])
    .slice(0, options.maxCols)
    .map((h) => (options.sanitizeHeaders ? sanitizeHeader(h) : h));

  if (headers.length > options.maxCols) {
    throw new SpreadsheetError("Column limit exceeded", "LIMIT_COLS");
  }

  const data: SheetData = [];

  if (rows.length === 0) {
    data.push([]);
  } else {
    data.push(headers.map((h) => ({ value: h } as Cell)));
  }

  for (let i = 0; i < rows.length; i++) {
    if (i + 1 > options.maxRows) {
      throw new SpreadsheetError(
        "Row limit exceeded during write",
        "LIMIT_ROWS"
      );
    }

    const r = rows[i];
    const values: Cell[] = [];

    for (let c = 0; c < headers.length; c++) {
      if (c + 1 > options.maxCols) {
        throw new SpreadsheetError(
          "Column limit exceeded during write",
          "LIMIT_COLS"
        );
      }
      const key = headers[c];
      let v = r[key];

      if (v === undefined || v === null) {
        values.push({ value: "" } as Cell);
        continue;
      }
      if (typeof v === "object") {
        try {
          v = JSON.stringify(v);
        } catch {
          v = String(v);
        }
      } else {
        v = String(v);
      }

      values.push({ value: enforceLimits(v, options.maxCellChars) } as Cell);
    }

    data.push(values);
  }

  if (isNodeRuntime()) {
    const nodeWriter = getNodeWriter();
    if (nodeWriter) {
      return nodeWriter(data, {
        buffer: true,
        sheet: options.sheetName,
      });
    }
  }

  const out = await writeXlsxFile(data, { sheet: options.sheetName });
  return out.arrayBuffer();
}
