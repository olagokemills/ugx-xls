import { sanitizeHeader, enforceLimits } from "./sanitizer";
import { SpreadsheetError } from "./errors";
import writeXlsxFile from "write-excel-file";
import type { Cell, SheetData } from "write-excel-file";

/** A single data row: any plain object whose values will be serialised to strings. */
export type WorksheetRow = Record<string, any>;

/** Options for {@link jsonToWorkbookBuffer} and {@link exportJsonToExcel}. */
export interface WriteOptions {
  /** Name of the worksheet tab. @default "Sheet1" */
  sheetName?: string;
  /** Strip ASCII control chars from header keys before writing. @default true */
  sanitizeHeaders?: boolean;
  /** Maximum number of data rows allowed (guard against runaway exports). @default 100_000 */
  maxRows?: number;
  /** Maximum number of columns (keys per row). @default 1_000 */
  maxCols?: number;
  /** Maximum characters per cell; longer values are hard-truncated. @default 10_000 */
  maxCellChars?: number;
}

/** Merged defaults applied when individual {@link WriteOptions} fields are omitted. */
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

/** Returns `true` when running inside Node.js (not a browser). */
function isNodeRuntime(): boolean {
  const g = globalThis as {
    process?: { versions?: { node?: string } };
  };
  return !!g.process?.versions?.node;
}

/**
 * Dynamically loads `write-excel-file/node` via CJS `require`.
 * Returns `undefined` when `require` is not available (browser environments).
 */
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

/**
 * Converts an array of plain objects into a raw XLSX workbook buffer.
 *
 * Works in both **Node.js** (returns `Uint8Array`) and **browsers**
 * (returns `ArrayBuffer`). Object keys become column headers; values are
 * coerced to strings. `null` / `undefined` cells are written as empty strings.
 * Objects are JSON-serialised; all other primitives use `String()`.
 *
 * @param rows - Array of plain objects to serialise. Column order is derived
 *               from the keys of the first row.
 * @param opts - Optional overrides for sheet name, safety limits, and sanitisation.
 * @returns A raw XLSX buffer suitable for saving to disk or streaming to a client.
 *
 * @throws {SpreadsheetError} `INVALID_INPUT` — `rows` is not an array.
 * @throws {SpreadsheetError} `LIMIT_ROWS`    — row count exceeds `maxRows`.
 * @throws {SpreadsheetError} `LIMIT_COLS`    — column count exceeds `maxCols`.
 *
 * @example
 * // Node.js — write buffer to disk
 * const buf = await jsonToWorkbookBuffer(
 *   [{ name: 'Alice', score: 42 }, { name: 'Bob', score: 98 }],
 *   { sheetName: 'Results' }
 * );
 * fs.writeFileSync('results.xlsx', Buffer.from(buf));
 *
 * @example
 * // Browser — wrap in a Blob for download
 * const buf = await jsonToWorkbookBuffer(rows);
 * const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
 * saveAs(blob, 'export.xlsx');
 */
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
