import { describe, it, expect } from "vitest";
import { jsonToCsvString } from "../src/csv";

describe("jsonToCsvString", () => {
  it("emits CSV with headers and rows", () => {
    const rows = [
      { Name: "Alice", Score: 95 },
      { Name: "Bob", Score: 88 },
    ];
    const csv = jsonToCsvString(rows);
    // Header + row separators must be CRLF
    expect(csv.startsWith("Name,Score\r\n")).toBe(true);
    expect(csv.includes("Alice")).toBe(true);
    expect(csv.includes("Bob")).toBe(true);
  });

  it("escapes commas, quotes, and newlines correctly", () => {
    const rows = [{ Note: 'he said "hi"', Desc: "x,y", Long: "first\nsecond" }];
    const csv = jsonToCsvString(rows);

    // Exact CSV we expect (CRLF between header and row;
    // CRLF also preserved *inside* the quoted field)
    const expected =
      'Note,Desc,Long\r\n"he said ""hi""","x,y","first\r\nsecond"';

    expect(csv).toBe(expected);
  });

  it("enforces limits", () => {
    const tooManyRows = Array.from({ length: 3 }, () => ({ A: "B" }));
    expect(() => jsonToCsvString(tooManyRows, { maxRows: 2 })).toThrow(
      /Row limit/
    );

    const tooManyCols = [{ a: 1, b: 2, c: 3 }];
    expect(() => jsonToCsvString(tooManyCols, { maxCols: 2 })).toThrow(
      /Column limit/
    );
  });

  it("includes BOM when requested", () => {
    const csv = jsonToCsvString([{ A: "1" }], { includeBom: true });
    expect(csv.startsWith("\uFEFF")).toBe(true);
  });
});
