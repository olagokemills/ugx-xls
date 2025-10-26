import { describe, it, expect } from "vitest";
import { jsonToWorkbookBuffer } from "../src/index.js";

describe("jsonToWorkbookBuffer", () => {
  it("builds an xlsx buffer from small JSON", async () => {
    const rows = [
      { A: "1", B: "two" },
      { A: "3", B: "four" },
    ];

    const buf = await jsonToWorkbookBuffer(rows, { sheetName: "Test" });
    const size = (buf as any).byteLength ?? (buf as any).length ?? 0;
    expect(size).toBeGreaterThan(0);
  });

  it("enforces row limit", async () => {
    const rows = Array.from({ length: 2 }, () => ({ X: "y" }));
    await expect(jsonToWorkbookBuffer(rows, { maxRows: 1 })).rejects.toThrow(
      /Row limit exceeded/
    );
  });

  it("handles empty array", async () => {
    const buf = await jsonToWorkbookBuffer([]);
    const size = (buf as any).byteLength ?? (buf as any).length ?? 0;
    expect(size).toBeGreaterThan(0);
  });
});
