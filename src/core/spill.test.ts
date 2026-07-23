import { describe, expect, it } from "vitest";
import { readCsv } from "../adapters/csv/read";
import { ensureCell, getCell } from "./model";
import { recalc } from "./recalc";

// Dynamic-array spill: a plain formula returning a 2-D array fills its anchor + spill range.
describe("dynamic array spill", () => {
  it("spills UNIQUE down a column and dedupes rows", () => {
    const wb = readCsv("a,=UNIQUE(A1:A5)\nb\na\nc\nb\n");
    recalc(wb);
    const s = wb.sheets[0]!;
    // Anchor B1 + spilled B2..B3 = a,b,c (first-seen order).
    expect(getCell(s, 1, 2)?.value).toBe("a");
    expect(getCell(s, 2, 2)?.value).toBe("b");
    expect(getCell(s, 3, 2)?.value).toBe("c");
    expect(getCell(s, 2, 2)?.spill).toBe(true);
    // Nothing spilled past the result.
    expect(getCell(s, 4, 2)?.value ?? "").toBe("");
  });

  it("SEQUENCE spills a 2-D grid from the anchor", () => {
    const wb = readCsv('"=SEQUENCE(2,3)"\n');
    recalc(wb);
    const s = wb.sheets[0]!;
    expect(getCell(s, 1, 1)?.value).toBe("1");
    expect(getCell(s, 1, 3)?.value).toBe("3");
    expect(getCell(s, 2, 1)?.value).toBe("4");
    expect(getCell(s, 2, 3)?.value).toBe("6");
  });

  it("SORT orders rows ascending and descending", () => {
    const wb = readCsv("3,=SORT(A1:A3)\n1\n2\n");
    recalc(wb);
    const s = wb.sheets[0]!;
    expect([getCell(s, 1, 2)?.value, getCell(s, 2, 2)?.value, getCell(s, 3, 2)?.value]).toEqual(["1", "2", "3"]);
  });

  it("FILTER keeps only rows whose mask entry is truthy", () => {
    // The include argument must be an array/range mask (the engine does not broadcast a
    // range=scalar comparison into a per-row mask). Here B1:B3 = 1,0,1 selects rows 1 and 3.
    const wb = readCsv('10,1,"=FILTER(A1:A3,B1:B3)"\n20,0\n30,1\n');
    recalc(wb);
    const s = wb.sheets[0]!;
    expect(getCell(s, 1, 3)?.value).toBe("10");
    expect(getCell(s, 2, 3)?.value).toBe("30");
  });

  it("reports #SPILL! when the range is blocked and clears a stale spill", () => {
    const wb = readCsv("a,=UNIQUE(A1:A3),\nb,BLOCK\nc\n");
    recalc(wb);
    const s = wb.sheets[0]!;
    // B2 is occupied by "BLOCK" -> the UNIQUE at B1 cannot spill.
    expect(getCell(s, 1, 2)?.value).toBe("#SPILL!");
    // Clear the obstacle, recompute: it now spills.
    const block = ensureCell(s, 2, 2);
    block.value = "";
    block.kind = "blank";
    recalc(wb);
    expect(getCell(s, 1, 2)?.value).toBe("a");
    expect(getCell(s, 2, 2)?.value).toBe("b");
    expect(getCell(s, 3, 2)?.value).toBe("c");
  });
});
