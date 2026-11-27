import { describe, expect, it } from "vitest";
import { getCell, type Cell, type Sheet, type Workbook } from "./model";
import { recalc } from "./recalc";

// A formula whose result is a reference to an empty cell shows 0 in Excel (=A9, INDEX, VLOOKUP,
// CHOOSE, OFFSET...). Only a real empty string ("" or a cell holding one) shows as empty.

function book(): Workbook {
  const s: Sheet = { name: "S", cells: new Map<string, Cell>(), maxRow: 0, maxCol: 0 };
  const put = (row: number, col: number, c: Partial<Cell>) => {
    s.cells.set(`${row}:${col}`, { row, col, value: "", kind: "blank", ...c });
    s.maxRow = Math.max(s.maxRow, row);
    s.maxCol = Math.max(s.maxCol, col);
  };
  put(1, 1, { value: "k1", kind: "s" });
  put(2, 1, { value: "k2", kind: "s" });
  put(3, 1, { value: "k3", kind: "s" });
  put(1, 2, { value: "10", kind: "n" });
  put(3, 2, { value: "30", kind: "n" }); // B2 left empty
  put(1, 3, { formula: '""' }); // C1 returns an empty string
  return { kind: "xlsx", files: {}, sheets: [s, { name: "Other", cells: new Map(), maxRow: 0, maxCol: 0 }] };
}

function ev(formulas: string[]): { value: string; kind: string }[] {
  const wb = book();
  const s = wb.sheets[0]!;
  formulas.forEach((f, i) => s.cells.set(`${20 + i}:5`, { row: 20 + i, col: 5, value: "", kind: "blank", formula: f }));
  recalc(wb);
  return formulas.map((_, i) => {
    const c = getCell(s, 20 + i, 5)!;
    return { value: c.value, kind: c.kind };
  });
}

describe("a formula that returns an empty cell", () => {
  it("shows 0, as in Excel", () => {
    const got = ev([
      "A9",
      "INDEX(B1:B3,2)",
      "INDEX(B:B,100)",
      'VLOOKUP("k2",A1:B3,2,FALSE)',
      "CHOOSE(2,A1,A9)",
      "OFFSET(A1,20,0)",
      'INDIRECT("Z99")',
      "IF(TRUE,A9)",
      "Other!Z9",
      'XLOOKUP("k2",A1:A3,B1:B3)',
      "A9+1",
    ]);
    expect(got.map((g) => g.value)).toEqual(["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "1"]);
    expect(got[0]!.kind).toBe("n");
  });

  it("keeps a real empty string empty", () => {
    const got = ev(['""', 'IF(1>0,"","x")', "C1", 'C1&""', 'A9&""', 'IF(C1="","blank","no")', 'VLOOKUP("k2",A1:B3,2,FALSE)&""', 'LEN(VLOOKUP("k2",A1:B3,2,FALSE))']);
    expect(got.map((g) => g.value)).toEqual(["", "", "", "", "", "blank", "", "0"]);
    expect(got[2]!.kind).toBe("s");
  });
});
