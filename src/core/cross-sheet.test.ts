import { describe, expect, it } from "vitest";
import { getCell, type Cell, type Sheet, type Workbook } from "./model";
import { recalc } from "./recalc";

// INDEX+MATCH into a lookup sheet is the single most common cross-sheet idiom in a real workbook.
// fast-formula-parser builds INDEX's result reference from the source range's row and column and
// drops the sheet, so the value came back off whichever sheet the formula lived on: a grade table
// answered with whatever happened to sit at the same coordinates next door, with no error to show
// for it. reference-fns puts the sheet back.

function sheet(name: string, cells: Record<string, string | number>): Sheet {
  const s: Sheet = { name, cells: new Map<string, Cell>(), maxRow: 0, maxCol: 0 };
  for (const [ref, v] of Object.entries(cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
    const col = [...m[1]!].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
    const row = Number(m[2]);
    s.cells.set(`${row}:${col}`, { row, col, value: String(v), kind: typeof v === "number" ? "n" : "s" });
    s.maxRow = Math.max(s.maxRow, row);
    s.maxCol = Math.max(s.maxCol, col);
  }
  return s;
}

/** Put `formula` at A1 of the first sheet and return what it computes to. */
function ev(formula: string, sheets: Sheet[]): string {
  const wb: Workbook = { kind: "xlsx", sheets, files: {} };
  const s = sheets[0]!;
  s.cells.set("1:1", { row: 1, col: 1, value: "", kind: "blank", formula });
  s.maxRow = Math.max(s.maxRow, 1);
  s.maxCol = Math.max(s.maxCol, 1);
  recalc(wb);
  return getCell(s, 1, 1)?.value ?? "";
}

// Both sheets carry data at B5:B7 and D5:D7, so reading the right cell off the wrong sheet
// returns a plausible wrong answer rather than an error. That is exactly how the bug hid.
const pair = (): Sheet[] => [
  sheet("Main", { B5: 1900, B6: 1901, B7: 1902, D5: 1910, D6: 1911, D7: 1912 }),
  sheet("Grid", { B5: 0, B6: 1, B7: 2, D5: "A", D6: "B", D7: "C" }),
];

describe("INDEX into another sheet", () => {
  it("reads the named sheet, not the one the formula sits on", () => {
    expect(ev("INDEX(Grid!$D$5:$D$7,2)", pair())).toBe("B");
  });

  it("carries the sheet through the INDEX+MATCH idiom", () => {
    expect(ev("INDEX(Grid!$D$5:$D$7,MATCH(2,Grid!$B$5:$B$7,0))", pair())).toBe("C");
  });

  it("works for a sheet whose name is not ASCII", () => {
    const sheets = [sheet("0", { D6: 1911 }), sheet("グリル", { D5: "A", D6: "B", D7: "C" })];
    expect(ev("INDEX(グリル!$D$5:$D$7,2)", sheets)).toBe("B");
  });

  // A whole-column INDEX returns the range itself (which kept its sheet); check the fix did not
  // disturb that form, and that a plain same-sheet INDEX still reads its own sheet.
  it("leaves the same-sheet and whole-column forms alone", () => {
    expect(ev("INDEX($D$5:$D$7,2)", pair())).toBe("1911");
    expect(ev("SUM(INDEX(Grid!$B$5:$B$7,0,1))", pair())).toBe("3"); // 0+1+2 off Grid
  });
});
