import { describe, expect, it } from "vitest";
import { getCell, type Cell, type Sheet, type Workbook } from "./model";
import { recalc } from "./recalc";
import { unanchorLines } from "./whole-lines";

// $A:$A read as about 727 columns of a million rows each, and $1:$1 as an error: a workbook with
// dynamic names over whole columns (OFFSET(..., COUNTA($A:$A))) froze on its first edit.

describe("whole columns and rows", () => {
  it("lose their anchors before parsing", () => {
    expect(unanchorLines("COUNTA(Data!$A:$A)+SUM($B:C)")).toBe("COUNTA(Data!A:A)+SUM(B:C)");
    expect(unanchorLines("SUM($1:$2)")).toBe("SUM(1:2)");
    expect(unanchorLines('SUM($A$1:$B$2)&"$A:$A"')).toBe('SUM($A$1:$B$2)&"$A:$A"');
  });

  it("compute over the used cells only", () => {
    const s: Sheet = { name: "Data", cells: new Map<string, Cell>(), maxRow: 3, maxCol: 1 };
    for (const r of [1, 2, 3]) s.cells.set(`${r}:1`, { row: r, col: 1, value: String(r), kind: "n" });
    s.cells.set("1:3", { row: 1, col: 3, value: "", kind: "blank", formula: "SUM($A:$A)+COUNTA(Data!$A:$A)" });
    s.maxCol = 3;
    const wb: Workbook = { kind: "xlsx", sheets: [s], files: {} };
    const t = performance.now();
    recalc(wb);
    expect(getCell(s, 1, 3)?.value).toBe("9");
    expect(performance.now() - t).toBeLessThan(2000);
  });
});
