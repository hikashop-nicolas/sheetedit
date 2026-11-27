import { describe, expect, it } from "vitest";
import { getCell, type Cell, type Sheet, type Workbook } from "./model";
import { recalc } from "./recalc";
import { cutWholeColumnArgs, unanchorLines } from "./whole-lines";

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

describe("a whole column passed straight to a blank-proof function", () => {
  it("is cut at the used rows whatever encloses it", () => {
    expect(cutWholeColumnArgs("OFFSET(Data!$A$3,0,0,COUNTA(Data!A:A)-2,1)", 40)).toBe("OFFSET(Data!$A$3,0,0,COUNTA(Data!A1:A40)-2,1)");
    expect(cutWholeColumnArgs("IF(x,MATCH(1,'My data'!B:B,0),SUM(C:D))", 9)).toBe("IF(x,MATCH(1,'My data'!B1:B9,0),SUM(C1:D9))");
  });

  it("is left whole where blanks or positions count", () => {
    for (const f of ["INDEX(A:A,500)", 'COUNTIF(A:A,"")', "COUNTBLANK(A:A)", "SUM(A:A*2)", "ROWS(A:A)", 'SUM("A:A")', "OFFSET(A:A,1,0)"])
      expect(cutWholeColumnArgs(f, 9)).toBe(f);
  });

  // The shape of POI's Basic_Expense_Template_2011.xls: dynamic names over whole columns, read
  // through INDIRECT(ADDRESS(...)) by hundreds of cells. It took minutes and froze the tab.
  it("keeps dynamic whole-column names cheap and right", () => {
    const data: Sheet = { name: "Types", cells: new Map<string, Cell>(), maxRow: 0, maxCol: 0 };
    const put = (s: Sheet, r: number, c: number, cell: Partial<Cell>) => {
      s.cells.set(`${r}:${c}`, { row: r, col: c, value: "", kind: "blank", ...cell });
      s.maxRow = Math.max(s.maxRow, r);
      s.maxCol = Math.max(s.maxCol, c);
    };
    put(data, 1, 1, { value: "Name", kind: "s" });
    put(data, 1, 2, { value: "Key", kind: "s" });
    for (let r = 2; r <= 101; r++) {
      put(data, r, 1, { value: `type ${r}`, kind: "s" });
      put(data, r, 2, { value: `K${r}`, kind: "s" });
    }
    const input: Sheet = { name: "Input", cells: new Map<string, Cell>(), maxRow: 0, maxCol: 0 };
    const out: Sheet = { name: "Out", cells: new Map<string, Cell>(), maxRow: 0, maxCol: 0 };
    for (let r = 1; r <= 400; r++) {
      put(input, r, 1, { value: `type ${(r % 100) + 2}`, kind: "s" });
      put(out, r, 1, {
        formula:
          'IF(ISNA(MATCH(INDIRECT(ADDRESS(ROW(),1,1,TRUE,"Input")),TypeList,0)),"",INDEX(TypeKeys,MATCH(INDIRECT(ADDRESS(ROW(),1,1,TRUE,"Input")),TypeList,0),1,1))',
      });
    }
    const wb: Workbook = {
      kind: "xlsx",
      files: {},
      sheets: [input, out, data],
      definedNames: new Map([
        ["TypeList", "OFFSET(Types!$A$2,0,0,COUNTA(Types!$A:$A)-1,1)"],
        ["TypeKeys", "OFFSET(Types!$B$2,0,0,COUNTA(Types!$B:$B)-1,1)"],
      ]),
    };
    let t = performance.now();
    recalc(wb);
    expect(performance.now() - t).toBeLessThan(10000);
    expect(getCell(out, 7, 1)?.value).toBe("K9");
    expect(getCell(out, 98, 1)?.value).toBe("K100");
    // An edit recalculates as fast, and a name pointed elsewhere is followed.
    input.cells.get("7:1")!.value = "type 50";
    t = performance.now();
    recalc(wb);
    expect(performance.now() - t).toBeLessThan(5000);
    expect(getCell(out, 7, 1)?.value).toBe("K50");
    wb.definedNames!.set("TypeKeys", "OFFSET(Types!$A$2,0,0,COUNTA(Types!$A:$A)-1,1)");
    recalc(wb);
    expect(getCell(out, 7, 1)?.value).toBe("type 50");
  });
});
