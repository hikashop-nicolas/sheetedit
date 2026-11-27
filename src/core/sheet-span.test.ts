import { describe, expect, it } from "vitest";
import { getCell, shiftFormula, type Cell, type Sheet, type Workbook } from "./model";
import { rewriteFormula } from "./structure";
import { formulaFromLocal, formulaToLocal } from "./formula-locale";
import { setRegionalLocale } from "./regional";
import { recalc } from "./recalc";
import { setCellInput } from "./workbook";
import { deleteSheet, moveSheet, renameSheet } from "./sheet-ops";
import {
  deleteSheetInFormula,
  expandSheetSpans,
  normalizeSheetSpans,
  renameSheetInFormula,
  scanSheetPrefixes,
} from "./sheet-span";

// 3D references (Sheet1:Sheet3!A1). fast-formula-parser read them as a range operator that
// produced a whole-sheet range, and materializing it ran the tab out of memory.

function sheet(name: string, cells: Record<string, string | number> = {}): Sheet {
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

/** The POI 55906 layout: Sheet1..Sheet3 with numbers in A1(/B1) and text in C1/E1. */
function book(): Workbook {
  return {
    kind: "csv", // no file parts: sheet ops then touch only the model
    files: {},
    sheets: [
      sheet("Sheet1", { A1: 11, C1: "a", E1: "x" }),
      sheet("Sheet2", { A1: 22, B1: 44, C1: "b", E1: "y" }),
      sheet("Sheet3", { A1: 33, C1: "c", E1: "z" }),
      sheet("Total"),
    ],
  };
}

function put(wb: Workbook, at: string, formula: string, on = "Total"): void {
  const s = wb.sheets.find((x) => x.name === on)!;
  const m = /^([A-Z]+)(\d+)$/.exec(at)!;
  const col = [...m[1]!].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  setCellInput(s, Number(m[2]), col, "=" + formula);
}

function val(wb: Workbook, at: string, on = "Total"): string {
  const s = wb.sheets.find((x) => x.name === on)!;
  const m = /^([A-Z]+)(\d+)$/.exec(at)!;
  const col = [...m[1]!].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  return getCell(s, Number(m[2]), col)?.value ?? "";
}

function ev(formula: string, wb = book()): string {
  put(wb, "A1", formula);
  recalc(wb);
  return val(wb, "A1");
}

describe("3D reference parsing", () => {
  it("finds plain, span and quoted prefixes with their reference", () => {
    const p = scanSheetPrefixes("SUM(Sheet1:Sheet3!A1:B2)+'Jan 2026':'Mar 2026'!$C$1+'A b:C d'!D:D+Other!E5");
    expect(p.map((x) => [x.first, x.last, x.ref])).toEqual([
      ["Sheet1", "Sheet3", "A1:B2"],
      ["Jan 2026", "Mar 2026", "$C$1"],
      ["A b", "C d", "D:D"],
      ["Other", undefined, "E5"],
    ]);
    expect(p[0]!.wholeArg).toBe(true);
    expect(p[0]!.fn).toBe("SUM");
    expect(p[1]!.wholeArg).toBe(false);
  });

  it("ignores string literals and a cell-shaped range start", () => {
    expect(scanSheetPrefixes('"Sheet1:Sheet3!A1"&A1')).toEqual([]);
    expect(scanSheetPrefixes("SUM(A1:Sheet3!B2)").map((x) => [x.first, x.last])).toEqual([["Sheet3", undefined]]);
  });

  it("expands a span into per-sheet arguments in tab order", () => {
    const names = ["Sheet1", "Sheet2", "Sheet3"];
    expect(expandSheetSpans("SUM(Sheet1:Sheet3!A1)", names)).toBe("SUM(Sheet1!A1,Sheet2!A1,Sheet3!A1)");
    expect(expandSheetSpans("SUM(Sheet3:Sheet1!A1)", names)).toBe("SUM(Sheet1!A1,Sheet2!A1,Sheet3!A1)");
    expect(expandSheetSpans("_xlfn.STDEV.S(Sheet2:Sheet3!A1,1)", names)).toBe("_xlfn.STDEV.S(Sheet2!A1,Sheet3!A1,1)");
    expect(expandSheetSpans("SUM(Sheet1:Nope!A1)", names)).toBe("SUM(#REF!)");
    expect(expandSheetSpans("VLOOKUP(1,Sheet1:Sheet3!A1:B2,2)", names)).toBe("VLOOKUP(1,#VALUE!,2)");
    expect(expandSheetSpans("Sheet2:Sheet2!A1+1", names)).toBe("Sheet2!A1+1");
  });

  it("quotes expanded names that need it", () => {
    expect(expandSheetSpans("SUM('Jan 2026':'Feb 2026'!A1)", ["Jan 2026", "Feb 2026"])).toBe("SUM('Jan 2026'!A1,'Feb 2026'!A1)");
    expect(expandSheetSpans("SUM('Jan 2026:Feb 2026'!A1)", ["Jan 2026", "Feb 2026"])).toBe("SUM('Jan 2026'!A1,'Feb 2026'!A1)");
  });

  it("normalizes a separately quoted span to Excel's form", () => {
    expect(normalizeSheetSpans("SUM('Jan 2026':'Mar 2026'!A1)")).toBe("SUM('Jan 2026:Mar 2026'!A1)");
    expect(normalizeSheetSpans("SUM('Jan':'Mar'!A1)")).toBe("SUM(Jan:Mar!A1)");
    expect(normalizeSheetSpans("SUM(Jan:Mar!A1)")).toBe("SUM(Jan:Mar!A1)");
  });
});

describe("3D reference evaluation (POI 55906 values)", () => {
  it.each([
    ["SUM(Sheet1:Sheet3!A1)", "66"],
    ["AVERAGE(Sheet1:Sheet3!A1)", "22"],
    ["COUNTA(Sheet1:Sheet3!C1)", "3"],
    ["COUNTA(Sheet1:Sheet3!D1)", "0"],
    ["SUM(Sheet1:Sheet3!A1:B2)", "110"],
    ["AVERAGE(Sheet1:Sheet3!A1:B2)", "27.5"],
    ["MIN(Sheet1:Sheet3!A$1)", "11"],
    ["MAX(Sheet1:Sheet3!A$1:B$2)", "44"],
    ["COUNT(Sheet1:Sheet3!$A$1:$B$2)", "4"],
    ["PRODUCT(Sheet1:Sheet3!A1)", "7986"],
    ["SUM(Sheet1:Sheet3!A1,100)", "166"],
    ["MAXA(Sheet1:Sheet3!A1)", "33"],
    ["MINA(Sheet1:Sheet3!A1)", "11"],
    ["AVERAGEA(Sheet1:Sheet3!A1)", "22"],
    ["VAR.P(Sheet1:Sheet3!A1)", String(242 / 3)],
    ["STDEV(Sheet1:Sheet3!A1)", "11"],
    ["sum(sheet1:SHEET3!a1)", "66"],
  ])("%s = %s", (f, want) => {
    expect(Number(ev(f))).toBeCloseTo(Number(want), 9);
  });

  it("is #REF! for a missing end sheet and #VALUE! outside a 3D function", () => {
    expect(ev("SUM(Sheet1:Gone!A1)")).toBe("#REF!");
    expect(ev("Sheet1:Sheet3!A1+1")).toBe("#VALUE!");
    expect(ev("ABS(Sheet1:Sheet3!A1)")).toBe("#VALUE!");
  });

  it("follows Excel for text and booleans in a span (LibreOffice counts booleans as numbers)", () => {
    const wb = book();
    wb.sheets[0]!.cells.set("1:2", { row: 1, col: 2, value: "TRUE", kind: "b" });
    wb.sheets[1]!.cells.set("1:3", { row: 1, col: 3, value: "5", kind: "s" });
    expect(ev("SUM(Sheet1:Sheet3!B1)", wb)).toBe("44");
    expect(ev("COUNT(Sheet1:Sheet3!A1:C1)", wb)).toBe("4");
    expect(ev("MINA(Sheet1:Sheet3!B1:C1)", wb)).toBe("0");
    expect(ev("AVERAGEA(Sheet1:Sheet3!A1:B1)", wb)).toBe(String((11 + 1 + 22 + 44 + 33) / 5));
    expect(ev("MAXA(Sheet1:Sheet3!C1)", wb)).toBe("0");
  });

  it("includes hidden sheets between the ends", () => {
    const wb = book();
    wb.sheets[1]!.visibility = "hidden";
    expect(ev("SUM(Sheet1:Sheet3!A1)", wb)).toBe("66");
  });

  it("handles quoted names in both forms", () => {
    const wb: Workbook = { kind: "csv", files: {}, sheets: [sheet("Jan 2026", { A1: 1 }), sheet("Feb 2026", { A1: 2 }), sheet("Mar 2026", { A1: 4 }), sheet("Total")] };
    expect(ev("SUM('Jan 2026':'Mar 2026'!A1)", wb)).toBe("7");
    expect(ev("SUM('Jan 2026:Feb 2026'!A1)", wb)).toBe("3");
  });
});

describe("3D reference recalculation and sheet changes", () => {
  it("recomputes when a cell on a middle sheet changes, through a dependent chain", () => {
    const wb = book();
    put(wb, "A1", "SUM(Sheet1:Sheet3!A1)");
    put(wb, "A2", "A1*2");
    put(wb, "A1", "Sheet1!A1*10", "Sheet2"); // formula inside the span, evaluated first
    recalc(wb);
    expect(val(wb, "A1")).toBe(String(11 + 110 + 33));
    expect(val(wb, "A2")).toBe(String((11 + 110 + 33) * 2));
    setCellInput(wb.sheets[0]!, 1, 1, "1");
    recalc(wb);
    expect(val(wb, "A1")).toBe(String(1 + 10 + 33));
  });

  it("follows a sheet moved into or out of the span", () => {
    const wb = book();
    put(wb, "A1", "SUM(Sheet1:Sheet2!A1)");
    recalc(wb);
    expect(val(wb, "A1")).toBe("33");
    moveSheet(wb, 2, 1); // Sheet3 now sits between Sheet1 and Sheet2
    recalc(wb);
    expect(val(wb, "A1")).toBe("66");
    moveSheet(wb, 1, 3); // and out again, past Total
    recalc(wb);
    expect(val(wb, "A1")).toBe("33");
  });

  it("rewrites the formula when an end sheet is renamed", () => {
    const wb = book();
    put(wb, "A1", "SUM(Sheet1:Sheet3!A1)+Sheet3!A1");
    renameSheet(wb, 2, "Mar 2026");
    expect(getCell(wb.sheets[3]!, 1, 1)!.formula).toBe("SUM('Sheet1:Mar 2026'!A1)+'Mar 2026'!A1");
    recalc(wb);
    expect(val(wb, "A1")).toBe("99");
  });

  it("shrinks the span when an end sheet is deleted, and keeps it when a middle one is", () => {
    const wb = book();
    put(wb, "A1", "SUM(Sheet1:Sheet3!A1)");
    deleteSheet(wb, 2);
    expect(getCell(wb.sheets[2]!, 1, 1)!.formula).toBe("SUM(Sheet1:Sheet2!A1)");
    recalc(wb);
    expect(val(wb, "A1")).toBe("33");
    deleteSheet(wb, 1);
    expect(getCell(wb.sheets[1]!, 1, 1)!.formula).toBe("SUM(Sheet1!A1)");
    recalc(wb);
    expect(val(wb, "A1")).toBe("11");

    const wb2 = book();
    put(wb2, "A1", "SUM(Sheet1:Sheet3!A1)");
    deleteSheet(wb2, 1);
    recalc(wb2);
    expect(val(wb2, "A1")).toBe("44");
  });

  it("rewrites helpers directly", () => {
    expect(renameSheetInFormula("SUM(Jan:Mar!A1)+Jan!B2", "jan", "Q1 start")).toBe("SUM('Q1 start:Mar'!A1)+'Q1 start'!B2");
    expect(deleteSheetInFormula("SUM(Jan:Mar!A1)", ["Jan", "Feb", "Mar"], 0)).toBe("SUM(Feb:Mar!A1)");
    expect(deleteSheetInFormula("SUM(Feb:Feb!A1)", ["Jan", "Feb", "Mar"], 1)).toBe("SUM(#REF!)");
  });

  it("normalizes a typed span on entry", () => {
    const wb: Workbook = { kind: "csv", files: {}, sheets: [sheet("Jan 2026"), sheet("Mar 2026")] };
    setCellInput(wb.sheets[0]!, 2, 2, "=SUM('Jan 2026':'Mar 2026'!A1)");
    expect(getCell(wb.sheets[0]!, 2, 2)!.formula).toBe("SUM('Jan 2026:Mar 2026'!A1)");
  });
});

describe("3D reference cost", () => {
  it("evaluates a whole-column span quickly", () => {
    const wb = book();
    for (const s of wb.sheets.slice(0, 3)) for (let r = 1; r <= 2000; r++) s.cells.set(`${r}:4`, { row: r, col: 4, value: "1", kind: "n" });
    put(wb, "A1", "SUM(Sheet1:Sheet3!D:D)");
    put(wb, "A2", "SUM(Sheet1:Sheet3!A:B)");
    put(wb, "A3", "COUNTA(Sheet1:Sheet3!1:1048576)");
    const t = performance.now();
    recalc(wb);
    expect(performance.now() - t).toBeLessThan(1000);
    expect(val(wb, "A1")).toBe("6000");
    expect(val(wb, "A2")).toBe("110");
    expect(val(wb, "A3")).toBe(String(6000 + 10));
  });
});

describe("3D references in other formula rewrites", () => {
  it("shift their cell on copy or fill but keep the sheet names", () => {
    expect(shiftFormula("SUM(Sheet1:Sheet3!A1)", 1, 1)).toBe("SUM(Sheet1:Sheet3!B2)");
    expect(shiftFormula("SUM('Jan 2026:Mar 2026'!A1:A3)", 2, 0)).toBe("SUM('Jan 2026:Mar 2026'!A3:A5)");
  });

  it("are left alone by a row insert on one sheet of the span", () => {
    const op = { axis: "row", kind: "insert", at: 2, count: 1 } as const;
    expect(rewriteFormula("SUM(Sheet1:Sheet3!A5)+Sheet3!A5", "Total", "Sheet3", op)).toBe("SUM(Sheet1:Sheet3!A5)+Sheet3!A6");
    expect(rewriteFormula("SUM('A b':'C d'!A5)", "Total", "C d", op)).toBe("SUM('A b':'C d'!A5)");
  });

  it("survive the French formula display and entry", () => {
    setRegionalLocale("fr");
    try {
      const f = "SUM('Jan 2026:Mar 2026'!A1,Jan:Mar!B2:B10,1.5)";
      expect(formulaToLocal(f)).toBe("SOMME('Jan 2026:Mar 2026'!A1;Jan:Mar!B2:B10;1,5)");
      expect(formulaFromLocal(formulaToLocal(f))).toBe(f);
    } finally {
      setRegionalLocale(null);
    }
  });

  it("cut a whole-sheet range to the used extent instead of materializing it", () => {
    const t = performance.now();
    expect(ev("SUM(Sheet2!A1:XFD1048576)")).toBe("66");
    expect(performance.now() - t).toBeLessThan(1000);
  });
});
