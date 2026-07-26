import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../index";
import { getCell, type Workbook } from "./model";
import { parseModule } from "./vba-parse";
import { VbaInterpreter, type RunResult } from "./vba-run";
import { EMPTY, VbaArray, type VbaValue } from "./vba-value";
import { excelGlobals, type ExcelHost } from "./vba-excel";

// Stage 3 of _plans/VBA_PLAN.md. The object model a macro talks to, mapped onto sheetedit's own
// model so a macro edit is an ordinary edit. What is NOT modelled must refuse by name, which is
// half of what these tests check: a plausible wrong answer is worse than a stop.

/** A workbook with no file behind it: enough for everything except styling. */
function bareWorkbook(rows: (string | number)[][] = []): Workbook {
  const wb: Workbook = { kind: "xlsx", sheets: [], files: {} };
  const sheet = { name: "Sheet1", cells: new Map(), maxRow: 0, maxCol: 0 };
  wb.sheets.push(sheet);
  rows.forEach((row, r) => row.forEach((v, c) => {
    const numeric = typeof v === "number";
    sheet.cells.set(`${r + 1}:${c + 1}`, {
      row: r + 1, col: c + 1, value: String(v), kind: numeric ? "n" : "s",
    });
    sheet.maxRow = Math.max(sheet.maxRow, r + 1);
    sheet.maxCol = Math.max(sheet.maxCol, c + 1);
  }));
  return wb;
}

function run(src: string, wb: Workbook, host: Partial<ExcelHost> = {}): { wb: Workbook; result: RunResult } {
  const full: ExcelHost = { wb, activeSheet: 0, ...host };
  const m = parseModule(`Sub T()\n${src}\nEnd Sub`);
  return { wb, result: new VbaInterpreter(m, { globals: excelGlobals(full) }).run("T") };
}

/** Run an expression and hand back its value. */
function evalIn(expr: string, wb: Workbook, host: Partial<ExcelHost> = {}): VbaValue {
  const full: ExcelHost = { wb, activeSheet: 0, ...host };
  const m = parseModule(`Function T()\nT = ${expr}\nEnd Function`);
  return new VbaInterpreter(m, { globals: excelGlobals(full) }).run("T").value ?? EMPTY;
}

const at = (wb: Workbook, r: number, c: number): string | undefined => getCell(wb.sheets[0]!, r, c)?.value;

describe("Range values", () => {
  it("reads a cell, and a blank one as Empty", () => {
    const wb = bareWorkbook([[1, "two"]]);
    expect(evalIn('Range("A1").Value', wb)).toBe(1);
    expect(evalIn('Range("B1").Value', wb)).toBe("two");
    expect(evalIn('Range("Z9").Value', wb)).toBe(EMPTY);
    // A Range used where a value is wanted gives its Value, which is its default member.
    expect(evalIn('Range("A1") + 1', wb)).toBe(2);
  });

  it("writes numbers, strings and Booleans with the right kind", () => {
    const { wb } = run('Range("A1").Value = 42\nRange("A2").Value = "hi"\nRange("A3").Value = True', bareWorkbook());
    expect(at(wb, 1, 1)).toBe("42");
    expect(getCell(wb.sheets[0]!, 1, 1)?.kind).toBe("n");
    expect(at(wb, 2, 1)).toBe("hi");
    expect(getCell(wb.sheets[0]!, 3, 1)?.kind).toBe("b");
  });

  it("writes a leading = as a formula, as Excel does", () => {
    const { wb } = run('Range("A1").Value = "=SUM(1,2)"', bareWorkbook());
    expect(getCell(wb.sheets[0]!, 1, 1)?.formula).toBe("SUM(1,2)");
  });

  it("fills every cell of a multi-cell range from one value", () => {
    const { wb } = run('Range("A1:B2").Value = 7', bareWorkbook());
    expect([at(wb, 1, 1), at(wb, 1, 2), at(wb, 2, 1), at(wb, 2, 2)]).toEqual(["7", "7", "7", "7"]);
  });

  it("reads a block as a 1-based two-dimensional array", () => {
    const wb = bareWorkbook([[1, 2], [3, 4]]);
    const v = evalIn('Range("A1:B2").Value', wb);
    expect(v).toBeInstanceOf(VbaArray);
    const arr = v as VbaArray;
    expect(arr.lower).toEqual([1, 1]);
    expect(arr.upper).toEqual([2, 2]);
    expect(arr.get([2, 1])).toBe(3);
  });

  it("writes a block back from an array corner to corner", () => {
    const wb = bareWorkbook([[1, 2], [3, 4]]);
    run('Dim v\nv = Range("A1:B2").Value\nRange("D1:E2").Value = v', wb);
    expect([at(wb, 1, 4), at(wb, 1, 5), at(wb, 2, 4), at(wb, 2, 5)]).toEqual(["1", "2", "3", "4"]);
  });

  it("gives Text the formatted display and Formula the formula", () => {
    const wb = bareWorkbook();
    const sheet = wb.sheets[0]!;
    sheet.cells.set("1:1", { row: 1, col: 1, value: "0.5", kind: "n", display: "50%", formula: "A2/2" });
    sheet.maxRow = 1; sheet.maxCol = 1;
    expect(evalIn('Range("A1").Text', wb)).toBe("50%");
    expect(evalIn('Range("A1").Formula', wb)).toBe("=A2/2");
  });

  it("clears with ClearContents", () => {
    const wb = bareWorkbook([[1, 2]]);
    run('Range("A1:B1").ClearContents', wb);
    expect(at(wb, 1, 1)).toBe("");
    expect(at(wb, 1, 2)).toBe("");
  });
});

describe("Range navigation", () => {
  const wb = (): Workbook => bareWorkbook([[1, 2, 3], [4, 5, 6], [7, 8, 9]]);

  it("counts rows, columns and cells", () => {
    expect(evalIn('Range("A1:C3").Count', wb())).toBe(9);
    expect(evalIn('Range("A1:C3").Rows.Count', wb())).toBe(3);
    expect(evalIn('Range("A1:C3").Columns.Count', wb())).toBe(3);
  });

  it("reports Row, Column and Address", () => {
    expect(evalIn('Range("B2").Row', wb())).toBe(2);
    expect(evalIn('Range("B2").Column', wb())).toBe(2);
    expect(evalIn('Range("B2:C3").Address', wb())).toBe("$B$2:$C$3");
    expect(evalIn('Range("B2").Address(False, False)', wb())).toBe("B2");
  });

  it("offsets and resizes", () => {
    expect(evalIn('Range("A1").Offset(1, 1).Value', wb())).toBe(5);
    expect(evalIn('Range("A1").Resize(2, 2).Address', wb())).toBe("$A$1:$B$2");
  });

  it("indexes Cells within the range, and Cells(n) walks across then down", () => {
    expect(evalIn('Range("A1:C3").Cells(2, 3).Value', wb())).toBe(6);
    expect(evalIn('Range("A1:C3")(4).Value', wb())).toBe(4); // the 4th cell is A2
  });

  it("takes a row or a column out of the range", () => {
    expect(evalIn('Range("A1:C3").Rows(2).Address', wb())).toBe("$A$2:$C$2");
    expect(evalIn('Range("A1:C3").Columns(3).Address', wb())).toBe("$C$1:$C$3");
  });

  it("walks the cells with For Each", () => {
    const wbx = wb();
    const m = parseModule('Function T()\nDim c, s\ns = 0\nFor Each c In Range("A1:B2")\ns = s + c.Value\nNext c\nT = s\nEnd Function');
    const v = new VbaInterpreter(m, { globals: excelGlobals({ wb: wbx, activeSheet: 0 }) }).run("T").value;
    expect(v).toBe(1 + 2 + 4 + 5);
  });

  it("accepts a defined name where a reference is wanted", () => {
    const wbx = wb();
    wbx.definedNames = new Map([["Total", "Sheet1!B2"]]);
    expect(evalIn('Range("Total").Value', wbx)).toBe(5);
  });

  it("stops on a reference it cannot read rather than guessing", () => {
    expect(() => evalIn('Range("not a ref").Value', wb())).toThrow(/not a reference/);
  });
});

describe("Worksheets", () => {
  const two = (): Workbook => {
    const wb = bareWorkbook([[1]]);
    wb.sheets.push({ name: "Data", cells: new Map(), maxRow: 0, maxCol: 0 });
    return wb;
  };

  it("indexes by position and by name", () => {
    expect(evalIn("Worksheets(2).Name", two())).toBe("Data");
    expect(evalIn('Worksheets("Data").Index', two())).toBe(2);
    expect(evalIn("Worksheets.Count", two())).toBe(2);
  });

  it("says which sheet is missing", () => {
    expect(() => evalIn('Worksheets("Nope").Name', two())).toThrow(/no sheet called Nope/);
  });

  it("writes to another sheet through Worksheets", () => {
    const wb = two();
    run('Worksheets("Data").Range("A1").Value = 5', wb);
    expect(getCell(wb.sheets[1]!, 1, 1)?.value).toBe("5");
  });

  it("renames a sheet", () => {
    const wb = two();
    run('Worksheets(2).Name = "Renamed"', wb);
    expect(wb.sheets[1]!.name).toBe("Renamed");
  });

  it("moves ActiveSheet with Activate, and bare Range follows it", () => {
    const wb = two();
    const host: ExcelHost = { wb, activeSheet: 0 };
    const m = parseModule('Sub T()\nWorksheets(2).Activate\nRange("A1").Value = 9\nEnd Sub');
    new VbaInterpreter(m, { globals: excelGlobals(host) }).run("T");
    expect(host.activeSheet).toBe(1);
    expect(getCell(wb.sheets[1]!, 1, 1)?.value).toBe("9");
  });

  it("walks the sheets with For Each", () => {
    const wb = two();
    const m = parseModule('Function T()\nDim s, n\nn = ""\nFor Each s In Worksheets\nn = n & s.Name & ";"\nNext s\nT = n\nEnd Function');
    expect(new VbaInterpreter(m, { globals: excelGlobals({ wb, activeSheet: 0 }) }).run("T").value)
      .toBe("Sheet1;Data;");
  });

  it("reports UsedRange over what the sheet actually holds", () => {
    expect(evalIn("ActiveSheet.UsedRange.Address", bareWorkbook([[1, 2], [3, 4]]))).toBe("$A$1:$B$2");
  });
});

describe("Selection and ActiveCell", () => {
  it("read and write through whatever is selected", () => {
    const wb = bareWorkbook([[1, 2, 3]]);
    const host = { selection: { r1: 1, c1: 2, r2: 1, c2: 3 }, activeCell: { r: 1, c: 2 } };
    expect(evalIn("Selection.Address", wb, host)).toBe("$B$1:$C$1");
    expect(evalIn("ActiveCell.Value", wb, host)).toBe(2);
    run("Selection.Value = 0", wb, host);
    expect([at(wb, 1, 2), at(wb, 1, 3)]).toEqual(["0", "0"]);
  });

  it("moves the selection with Select", () => {
    const wb = bareWorkbook([[1]]);
    const host: ExcelHost = { wb, activeSheet: 0 };
    const m = parseModule('Sub T()\nRange("C3").Select\nEnd Sub');
    new VbaInterpreter(m, { globals: excelGlobals(host) }).run("T");
    expect(host.selection).toEqual({ r1: 3, c1: 3, r2: 3, c2: 3 });
    expect(host.activeCell).toEqual({ r: 3, c: 3 });
  });
});

describe("hiding rows and columns", () => {
  it("hides the rows an EntireRow covers", () => {
    const wb = bareWorkbook([[1], [2], [3]]);
    run('Range("A2").EntireRow.Hidden = True', wb);
    expect([...wb.sheets[0]!.hiddenRows ?? []]).toEqual([2]);
    run('Range("A2").EntireRow.Hidden = False', wb);
    expect([...wb.sheets[0]!.hiddenRows ?? []]).toEqual([]);
  });

  it("hides columns through EntireColumn", () => {
    const wb = bareWorkbook([[1, 2, 3]]);
    run('Range("B1").EntireColumn.Hidden = True', wb);
    expect([...wb.sheets[0]!.hiddenCols ?? []]).toEqual([2]);
  });
});

describe("SpecialCells", () => {
  it("finds the blanks in a range and keeps them as separate areas", () => {
    const wb = bareWorkbook([[1], [""], [3], [""], [""]]);
    // Rows 2, 4 and 5 are blank; 4 and 5 are contiguous but sit in different rows, so three areas.
    expect(evalIn('Range("A1:A5").SpecialCells(xlCellTypeBlanks).Areas.Count', wb)).toBe(3);
    expect(evalIn('Range("A1:A5").SpecialCells(xlCellTypeBlanks).Address', wb)).toBe("$A$2,$A$4,$A$5");
  });

  it("hides every blank row in one statement, which is what real macros do with it", () => {
    const wb = bareWorkbook([[1], [""], [3], [""]]);
    run('Range("A1:A4").SpecialCells(xlCellTypeBlanks).EntireRow.Hidden = True', wb);
    expect([...wb.sheets[0]!.hiddenRows ?? []].sort((a, b) => a - b)).toEqual([2, 4]);
  });

  it("raises when nothing matches, as Excel does", () => {
    const wb = bareWorkbook([[1], [2]]);
    expect(() => evalIn('Range("A1:A2").SpecialCells(xlCellTypeBlanks).Address', wb)).toThrow(/no cells were found/);
  });

  it("refuses the kinds it cannot answer honestly", () => {
    const wb = bareWorkbook([[1]]);
    expect(() => evalIn('Range("A1").SpecialCells(xlCellTypeLastCell).Address', wb))
      .toThrow(/SpecialCells\(11\) is not supported/);
  });
});

describe("Application", () => {
  it("accepts the UI switches a macro flips, and reads them back", () => {
    const wb = bareWorkbook();
    const m = parseModule("Function T()\nApplication.ScreenUpdating = False\nT = Application.ScreenUpdating\nEnd Function");
    expect(new VbaInterpreter(m, { globals: excelGlobals({ wb, activeSheet: 0 }) }).run("T").value).toBe(false);
  });

  it("runs WorksheetFunction through sheetedit's own formula engine", () => {
    const wb = bareWorkbook([[1], [2], [3]]);
    expect(evalIn('Application.WorksheetFunction.Sum(Range("A1:A3"))', wb)).toBe(6);
    expect(evalIn('Application.WorksheetFunction.Max(Range("A1:A3"))', wb)).toBe(3);
    expect(evalIn("Application.WorksheetFunction.Round(2.567, 1)", wb)).toBe(2.6);
  });

  it("refuses everything that would reach outside the page", () => {
    const wb = bareWorkbook();
    for (const call of ["Application.Quit", 'Application.Shell("x")', 'Application.CreateObject("Scripting.FileSystemObject")']) {
      expect(() => run(call, wb), call).toThrow(/reaches outside the page/);
    }
  });

  it("refuses to save on the macro's behalf", () => {
    const wb = bareWorkbook();
    expect(() => run("ActiveWorkbook.Save", wb)).toThrow(/save from the toolbar/);
  });
});

describe("protection", () => {
  it("stops a write to a locked cell on a protected sheet", () => {
    const wb = bareWorkbook([[1]]);
    wb.sheets[0]!.protection = { sheet: true };
    expect(() => run('Range("A1").Value = 2', wb)).toThrow(/locked on a protected sheet/);
    expect(at(wb, 1, 1)).toBe("1");
  });

  it("lets the macro unprotect first, which is what they all do", () => {
    const wb = bareWorkbook([[1]]);
    wb.sheets[0]!.protection = { sheet: true };
    run('ActiveSheet.Unprotect\nRange("A1").Value = 2\nActiveSheet.Protect', wb);
    expect(at(wb, 1, 1)).toBe("2");
    expect(wb.sheets[0]!.protection?.sheet).toBe(true);
  });
});

describe("what is not modelled", () => {
  it("names the member instead of evaluating to Empty", () => {
    const wb = bareWorkbook([[1]]);
    const cases: [string, RegExp][] = [
      ['Range("A1").PivotTable', /Range.PivotTable is not supported/],
      ['Range("A1").Sort', /Range.Sort is not supported by sheetedit yet/],
      ['Range("A1").FormulaR1C1 = "=RC[-1]"', /FormulaR1C1 is not supported/],
      ["ActiveSheet.Visible", /Worksheet.Visible is not supported by sheetedit yet/],
      ["Worksheets.Add", /Worksheets.Add is not supported by sheetedit yet/],
      ['Range("A1").Font.Shadow = True', /Font.Shadow is not supported/],
    ];
    for (const [src, message] of cases) expect(() => run(src, wb), src).toThrow(message);
  });

  it("stops on an xl constant it does not carry", () => {
    expect(() => evalIn("xlSomeConstantWeNeverHeardOf", bareWorkbook())).toThrow(/is not defined/);
  });
});

// --- styling, which needs a real package behind it ------------------------------

/** A minimal but complete .xlsx, so a style write goes through the real style pool. */
function makeXlsx(): Uint8Array {
  const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const rel = "http://schemas.openxmlformats.org/package/2006/relationships";
  const od = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
      + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + `<Default Extension="xml" ContentType="application/xml"/>`
      + `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`
      + `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
      + `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${od}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${ns}" xmlns:r="${od}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="${rel}">`
      + `<Relationship Id="rId1" Type="${od}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${od}/styles" Target="styles.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`),
    "xl/styles.xml": strToU8(
      `<styleSheet xmlns="${ns}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>`
      + `<fills count="1"><fill><patternFill patternType="none"/></patternFill></fill></fills>`
      + `<borders count="1"><border/></borders>`
      + `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`
      + `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
    ),
  });
}

describe("Font and Interior", () => {
  const xlsx = (): Workbook => readWorkbook(makeXlsx(), "book.xlsx");

  it("sets bold through the format's own style writer", () => {
    const wb = xlsx();
    run('Range("A1").Font.Bold = True', wb);
    expect(getCell(wb.sheets[0]!, 1, 1)?.cellStyle?.bold).toBe(true);
    expect(wb.stylesDirty).toBe(true);
  });

  it("reads a colour back as the BGR long Excel uses, not RGB", () => {
    const wb = xlsx();
    // &HFF0000 is blue in Excel's BGR, so the CSS colour must come out #0000ff.
    run("Range(\"A1\").Interior.Color = &HFF0000", wb);
    expect(getCell(wb.sheets[0]!, 1, 1)?.cellStyle?.bg?.toLowerCase()).toBe("#0000ff");
    expect(evalIn('Range("A1").Interior.Color', wb)).toBe(0xff0000);
  });

  it("round-trips a font colour", () => {
    const wb = xlsx();
    run("Range(\"A1\").Font.Color = &H0000FF", wb); // red
    expect(getCell(wb.sheets[0]!, 1, 1)?.cellStyle?.color?.toLowerCase()).toBe("#ff0000");
    expect(evalIn('Range("A1").Font.Color', wb)).toBe(0x0000ff);
  });
});

describe("a real macro from the fixtures", () => {
  it("runs Modul1 end to end against the real object model", async () => {
    const { readFileSync } = await import("node:fs");
    const { unzipSync } = await import("fflate");
    const { readVbaProject } = await import("./vba");
    const bin = unzipSync(new Uint8Array(readFileSync("src/fixtures/macros-cp1252.xlsm")))["xl/vbaProject.bin"]!;
    const mod = readVbaProject(bin)!.modules.find((m) => m.name === "Modul1")!;

    // The macro unhides every blank row in five bands of a sheet called Miete, then protects it.
    const wb = bareWorkbook();
    wb.sheets[0]!.name = "Miete";
    for (const r of [5, 6, 10, 50, 104, 182, 254]) {
      wb.sheets[0]!.cells.set(`${r}:1`, { row: r, col: 1, value: `x${r}`, kind: "s" });
      wb.sheets[0]!.maxRow = Math.max(wb.sheets[0]!.maxRow, r);
    }
    wb.sheets[0]!.maxCol = 1;
    wb.sheets[0]!.hiddenRows = new Set([7, 8, 9]);

    new VbaInterpreter(parseModule(mod.source, mod.name), { globals: excelGlobals({ wb, activeSheet: 0 }) })
      .run("Plus1_Klicken");

    // The rows it was asked to unhide are visible again, and the sheet ends up protected.
    expect([...wb.sheets[0]!.hiddenRows]).toEqual([]);
    expect(wb.sheets[0]!.protection?.sheet).toBe(true);
  });
});

describe("the write hook", () => {
  it("reports every cell about to change, before it changes", () => {
    const wb = bareWorkbook([[1, 2]]);
    const seen: string[] = [];
    run('Range("A1:B1").Value = 9', wb, {
      onBeforeWrite: (s, r, c) => seen.push(`${s}:${r}:${c}=${getCell(wb.sheets[0]!, r, c)?.value ?? ""}`),
    });
    // The old value is still there when the hook fires, which is what makes one undo step possible.
    expect(seen).toEqual(["0:1:1=1", "0:1:2=2"]);
  });
});
