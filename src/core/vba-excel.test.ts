import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../index";
import { getCell, type Workbook } from "./model";
import { parseModule } from "vbalang";
import { VbaInterpreter, type RunResult } from "vbalang";
import { EMPTY, VbaArray, type VbaValue } from "vbalang";
import { excelGlobals, type ExcelHost } from "./vba-excel";

// Stage 3 of _plans/done/VBA_PLAN.md. The object model a macro talks to, mapped onto sheetedit's own
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
      ["ActiveSheet.ExportAsFixedFormat", /cannot export a PDF from a macro/],
      ['Range("A1").FormulaR1C1 = "=RC[-1]"', /FormulaR1C1 is not supported/],
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

describe("Range.Sort", () => {
  const data = (): Workbook => bareWorkbook([["Name", "Score"], ["Carol", 3], ["alice", 10], ["Bob", 7]]);

  it("sorts by a key column, keeping each row's cells together", () => {
    const wb = data();
    run('Range("A1:B4").Sort Key1:=Range("A1"), Header:=xlYes', wb);
    expect([at(wb, 2, 1), at(wb, 3, 1), at(wb, 4, 1)]).toEqual(["alice", "Bob", "Carol"]);
    expect([at(wb, 2, 2), at(wb, 3, 2), at(wb, 4, 2)]).toEqual(["10", "7", "3"]);
  });

  it("sorts descending on xlDescending", () => {
    const wb = data();
    run('Range("A1:B4").Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlYes', wb);
    expect([at(wb, 2, 2), at(wb, 3, 2), at(wb, 4, 2)]).toEqual(["10", "7", "3"]);
  });

  it("treats the first row as data unless Header says otherwise", () => {
    const wb = data();
    run('Range("A1:B4").Sort Key1:=Range("A1")', wb);
    // "Name" sorts among the values, which is what xlNo (the default) means.
    expect(at(wb, 4, 1)).toBe("Name");
  });

  it("takes a second key for ties", () => {
    const wb = bareWorkbook([["b", 2], ["a", 2], ["a", 1]]);
    run('Range("A1:B3").Sort Key1:=Range("A1"), Key2:=Range("B1")', wb);
    expect([at(wb, 1, 1), at(wb, 1, 2)]).toEqual(["a", "1"]);
    expect([at(wb, 2, 1), at(wb, 2, 2)]).toEqual(["a", "2"]);
  });

  it("puts blanks last in both directions", () => {
    const wb = bareWorkbook([["b"], [""], ["a"]]);
    run('Range("A1:A3").Sort Key1:=Range("A1")', wb);
    expect([at(wb, 1, 1), at(wb, 2, 1), at(wb, 3, 1)]).toEqual(["a", "b", ""]);
    run('Range("A1:A3").Sort Key1:=Range("A1"), Order1:=xlDescending', wb);
    expect([at(wb, 1, 1), at(wb, 2, 1), at(wb, 3, 1)]).toEqual(["b", "a", ""]);
  });

  it("stops when no key was given", () => {
    expect(() => run('Range("A1:B4").Sort', data())).toThrow(/needs a key column/);
  });
});

describe("Range.AutoFilter", () => {
  const data = (): Workbook => bareWorkbook([["Fruit", "Qty"], ["apple", 5], ["pear", 2], ["apple", 9]]);

  it("turns the filter on over the range, and off again", () => {
    const wb = data();
    run('Range("A1:B4").AutoFilter', wb);
    expect(wb.sheets[0]!.autoFilter).toEqual({ r1: 1, c1: 1, r2: 4, c2: 2 });
    run('Range("A1:B4").AutoFilter', wb);
    expect(wb.sheets[0]!.autoFilter).toBeUndefined();
  });

  it("hides the rows a criterion excludes", () => {
    const wb = data();
    run('Range("A1:B4").AutoFilter Field:=1, Criteria1:="apple"', wb);
    expect([...wb.sheets[0]!.filterHidden ?? []]).toEqual([3]); // the pear row
  });

  it("understands the comparison criteria", () => {
    const wb = data();
    run('Range("A1:B4").AutoFilter Field:=2, Criteria1:=">4"', wb);
    expect([...wb.sheets[0]!.filterHidden ?? []]).toEqual([3]);
  });

  it("understands a wildcard", () => {
    const wb = data();
    run('Range("A1:B4").AutoFilter Field:=1, Criteria1:="=p*"', wb);
    expect([...wb.sheets[0]!.filterHidden ?? []].sort((a, b) => a - b)).toEqual([2, 4]);
  });

  it("shows everything again when the criterion is dropped", () => {
    const wb = data();
    run('Range("A1:B4").AutoFilter Field:=1, Criteria1:="apple"', wb);
    run('Range("A1:B4").AutoFilter Field:=1', wb);
    expect(wb.sheets[0]!.filterHidden).toBeUndefined();
  });

  it("marks the hidden rows in the file, not only in the model", async () => {
    // filterHidden lives beside hiddenRows and is not persisted by the outline writer, so the
    // row elements have to be flagged directly or a save loses the filtering.
    const { writeWorkbook } = await import("./workbook");
    const { strFromU8, unzipSync } = await import("fflate");
    const wb = readWorkbook(makeXlsx(), "book.xlsx");
    run('Range("A1:A3").Value = "x"\nRange("A2").Value = "y"\nRange("A1:A3").AutoFilter Field:=1, Criteria1:="x"', wb);
    expect([...wb.sheets[0]!.filterHidden ?? []]).toEqual([2]);
    const sheetXml = strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]!);
    expect(sheetXml).toMatch(/<row[^>]*r="2"[^>]*hidden="1"/);
  });

  it("stops on a field outside the range", () => {
    expect(() => run('Range("A1:B4").AutoFilter Field:=9, Criteria1:="x"', data())).toThrow(/outside the range/);
  });
});

describe("Range.Find", () => {
  const data = (): Workbook => bareWorkbook([["alpha", "beta"], ["gamma", "delta"]]);

  it("returns the first matching cell as a Range", () => {
    expect(evalIn('Range("A1:B2").Find("delta").Address', data())).toBe("$B$2");
  });

  it("matches a substring by default, and the whole cell on xlWhole", () => {
    expect(evalIn('Range("A1:B2").Find("amm").Address', data())).toBe("$A$2");
    expect(evalIn('Range("A1:B2").Find(What:="amm", LookAt:=xlWhole) Is Nothing', data())).toBe(true);
  });

  it("ignores case unless MatchCase says not to", () => {
    expect(evalIn('Range("A1:B2").Find("ALPHA").Address', data())).toBe("$A$1");
    expect(evalIn('Range("A1:B2").Find(What:="ALPHA", MatchCase:=True) Is Nothing', data())).toBe(true);
  });

  it("gives Nothing when there is no match, which is what macros test for", () => {
    expect(evalIn('Range("A1:B2").Find("nope") Is Nothing', data())).toBe(true);
  });
});

describe("Copy, Cut and Paste", () => {
  it("copies values and formulas to another place", () => {
    const wb = bareWorkbook([[1, 2], [3, 4]]);
    run('Range("A1:B2").Copy Range("D1")', wb);
    expect([at(wb, 1, 4), at(wb, 1, 5), at(wb, 2, 4), at(wb, 2, 5)]).toEqual(["1", "2", "3", "4"]);
    expect(at(wb, 1, 1)).toBe("1"); // the source is untouched
  });

  it("pastes in a second statement, as a macro usually writes it", () => {
    const wb = bareWorkbook([[7]]);
    run('Range("A1").Copy\nRange("C3").PasteSpecial', wb);
    expect(at(wb, 3, 3)).toBe("7");
  });

  it("carries a formula across as a formula", () => {
    const wb = bareWorkbook([[1], [2]]);
    run('Range("B1").Formula = "=A1+A2"\nRange("B1").Copy Range("C1")', wb);
    expect(getCell(wb.sheets[0]!, 1, 3)?.formula).toBe("A1+A2");
  });

  it("clears the source on a Cut, which is what makes it a move", () => {
    const wb = bareWorkbook([[1, 2]]);
    run('Range("A1:B1").Cut Range("A3")', wb);
    expect([at(wb, 3, 1), at(wb, 3, 2)]).toEqual(["1", "2"]);
    expect([at(wb, 1, 1), at(wb, 1, 2)]).toEqual(["", ""]);
  });

  it("copies between sheets", () => {
    const wb = bareWorkbook([[5]]);
    wb.sheets.push({ name: "Other", cells: new Map(), maxRow: 0, maxCol: 0 });
    run('Range("A1").Copy\nWorksheets("Other").Range("B2").PasteSpecial', wb);
    expect(getCell(wb.sheets[1]!, 2, 2)?.value).toBe("5");
  });

  it("says so when there is nothing on the clipboard", () => {
    expect(() => run('Range("A1").PasteSpecial', bareWorkbook())).toThrow(/nothing to paste/);
  });
});

describe("adding and deleting sheets", () => {
  it("adds a sheet and hands it back for chaining", () => {
    // Adding a sheet writes real parts, so this needs a real package rather than a bare model.
    const wb = readWorkbook(makeXlsx(), "book.xlsx");
    run('Worksheets.Add.Name = "Fresh"', wb);
    expect(wb.sheets.map((s) => s.name)).toContain("Fresh");
  });

  it("deletes a sheet", () => {
    const wb = bareWorkbook([[1]]);
    wb.sheets.push({ name: "Doomed", cells: new Map(), maxRow: 0, maxCol: 0 });
    run('Worksheets("Doomed").Delete', wb);
    expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1"]);
  });

  it("refuses to delete the only sheet", () => {
    expect(() => run("ActiveSheet.Delete", bareWorkbook([[1]]))).toThrow(/at least one sheet/);
  });
});

describe("Worksheet.Visible", () => {
  const two = (): Workbook => {
    const wb = readWorkbook(makeXlsx(), "book.xlsx");
    // makeXlsx has one sheet; a second is needed before one can be hidden.
    wb.sheets.push({ name: "Data", cells: new Map(), maxRow: 0, maxCol: 0 });
    return wb;
  };

  it("reads the xl constants Excel uses, not True/False", () => {
    const wb = two();
    expect(evalIn("Worksheets(2).Visible", wb)).toBe(-1); // xlSheetVisible
    wb.sheets[1]!.visibility = "hidden";
    expect(evalIn("Worksheets(2).Visible", wb)).toBe(0);
    wb.sheets[1]!.visibility = "veryHidden";
    expect(evalIn("Worksheets(2).Visible", wb)).toBe(2);
  });

  it("hides on False and on xlSheetHidden alike", () => {
    for (const src of ["Worksheets(2).Visible = False", "Worksheets(2).Visible = xlSheetHidden"]) {
      const wb = two();
      run(src, wb);
      expect(wb.sheets[1]!.visibility, src).toBe("hidden");
    }
  });

  it("sets very hidden, which is the state only a macro can reach", () => {
    const wb = two();
    run("Worksheets(2).Visible = xlSheetVeryHidden", wb);
    expect(wb.sheets[1]!.visibility).toBe("veryHidden");
  });

  it("shows a hidden sheet again", () => {
    const wb = two();
    wb.sheets[1]!.visibility = "hidden";
    run("Worksheets(2).Visible = True", wb);
    expect(wb.sheets[1]!.visibility).toBeUndefined();
  });

  it("stops rather than leaving a workbook with no reachable sheet", () => {
    const wb = two();
    wb.sheets[1]!.visibility = "hidden";
    expect(() => run("Worksheets(1).Visible = False", wb)).toThrow(/at least one visible sheet/);
  });
});

describe("Worksheet.Move and PrintOut", () => {
  const three = (): Workbook => {
    const wb = bareWorkbook([[1]]);
    wb.sheets.push({ name: "B", cells: new Map(), maxRow: 0, maxCol: 0 });
    wb.sheets.push({ name: "C", cells: new Map(), maxRow: 0, maxCol: 0 });
    return wb;
  };

  it("moves a sheet before another", () => {
    const wb = three();
    run('Worksheets("C").Move Before:=Worksheets(1)', wb);
    expect(wb.sheets.map((s) => s.name)).toEqual(["C", "Sheet1", "B"]);
  });

  it("moves a sheet after another", () => {
    const wb = three();
    run('Worksheets(1).Move After:=Worksheets("C")', wb);
    expect(wb.sheets.map((s) => s.name)).toEqual(["B", "C", "Sheet1"]);
  });

  it("stops when told neither Before nor After", () => {
    expect(() => run('Worksheets("B").Move', three())).toThrow(/Before or After/);
  });

  it("prints through the host, and says so when there is no host to print with", () => {
    const printed: number[] = [];
    const wb = three();
    run("Worksheets(2).PrintOut", wb, { print: (i) => printed.push(i) });
    expect(printed).toEqual([1]);
    expect(() => run("Worksheets(2).PrintOut", three())).toThrow(/printing is not available/);
  });
});

describe("Range.Replace", () => {
  const data = (): Workbook => bareWorkbook([["one two", "TWO"], ["two", 2]]);

  it("rewrites every occurrence and reports that it changed something", () => {
    const wb = data();
    expect(evalIn('Range("A1:B2").Replace("two", "2")', wb)).toBe(true);
    expect([at(wb, 1, 1), at(wb, 1, 2), at(wb, 2, 1)]).toEqual(["one 2", "2", "2"]);
  });

  it("ignores case unless MatchCase says not to", () => {
    const wb = data();
    run('Range("A1:B2").Replace What:="two", Replacement:="x", MatchCase:=True', wb);
    expect([at(wb, 1, 1), at(wb, 1, 2)]).toEqual(["one x", "TWO"]);
  });

  it("matches the whole cell on xlWhole", () => {
    const wb = data();
    run('Range("A1:B2").Replace What:="two", Replacement:="x", LookAt:=xlWhole', wb);
    expect([at(wb, 1, 1), at(wb, 2, 1)]).toEqual(["one two", "x"]);
  });

  it("says nothing changed when nothing matched", () => {
    expect(evalIn('Range("A1:B2").Replace("zzz", "x")', data())).toBe(false);
  });

  it("replaces inside a formula, and it stays a formula", () => {
    const wb = bareWorkbook([[1], [2]]);
    run('Range("C1").Formula = "=A1+A2"\nRange("C1").Replace "A1", "A2"', wb);
    expect(getCell(wb.sheets[0]!, 1, 3)?.formula).toBe("A2+A2");
  });
});

describe("Range.RemoveDuplicates", () => {
  it("drops repeated rows and pulls the rest up", () => {
    const wb = bareWorkbook([["a", 1], ["b", 2], ["a", 1], ["c", 3]]);
    run('Range("A1:B4").RemoveDuplicates', wb);
    expect([at(wb, 1, 1), at(wb, 2, 1), at(wb, 3, 1), at(wb, 4, 1)]).toEqual(["a", "b", "c", ""]);
  });

  it("compares only the key columns it is given", () => {
    const wb = bareWorkbook([["a", 1], ["a", 2], ["b", 3]]);
    run('Range("A1:B3").RemoveDuplicates Columns:=1', wb);
    // Row 2 repeats "a" in the key column even though its second column differs.
    expect([at(wb, 1, 1), at(wb, 2, 1), at(wb, 3, 1)]).toEqual(["a", "b", ""]);
  });

  it("keeps the header row when told there is one", () => {
    const wb = bareWorkbook([["Name"], ["a"], ["a"]]);
    run('Range("A1:A3").RemoveDuplicates Columns:=1, Header:=xlYes', wb);
    expect([at(wb, 1, 1), at(wb, 2, 1), at(wb, 3, 1)]).toEqual(["Name", "a", ""]);
  });

  it("stops on a key column outside the range", () => {
    expect(() => run('Range("A1:B2").RemoveDuplicates Columns:=5', bareWorkbook([["a", 1], ["a", 1]])))
      .toThrow(/outside the range/);
  });
});

describe("Range.TextToColumns", () => {
  it("splits the first column on a delimiter, in place", () => {
    const wb = bareWorkbook([["a,b,c"], ["d,e,f"]]);
    run('Range("A1:A2").TextToColumns Comma:=True', wb);
    expect([at(wb, 1, 1), at(wb, 1, 2), at(wb, 1, 3)]).toEqual(["a", "b", "c"]);
    expect([at(wb, 2, 1), at(wb, 2, 2), at(wb, 2, 3)]).toEqual(["d", "e", "f"]);
  });

  it("writes to a Destination when given one", () => {
    const wb = bareWorkbook([["a;b"]]);
    run('Range("A1").TextToColumns Destination:=Range("C5"), Semicolon:=True', wb);
    expect([at(wb, 5, 3), at(wb, 5, 4)]).toEqual(["a", "b"]);
    expect(at(wb, 1, 1)).toBe("a;b"); // the source is left alone
  });

  it("takes several delimiters at once", () => {
    const wb = bareWorkbook([["a,b;c"]]);
    run('Range("A1").TextToColumns Comma:=True, Semicolon:=True', wb);
    expect([at(wb, 1, 1), at(wb, 1, 2), at(wb, 1, 3)]).toEqual(["a", "b", "c"]);
  });

  it("refuses fixed width rather than splitting somewhere of its own choosing", () => {
    expect(() => run('Range("A1").TextToColumns DataType:=xlFixedWidth', bareWorkbook([["ab"]])))
      .toThrow(/fixed widths/);
  });

  it("stops when no delimiter was named", () => {
    expect(() => run('Range("A1").TextToColumns', bareWorkbook([["a,b"]]))).toThrow(/at least one delimiter/);
  });
});

describe("Range.AdvancedFilter", () => {
  // A1:B4 is the data with a header row; D1:E3 is the criteria range, also with a header row.
  const data = (): Workbook => bareWorkbook([
    ["Fruit", "Qty", "", "Fruit", "Qty"],
    ["apple", 5, "", "apple", ""],
    ["pear", 2, "", "", ">4"],
    ["apple", 9],
  ]);

  it("hides the rows that fail the criteria, in place", () => {
    const wb = data();
    // Two criteria rows: Fruit = apple, OR Qty > 4. Row 3 (pear, 2) fails both.
    run('Range("A1:B4").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Range("D1:E3")', wb);
    expect([...wb.sheets[0]!.filterHidden ?? []]).toEqual([3]);
  });

  it("ANDs the cells of one criteria row", () => {
    const wb = bareWorkbook([
      ["Fruit", "Qty", "", "Fruit", "Qty"],
      ["apple", 5, "", "apple", ">6"],
      ["apple", 9],
      ["pear", 9],
    ]);
    // One row: Fruit = apple AND Qty > 6. Only row 3 qualifies.
    run('Range("A1:B4").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Range("D1:E2")', wb);
    expect([...wb.sheets[0]!.filterHidden ?? []].sort((a, b) => a - b)).toEqual([2, 4]);
  });

  it("copies the matches, header and all, to CopyToRange", () => {
    const wb = data();
    run('Range("A1:B4").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("D1:E3"), CopyToRange:=Range("A10")', wb);
    expect([at(wb, 10, 1), at(wb, 10, 2)]).toEqual(["Fruit", "Qty"]);
    expect([at(wb, 11, 1), at(wb, 12, 1)]).toEqual(["apple", "apple"]);
    expect(at(wb, 13, 1)).toBeUndefined(); // a copy writes its rows and stops; it clears no tail
    expect(wb.sheets[0]!.filterHidden).toBeUndefined(); // a copy hides nothing
  });

  it("drops repeats when asked for unique rows only", () => {
    const wb = bareWorkbook([["Fruit"], ["apple"], ["apple"], ["pear"]]);
    run('Range("A1:A4").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("C1"), Unique:=True', wb);
    expect([at(wb, 1, 3), at(wb, 2, 3), at(wb, 3, 3)]).toEqual(["Fruit", "apple", "pear"]);
    expect(at(wb, 4, 3)).toBeUndefined();
  });

  it("keeps every row when there is no criteria range at all", () => {
    const wb = data();
    run('Range("A1:B4").AdvancedFilter Action:=xlFilterInPlace', wb);
    expect(wb.sheets[0]!.filterHidden).toBeUndefined();
  });

  it("stops when the criteria name a column the range does not have", () => {
    const wb = bareWorkbook([["Fruit", "", "Colour"], ["apple", "", "red"]]);
    expect(() => run('Range("A1:A2").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Range("C1:C2")', wb))
      .toThrow(/does not have/);
  });

  it("stops when a copy has nowhere to go", () => {
    expect(() => run('Range("A1:B4").AdvancedFilter Action:=xlFilterCopy', data())).toThrow(/needs a CopyToRange/);
  });
});

describe("Range.Clear and ClearFormats", () => {
  const styled = (): Workbook => {
    const wb = readWorkbook(makeXlsx(), "book.xlsx");
    run('Range("A1").Value = 5\nRange("A1").Font.Bold = True\nRange("A1").Interior.Color = &HFF0000', wb);
    return wb;
  };

  it("ClearFormats takes the styling off and leaves the value", () => {
    const wb = styled();
    expect(getCell(wb.sheets[0]!, 1, 1)?.cellStyle?.bold).toBe(true);
    run('Range("A1").ClearFormats', wb);
    const cell = getCell(wb.sheets[0]!, 1, 1)!;
    expect(cell.cellStyle).toBeUndefined();
    expect(cell.style).toBeUndefined();
    expect(cell.value).toBe("5");
  });

  it("Clear takes both", () => {
    const wb = styled();
    run('Range("A1").Clear', wb);
    const cell = getCell(wb.sheets[0]!, 1, 1)!;
    expect(cell.cellStyle).toBeUndefined();
    expect(cell.value).toBe("");
  });

  it("drops the style index from the file, not only from the model", async () => {
    const { writeWorkbook } = await import("./workbook");
    const { strFromU8, unzipSync } = await import("fflate");
    const wb = styled();
    expect(strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]!)).toMatch(/<c r="A1"[^>]*\ss="/);
    run('Range("A1").ClearFormats', wb);
    expect(strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]!)).not.toMatch(/<c r="A1"[^>]*\ss="/);
  });
});

describe("Worksheet.Copy", () => {
  const book = (): Workbook => {
    const wb = readWorkbook(makeXlsx(), "book.xlsx");
    run('Range("A1").Value = 7\nRange("B2").Formula = "=A1*2"\nRange("A1").Font.Bold = True', wb);
    return wb;
  };

  it("duplicates the grid next to the original", () => {
    const wb = book();
    run("Worksheets(1).Copy After:=Worksheets(1)", wb);
    expect(wb.sheets).toHaveLength(2);
    const copy = wb.sheets[1]!;
    expect(getCell(copy, 1, 1)?.value).toBe("7");
    expect(getCell(copy, 2, 2)?.formula).toBe("A1*2");
    expect(getCell(copy, 1, 1)?.cellStyle?.bold).toBe(true);
  });

  it("gives the copy its own cells, so editing one leaves the other alone", () => {
    const wb = book();
    run("Worksheets(1).Copy After:=Worksheets(1)", wb);
    run('Worksheets(2).Range("A1").Value = 99', wb);
    expect(getCell(wb.sheets[0]!, 1, 1)?.value).toBe("7");
    expect(getCell(wb.sheets[1]!, 1, 1)?.value).toBe("99");
  });

  it("puts it before when told Before", () => {
    const wb = book();
    const original = wb.sheets[0]!.name;
    run("Worksheets(1).Copy Before:=Worksheets(1)", wb);
    expect(wb.sheets[1]!.name).toBe(original);
  });

  it("survives a save and re-read", async () => {
    const { writeWorkbook } = await import("./workbook");
    const wb = book();
    run("Worksheets(1).Copy After:=Worksheets(1)", wb);
    const back = readWorkbook(writeWorkbook(wb), "book.xlsx");
    expect(back.sheets).toHaveLength(2);
    expect(getCell(back.sheets[1]!, 1, 1)?.value).toBe("7");
  });

  it("refuses the no-argument form, which makes a new workbook in Excel", () => {
    expect(() => run("Worksheets(1).Copy", book())).toThrow(/cannot open a new workbook/);
  });
});

// OLEObjects: an ActiveX control is reachable from a macro, which is how the button on a real
// worksheet drives its combo box. Anything not modelled refuses by name, as everywhere else.
describe("OLEObjects", () => {
  const withCombo = (): Workbook => {
    const wb = bareWorkbook([["Mon"], ["Tue"], ["Wed"]]);
    wb.sheets[0]!.controls = [
      { kind: "dropdown", name: "ComboDay", activeX: true, activeXValue: "Tue", sourceRange: "A1:A3", linkedCell: "$C$1" },
      { kind: "button", name: "PlainButton" }, // a form control: not an OLE object
    ];
    return wb;
  };
  const items = (): string[] => ["Mon", "Tue", "Wed"];

  it("reaches a control by name and reads its list state", () => {
    const wb = withCombo();
    expect(evalIn(`ActiveSheet.OLEObjects("ComboDay").Object.ListCount`, wb, { controlItems: items })).toBe(3);
    expect(evalIn(`ActiveSheet.OLEObjects("ComboDay").Object.ListIndex`, wb, { controlItems: items })).toBe(1);
    expect(evalIn(`ActiveSheet.OLEObjects("ComboDay").Object.List(2)`, wb, { controlItems: items })).toBe("Wed");
    expect(evalIn(`ActiveSheet.OLEObjects.Count`, wb, { controlItems: items })).toBe(1); // the form control is not one
  });

  it("sets a control's value, which moves its linked cell too", () => {
    const wb = withCombo();
    let notified = 0;
    run(`ActiveSheet.OLEObjects("ComboDay").Object.Value = "Wed"`, wb, { controlItems: items, onControlChange: () => { notified++; } });
    expect(wb.sheets[0]!.controls![0]!.activeXValue).toBe("Wed");
    expect(at(wb, 1, 3)).toBe("Wed"); // the linked cell follows, as it does in Excel
    expect(notified).toBe(1);
  });

  it("runs the advance-the-combo pattern a real button uses", () => {
    const wb = withCombo();
    run(
      `With ActiveSheet.OLEObjects("ComboDay").Object\n` +
        `  If .ListIndex = .ListCount - 1 Then\n    .Value = .List(0)\n  Else\n    .Value = .List(.ListIndex + 1)\n  End If\n` +
        `End With`,
      wb,
      { controlItems: items },
    );
    expect(wb.sheets[0]!.controls![0]!.activeXValue).toBe("Wed");
  });

  it("refuses a member it does not model, and a name that is not there", () => {
    const wb = withCombo();
    expect(() => evalIn(`ActiveSheet.OLEObjects("ComboDay").Object.BackStyle`, wb, { controlItems: items })).toThrow(/BackStyle/);
    expect(() => evalIn(`ActiveSheet.OLEObjects("Nope").Object.Value`, wb, { controlItems: items })).toThrow(/Nope/);
  });
});
