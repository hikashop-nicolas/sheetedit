import { describe, expect, it } from "vitest";
import { getCell, type Workbook } from "./model";
import { runnableSubs, runWorkbookMacro } from "./vba-macro";

// Stage 4 of _plans/VBA_PLAN.md. The run itself is one undo step, and a run that stops part-way is
// rolled back before the caller ever sees it: a half-run macro that the user then saves is the one
// failure mode this whole feature has to rule out.

function workbook(rows: (string | number)[][] = [], sheets = 1): Workbook {
  const wb: Workbook = { kind: "xlsx", sheets: [], files: {} };
  for (let i = 0; i < sheets; i++) wb.sheets.push({ name: i ? `Sheet${i + 1}` : "Sheet1", cells: new Map(), maxRow: 0, maxCol: 0 });
  const sheet = wb.sheets[0]!;
  rows.forEach((row, r) => row.forEach((v, c) => {
    sheet.cells.set(`${r + 1}:${c + 1}`, { row: r + 1, col: c + 1, value: String(v), kind: typeof v === "number" ? "n" : "s" });
    sheet.maxRow = Math.max(sheet.maxRow, r + 1);
    sheet.maxCol = Math.max(sheet.maxCol, c + 1);
  }));
  return wb;
}

const at = (wb: Workbook, r: number, c: number, s = 0): string | undefined => getCell(wb.sheets[s]!, r, c)?.value;

describe("choosing what to run", () => {
  it("offers the public Subs that need no argument", () => {
    const src = [
      "Public Sub Go()", "End Sub",
      "Private Sub Helper()", "End Sub",
      "Sub NeedsArgs(n As Long)", "End Sub",
      "Function Calc()", "End Function",
      "Sub AlsoGo(Optional n As Long = 1)", "End Sub",
    ].join("\n");
    expect(runnableSubs(src)).toEqual(["Go", "AlsoGo"]);
  });

  it("offers nothing from a module it cannot parse, rather than throwing at the viewer", () => {
    expect(runnableSubs("Sub Broken(\nEnd Sub")).toEqual([]);
  });
});

describe("a successful run", () => {
  it("writes the cells and hands back one undo/redo pair", () => {
    const wb = workbook([[1, 2]]);
    const res = runWorkbookMacro(wb, 'Sub Go()\nRange("A1").Value = 10\nRange("B1").Value = 20\nEnd Sub', "M", "Go");
    expect(res.ok).toBe(true);
    expect([at(wb, 1, 1), at(wb, 1, 2)]).toEqual(["10", "20"]);

    res.undo!();
    expect([at(wb, 1, 1), at(wb, 1, 2)]).toEqual(["1", "2"]);
    res.redo!();
    expect([at(wb, 1, 1), at(wb, 1, 2)]).toEqual(["10", "20"]);
  });

  it("undoes a cell the macro created back out of existence", () => {
    const wb = workbook();
    const res = runWorkbookMacro(wb, 'Sub Go()\nRange("C3").Value = "new"\nEnd Sub', "M", "Go");
    expect(at(wb, 3, 3)).toBe("new");
    res.undo!();
    expect(at(wb, 3, 3)).toBe("");
  });

  it("collects one undo step across several sheets", () => {
    const wb = workbook([[1]], 2);
    const res = runWorkbookMacro(wb, 'Sub Go()\nRange("A1").Value = 9\nWorksheets(2).Range("A1").Value = 8\nEnd Sub', "M", "Go");
    expect([at(wb, 1, 1, 0), at(wb, 1, 1, 1)]).toEqual(["9", "8"]);
    res.undo!();
    expect([at(wb, 1, 1, 0), at(wb, 1, 1, 1)]).toEqual(["1", ""]);
  });

  it("undoes hidden rows, a rename and protection, which touch no cell at all", () => {
    const wb = workbook([[1], [2], [3]]);
    const res = runWorkbookMacro(wb, [
      "Sub Go()",
      '  Rows(2).Hidden = True',
      '  ActiveSheet.Name = "Renamed"',
      "  ActiveSheet.Protect",
      "End Sub",
    ].join("\n"), "M", "Go");
    expect(res.ok).toBe(true);
    expect([...wb.sheets[0]!.hiddenRows!]).toEqual([2]);
    expect(wb.sheets[0]!.name).toBe("Renamed");
    expect(wb.sheets[0]!.protection?.sheet).toBe(true);

    res.undo!();
    expect([...wb.sheets[0]!.hiddenRows!]).toEqual([]);
    expect(wb.sheets[0]!.name).toBe("Sheet1");
    expect(wb.sheets[0]!.protection).toBeUndefined();
  });

  it("offers no undo when the macro changed nothing", () => {
    const wb = workbook([[1]]);
    const res = runWorkbookMacro(wb, 'Sub Go()\nDebug.Print Range("A1").Value\nEnd Sub', "M", "Go");
    expect(res.ok).toBe(true);
    expect(res.messages).toEqual(["1"]);
    expect(res.undo).toBeUndefined();
  });

  it("reports where the macro left the view", () => {
    const wb = workbook([[1]], 2);
    const res = runWorkbookMacro(wb, 'Sub Go()\nWorksheets(2).Activate\nRange("B3").Select\nEnd Sub', "M", "Go");
    expect(res.activeSheet).toBe(1);
    expect(res.selection).toEqual({ r1: 3, c1: 2, r2: 3, c2: 2 });
  });
});

describe("a run that stops", () => {
  it("rolls back everything it had already written", () => {
    const wb = workbook([[1, 2]]);
    const res = runWorkbookMacro(wb, [
      "Sub Go()",
      '  Range("A1").Value = 99',
      '  Range("B1").Value = 98',
      "  Application.Shell \"whatever\"", // stops here
      '  Range("A1").Value = 1000',
      "End Sub",
    ].join("\n"), "Modul1", "Go");
    expect(res.ok).toBe(false);
    expect(res.error?.message).toMatch(/reaches outside the page/);
    expect(res.error?.module).toBe("Modul1");
    // Nothing survived, which is the whole point.
    expect([at(wb, 1, 1), at(wb, 1, 2)]).toEqual(["1", "2"]);
  });

  it("rolls back sheet state as well as cells", () => {
    const wb = workbook([[1], [2]]);
    const res = runWorkbookMacro(wb, 'Sub Go()\nRows(1).Hidden = True\nApplication.Quit\nEnd Sub', "M", "Go");
    expect(res.ok).toBe(false);
    expect([...wb.sheets[0]!.hiddenRows ?? []]).toEqual([]);
  });

  it("names the line the macro stopped on", () => {
    const wb = workbook();
    const res = runWorkbookMacro(wb, "Sub Go()\nDim x\nx = 1 / 0\nEnd Sub", "M", "Go");
    expect(res.ok).toBe(false);
    expect(res.error?.line).toBe(3);
  });

  it("reports a syntax error against the module without running anything", () => {
    const wb = workbook([[1]]);
    const res = runWorkbookMacro(wb, "Sub Go()\nIf Then\nEnd Sub", "M", "Go");
    expect(res.ok).toBe(false);
    expect(at(wb, 1, 1)).toBe("1");
  });

  it("stops a runaway loop and leaves the workbook as it was", () => {
    const wb = workbook([[1]]);
    const res = runWorkbookMacro(wb, 'Sub Go()\nDo\nRange("A1").Value = 2\nLoop\nEnd Sub', "M", "Go", { maxSteps: 500 });
    expect(res.ok).toBe(false);
    expect(res.error?.message).toMatch(/ran too long/);
    expect(at(wb, 1, 1)).toBe("1");
  });

  it("refuses to write to a protected sheet, and rolls back what came before", () => {
    const wb = workbook([[1, 2]]);
    wb.sheets[0]!.protection = { sheet: true };
    const res = runWorkbookMacro(wb, 'Sub Go()\nActiveSheet.Unprotect\nRange("A1").Value = 5\nActiveSheet.Protect\nRange("B1").Value = 6\nEnd Sub', "M", "Go");
    expect(res.ok).toBe(false);
    expect(res.error?.message).toMatch(/locked on a protected sheet/);
    expect([at(wb, 1, 1), at(wb, 1, 2)]).toEqual(["1", "2"]);
    expect(wb.sheets[0]!.protection?.sheet).toBe(true);
  });
});

describe("MsgBox", () => {
  it("comes back as text rather than blocking the tab", () => {
    const wb = workbook();
    const res = runWorkbookMacro(wb, 'Sub Go()\nMsgBox "all done"\nEnd Sub', "M", "Go");
    expect(res.ok).toBe(true);
    expect(res.messages).toEqual(["all done"]);
  });
});
