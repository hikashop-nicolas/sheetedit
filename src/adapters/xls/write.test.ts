import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { getCell } from "../../core/model";
import { readWorkbook, setCellInput, writeWorkbook } from "../../core/workbook";
import { setXlsxCellNumFmt, setXlsxCellStyle, setXlsxComment, setXlsxCondFormat, setXlsxDataValidation, setXlsxHyperlink, setXlsxMerge } from "../xlsx";
import { decodeFormula, type FormulaContext } from "./formula-read";
import { compileFormula, type CompileEnv } from "./formula-write";
import { writeXls } from "./write";

const fixture = (): Uint8Array => new Uint8Array(readFileSync(join(__dirname, "../../fixtures/compound.xls")));

const ctx: FormulaContext = {
  biff: 8, row: 0, col: 0,
  xti: (i) => (["Sheet1!", "Sheet2!", "'My Sheet'!", "Sheet1:Sheet2!"][i] ?? null),
  tabs: () => null,
  name: (i) => ["Total", "Rate"][i - 1] ?? "#NAME?",
  nameX: () => "#NAME?",
};
const env: CompileEnv = {
  sheet: 0,
  sheetIndex: (n) => ["sheet1", "sheet2", "my sheet"].indexOf(n.toLowerCase()),
  xti: (a, b) => (a === 0 && b === 1 ? 3 : a),
  nameIndex: (n) => ["total", "rate"].indexOf(n.toLowerCase()) + 1,
  addinFunc: () => ({ ixti: 0, index: 1 }),
  funcName: () => 1,
  externName: () => ({ ixti: 0, index: 1 }),
  sheetName: () => ({ ixti: 0, index: 1 }),
};

describe("formula tokens", () => {
  const cases = [
    "SUM(A1:B2)", "A1+B$2*$C3-$D$4", "IF(A1>0,\"yes\",\"no\")", "-A1^2%", "(A1+B1)*C1", "Sheet2!A1+'My Sheet'!$B$2:C3",
    "SUM(Sheet1:Sheet2!A1)", "VLOOKUP(A1,Sheet2!A:B,2,FALSE)", "SUM(1:3)", "Total*Rate", "PI()", "ROUND(A1,2)", "A1&\"x\"\"y\"",
    "SUM({1,2;3,4})", "INDEX(A1:C3,2,2)", "IF(A1,,1)", "#N/A", "1.5E+25", "COUNTIF(A1:A9,\">3\")", "SUM(A1:B2 B2:C3)", "(A1,B1)",
  ];
  for (const f of cases) it(`compiles and decodes ${f}`, () => {
    const { rgce, rgcb } = compileFormula(f, env);
    expect(decodeFormula(rgce, rgcb, ctx)).toBe(f);
  });

  it("refuses what BIFF8 cannot hold", () => {
    expect(compileFormula("A70000", env).rgce[0]).toBe(0x43); // past IV65536 it is a name
    expect(() => compileFormula(`"${"x".repeat(300)}"`, env)).toThrow();
    expect(() => compileFormula("Table1[Col]", env)).toThrow();
  });
});

describe("xls write", () => {
  it("writes every edit made through the xlsx paths back as BIFF8", () => {
    const wb = readWorkbook(fixture());
    const s = wb.sheets[0]!;
    setCellInput(s, 14, 1, "Hello été 中");
    setCellInput(s, 14, 2, "=B11*2");
    setCellInput(s, 15, 2, "0.25");
    setXlsxCellNumFmt(wb, s, getCell(s, 15, 2)!, "0.00%");
    setXlsxCellStyle(wb, s, getCell(s, 14, 1)!, { bold: true, italic: true, color: "#ff0000", bg: "#ffff00", align: "center", border: true, fontSize: 14 });
    setXlsxComment(wb, s, 14, 1, "A note");
    setXlsxHyperlink(wb, s, 15, 1, { href: "https://example.com/", tip: "Example" });
    setXlsxDataValidation(s, [{ r1: 16, c1: 1, r2: 18, c2: 1 }], { type: "list", values: ["a", "b", "c"], allowBlank: true });
    setXlsxCondFormat(wb, s, [{ r1: 3, c1: 2, r2: 5, c2: 2 }], { kind: "cellIs", operator: "greaterThan", value: "200", fill: "#ffc7ce" });
    setXlsxMerge(s, 20, 1, 20, 3, true);
    s.freeze = { rows: 1, cols: 1 };
    s.protection = { sheet: true, locks: { formatCells: false }, password: { legacy: "CC3D" } };
    s.printSetup = { ...s.printSetup, orientation: "landscape", paperSize: 9, printArea: [{ r1: 1, c1: 1, r2: 11, c2: 2 }], titleRows: { from: 1, to: 1 } };
    (wb.definedNames ??= new Map()).set("Amounts", "Data!$B$3:$B$11");
    const out = writeWorkbook(wb);
    expect(out.subarray(0, 4)).toEqual(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]));
    const back = readWorkbook(out);
    const b = back.sheets[0]!;
    expect(getCell(b, 14, 1)?.value).toBe("Hello été 中");
    expect(getCell(b, 14, 2)?.formula).toBe("B11*2");
    expect(getCell(b, 14, 2)?.value).toBe("5600");
    expect(getCell(b, 15, 2)?.numFmt).toBe("0.00%");
    expect(getCell(b, 14, 1)?.cellStyle).toMatchObject({ bold: true, italic: true, color: "#ff0000", bg: "#ffff00", align: "center", fontSize: 14 });
    expect(getCell(b, 14, 1)?.cellStyle?.borders?.top).toBeTruthy();
    expect(getCell(b, 14, 1)?.comments?.[0]?.text).toBe("A note");
    expect(getCell(b, 15, 1)?.link).toMatchObject({ href: "https://example.com/", tip: "Example" });
    expect(b.validations?.[0]).toMatchObject({ type: "list", values: ["a", "b", "c"] });
    expect(b.condFormats?.[0]?.rules[0]).toMatchObject({ type: "cellIs", operator: "greaterThan", formulas: ["200"] });
    expect(b.condFormats?.[0]?.rules[0]?.dxf?.bg).toBe("#ffc7ce");
    expect(b.merges).toContainEqual({ r1: 20, c1: 1, r2: 20, c2: 3 });
    expect(b.freeze).toEqual({ rows: 1, cols: 1 });
    expect(b.protection).toMatchObject({ sheet: true, locks: { formatCells: false }, password: { legacy: "CC3D" } });
    expect(b.printSetup).toMatchObject({ orientation: "landscape", paperSize: 9, printArea: [{ r1: 1, c1: 1, r2: 11, c2: 2 }], titleRows: { from: 1, to: 1 } });
    expect(back.definedNames?.get("Amounts")).toBe("Data!$B$3:$B$11");
    expect(b.rowOutline).toEqual(s.rowOutline);
    // A second trip changes nothing.
    const again = readWorkbook(writeXls(back));
    expect([...again.sheets[0]!.cells.values()].filter((c) => c.value !== "" || c.formula).map((c) => `${c.row}:${c.col}=${c.value}|${c.formula ?? ""}`).sort())
      .toEqual([...b.cells.values()].filter((c) => c.value !== "" || c.formula).map((c) => `${c.row}:${c.col}=${c.value}|${c.formula ?? ""}`).sort());
  });

  it("splits a large shared string table across CONTINUE records", () => {
    const wb = readWorkbook(fixture());
    const s = wb.sheets[0]!;
    for (let r = 30; r < 2030; r++) setCellInput(s, r, 1, `row ${r} ` + "é".repeat(r % 40) + "x".repeat(r % 13));
    setCellInput(s, 2040, 1, "z".repeat(20000));
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    for (let r = 30; r < 2030; r += 97) expect(getCell(back, r, 1)?.value).toBe(`row ${r} ` + "é".repeat(r % 40) + "x".repeat(r % 13));
    expect(getCell(back, 2040, 1)?.value).toBe("z".repeat(20000));
  });
});
