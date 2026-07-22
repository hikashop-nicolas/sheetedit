import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook, getCell } from "../../index";
import { createWorksheet, uniqueSheetName } from "./sheet-create";
import { loadResultToNewSheet } from "./tables";
import type { MValue } from "mlang";

// A minimal but realistic .xlsx (one sheet, proper content-types/rels) so createWorksheet's
// registration is exercised end to end: create -> write -> re-read must surface the new sheet.
function makeXlsx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
    ),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

const num = (n: number): MValue => ({ kind: "number", value: n });
const txt = (s: string): MValue => ({ kind: "text", value: s });

describe("createWorksheet", () => {
  it("registers a new sheet that survives a write + re-read", () => {
    const wb = readWorkbook(makeXlsx());
    expect(wb.sheets).toHaveLength(1);
    const sheet = createWorksheet(wb, "Extra");
    expect(sheet.name).toBe("Extra");
    expect(sheet.path).toBe("xl/worksheets/sheet2.xml");
    // workbook.xml, its rels and content-types all gained the part.
    expect(new TextDecoder().decode(wb.files["xl/workbook.xml"])).toContain('name="Extra"');
    expect(new TextDecoder().decode(wb.files["xl/_rels/workbook.xml.rels"])).toContain("worksheets/sheet2.xml");
    expect(new TextDecoder().decode(wb.files["[Content_Types].xml"])).toContain("/xl/worksheets/sheet2.xml");

    const out = writeWorkbook(wb);
    const reread = readWorkbook(out);
    expect(reread.sheets.map((s) => s.name)).toEqual(["Sheet1", "Extra"]);
    // The re-read file is a valid zip with the new part.
    expect(Object.keys(unzipSync(out))).toContain("xl/worksheets/sheet2.xml");
  });

  it("uniqueSheetName avoids collisions and Excel's 31-char / illegal-char limits", () => {
    const wb = readWorkbook(makeXlsx());
    expect(uniqueSheetName(wb, "Sheet1")).toBe("Sheet1 (2)");
    expect(uniqueSheetName(wb, "a/b:c")).toBe("a b c");
    expect(uniqueSheetName(wb, "x".repeat(50)).length).toBeLessThanOrEqual(31);
  });

  it("loadResultToNewSheet writes header + data that re-read correctly", () => {
    const wb = readWorkbook(makeXlsx());
    const result: Extract<MValue, { kind: "table" }> = {
      kind: "table",
      columns: ["Product", "Qty"],
      rows: [[txt("Apples"), num(10)], [txt("Pears"), num(4)]],
    };
    const { sheetIndex, rows } = loadResultToNewSheet(wb, "Sales", result);
    expect(rows).toBe(2);
    const out = writeWorkbook(wb);
    const reread = readWorkbook(out);
    const s = reread.sheets[sheetIndex];
    expect(s.name).toBe("Sales");
    expect(getCell(s, 1, 1)?.value).toBe("Product");
    expect(getCell(s, 1, 2)?.value).toBe("Qty");
    expect(getCell(s, 2, 1)?.value).toBe("Apples");
    expect(getCell(s, 3, 2)?.value).toBe("4");
  });
});
