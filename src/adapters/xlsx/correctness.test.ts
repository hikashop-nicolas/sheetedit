import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, setCellInput, writeWorkbook } from "../../core/workbook";
import { getCell } from "../../core/model";

// Build a one-sheet book whose sharedStrings and A1..cells are supplied by the caller.
function book(sst: string, cells: string): ReturnType<typeof readWorkbook> {
  const xlsx = zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${cells}</row></sheetData></worksheet>`),
    "xl/sharedStrings.xml": strToU8(sst),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
  return readWorkbook(xlsx);
}

describe("edit correctness", () => {
  it("editing one shared-string cell does not change a cell that shares the string", () => {
    const wb = book(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="1"><si><t>hello</t></si></sst>`,
      `<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c>`,
    );
    setCellInput(wb.sheets[0], 1, 1, "world");
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.value).toBe("world");
    expect(getCell(re.sheets[0], 1, 2)?.value).toBe("hello");
  });

  it("editing a rich-text cell drops its stale per-run styling", () => {
    const wb = book(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>a</t></r><r><rPr><b/></rPr><t>B</t></r></si></sst>`,
      `<c r="A1" t="s"><v>0</v></c>`,
    );
    expect(getCell(wb.sheets[0], 1, 1)?.richRuns?.length).toBe(2);
    setCellInput(wb.sheets[0], 1, 1, "plain now");
    expect(getCell(wb.sheets[0], 1, 1)?.richRuns).toBeUndefined();
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.value).toBe("plain now");
    expect(getCell(re.sheets[0], 1, 1)?.richRuns).toBeUndefined();
  });
});
