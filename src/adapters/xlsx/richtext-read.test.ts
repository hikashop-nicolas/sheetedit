import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";

// A shared string with multiple styled runs must surface per-run styling in cell.richRuns.
function book(sst: string): ReturnType<typeof readWorkbook> {
  const xlsx = zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`),
    "xl/sharedStrings.xml": strToU8(sst),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
  return readWorkbook(xlsx);
}

describe("rich text reader", () => {
  it("captures per-run bold/italic/colour into richRuns", () => {
    const wb = book(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t xml:space="preserve">Hi </t></r><r><rPr><b/><color rgb="FFCC0000"/></rPr><t>red</t></r><r><rPr><i/></rPr><t> it</t></r></si></sst>`);
    const cell = wb.sheets[0].cells.get("1:1");
    expect(cell?.value).toBe("Hi red it");
    const runs = cell?.richRuns;
    expect(runs?.length).toBe(3);
    expect(runs?.[0]).toMatchObject({ text: "Hi " });
    expect(runs?.[1]).toMatchObject({ text: "red", bold: true });
    expect(runs?.[1].color?.toLowerCase()).toContain("cc0000");
    expect(runs?.[2]).toMatchObject({ text: " it", italic: true });
  });

  it("leaves richRuns undefined for a single-run (unstyled) string", () => {
    const wb = book(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>plain</t></si></sst>`);
    const cell = wb.sheets[0].cells.get("1:1");
    expect(cell?.value).toBe("plain");
    expect(cell?.richRuns).toBeUndefined();
  });
});
