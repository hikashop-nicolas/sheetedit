import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { setXlsxDataValidation, setXlsxHyperlink } from "./write";
import { serializeXml } from "../../core/model";

function base(): ReturnType<typeof readWorkbook> {
  const xlsx = zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
  return readWorkbook(xlsx);
}

describe("authoring writers", () => {
  it("writes a list data validation and round-trips it", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxDataValidation(sheet, [{ r1: 1, c1: 1, r2: 9, c2: 1 }], { values: ["Yes", "No", "Maybe"], allowBlank: true });
    const xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).toContain('<dataValidation type="list"');
    expect(xml).toContain('sqref="A1:A9"');
    expect(xml).toContain('"Yes,No,Maybe"');
    expect(sheet.validations?.[0].values).toEqual(["Yes", "No", "Maybe"]);
  });

  it("writes an internal and an external hyperlink", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxHyperlink(wb, sheet, 1, 1, { href: "Sheet1!B2", internal: true, tip: "go" });
    let xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).toContain('<hyperlink ref="A1" location="Sheet1!B2" tooltip="go"');
    setXlsxHyperlink(wb, sheet, 2, 2, { href: "https://example.com" });
    xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).toMatch(/<hyperlink ref="B2" r:id="rId\d+"/);
    const rels = new TextDecoder().decode(wb.files["xl/worksheets/_rels/sheet1.xml.rels"]);
    expect(rels).toContain('TargetMode="External"');
    expect(rels).toContain('Target="https://example.com"');
    // remove
    setXlsxHyperlink(wb, sheet, 1, 1, null);
    xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).not.toContain('ref="A1"');
  });
});
