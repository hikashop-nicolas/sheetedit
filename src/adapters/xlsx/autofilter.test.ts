import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { setXlsxAutoFilter, setXlsxRowHidden } from "./write";
import { serializeXml } from "../../core/model";

function xlsx(withAf: boolean): Uint8Array {
  const c = (ref: string, v: string | number, s = false) => (s ? `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>` : `<c r="${ref}"><v>${v}</v></c>`);
  const af = withAf ? `<autoFilter ref="A1:B3"/>` : "";
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${c("A1","Name",true)}${c("B1","Qty",true)}</row><row r="2">${c("A2","Pen",true)}${c("B2",3)}</row><row r="3">${c("A3","Book",true)}${c("B3",8)}</row></sheetData>${af}</worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

describe("autofilter", () => {
  it("reads the autoFilter range into the sheet", () => {
    const wb = readWorkbook(xlsx(true));
    expect(wb.sheets[0].autoFilter).toEqual({ r1: 1, c1: 1, r2: 3, c2: 2 });
  });

  it("creates, then removes, the autoFilter element", () => {
    const wb = readWorkbook(xlsx(false));
    const sheet = wb.sheets[0];
    setXlsxAutoFilter(sheet, "A1:B3");
    let xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).toContain('<autoFilter ref="A1:B3"');
    setXlsxAutoFilter(sheet, null);
    xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).not.toContain("autoFilter");
  });

  it("sets and clears a row's hidden attribute", () => {
    const wb = readWorkbook(xlsx(true));
    const sheet = wb.sheets[0];
    setXlsxRowHidden(sheet, 2, true);
    let xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).toMatch(/<row[^>]*r="2"[^>]*hidden="1"/);
    setXlsxRowHidden(sheet, 2, false);
    xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).not.toMatch(/<row[^>]*r="2"[^>]*hidden/);
  });
});
