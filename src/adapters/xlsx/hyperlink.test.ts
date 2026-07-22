import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, getCell } from "../../index";

function xlsxWithLinks(): Uint8Array {
  const sheet = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheetData>` +
      `<row r="1"><c r="A1" t="inlineStr"><is><t>site</t></is></c></row>` +
      `<row r="2"><c r="A2" t="inlineStr"><is><t>go</t></is></c></row>` +
      `</sheetData>` +
      `<hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="A2" location="Sheet1!B5" tooltip="jump"/></hyperlinks>` +
      `</worksheet>`,
  );
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": sheet,
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

describe("xlsx hyperlinks", () => {
  it("reads external and internal links onto their cells", () => {
    const wb = readWorkbook(xlsxWithLinks());
    const a1 = getCell(wb.sheets[0], 1, 1);
    expect(a1?.link).toEqual({ href: "https://example.com/", internal: false, tip: undefined });
    const a2 = getCell(wb.sheets[0], 2, 1);
    expect(a2?.link).toEqual({ href: "Sheet1!B5", internal: true, tip: "jump" });
  });
});
