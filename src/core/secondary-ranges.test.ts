import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../index";
import { applyLineOp } from "./structure";

// A workbook whose sheet carries conditional formatting, a data validation and a hyperlink, so
// inserting/deleting rows must shift their sqref/ref (which sheetedit otherwise leaves untouched).
function xlsx(): Uint8Array {
  const sheet = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheetData><row r="5"><c r="A5" t="inlineStr"><is><t>x</t></is></c></row></sheetData>` +
      `<conditionalFormatting sqref="A5:A10"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>` +
      `<dataValidations count="1"><dataValidation type="list" sqref="A5:A10"><formula1>"a,b"</formula1></dataValidation></dataValidations>` +
      `<hyperlinks><hyperlink ref="A5" r:id="rId1"/></hyperlinks>` +
      `</worksheet>`,
  );
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": sheet,
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://x/" TargetMode="External"/></Relationships>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}
const wsText = (bytes: Uint8Array): string => new TextDecoder().decode(unzipSync(bytes)["xl/worksheets/sheet1.xml"]);

describe("structural edits shift secondary ranges", () => {
  it("inserting rows above shifts CF, DV and hyperlink refs", () => {
    const wb = readWorkbook(xlsx());
    applyLineOp(wb, 0, { axis: "row", kind: "insert", at: 1, count: 2 }); // 2 rows at the top
    const xml = wsText(writeWorkbook(wb));
    expect(xml).toContain(`<conditionalFormatting sqref="A7:A12"`);
    expect(xml).toContain(`sqref="A7:A12"`); // dataValidation
    expect(xml).toContain(`<hyperlink ref="A7"`);
    // The model validation used for rendering shifted too.
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].validations![0].ranges[0]).toEqual({ r1: 7, c1: 1, r2: 12, c2: 1 });
  });
  it("deleting rows inside a range shrinks it", () => {
    const wb = readWorkbook(xlsx());
    applyLineOp(wb, 0, { axis: "row", kind: "delete", at: 6, count: 2 }); // remove rows 6-7 (inside A5:A10)
    const xml = wsText(writeWorkbook(wb));
    expect(xml).toContain(`sqref="A5:A8"`); // A5:A10 minus 2 rows
  });
});
