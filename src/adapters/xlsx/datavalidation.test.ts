import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";

function xlsxWithDv(): Uint8Array {
  const sheet = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<sheetData><row r="1"><c r="D1" t="inlineStr"><is><t>Red</t></is></c></row><row r="2"><c r="D2" t="inlineStr"><is><t>Green</t></is></c></row></sheetData>` +
      `<dataValidations count="2">` +
      `<dataValidation type="list" allowBlank="1" sqref="A1:A100"><formula1>"Yes,No,Maybe"</formula1></dataValidation>` +
      `<dataValidation type="list" sqref="B1"><formula1>$D$1:$D$2</formula1></dataValidation>` +
      `<dataValidation type="whole" sqref="C1"><formula1>1</formula1><formula2>9</formula2></dataValidation>` +
      `</dataValidations></worksheet>`,
  );
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": sheet,
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

describe("xlsx data validation (dropdowns)", () => {
  it("reads inline / range list validations and typed (non-list) rules", () => {
    const wb = readWorkbook(xlsxWithDv());
    const dvs = wb.sheets[0].validations!;
    expect(dvs).toHaveLength(3); // two list rules + the "whole" constraint
    const inline = dvs.find((d) => d.values);
    expect(inline?.values).toEqual(["Yes", "No", "Maybe"]);
    expect(inline?.allowBlank).toBe(true);
    expect(inline?.ranges[0]).toEqual({ r1: 1, c1: 1, r2: 100, c2: 1 });
    const range = dvs.find((d) => d.rangeRef);
    expect(range?.rangeRef).toBe("$D$1:$D$2");
    expect(range?.ranges[0]).toEqual({ r1: 1, c1: 2, r2: 1, c2: 2 });
    expect(dvs.find((d) => d.type === "whole")).toBeTruthy();
  });
});
