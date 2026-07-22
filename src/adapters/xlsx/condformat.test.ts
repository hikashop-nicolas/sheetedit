import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { computeCondVisuals } from "./condformat";
import { key } from "../../core/model";

function xlsx(): Uint8Array {
  const cell = (ref: string, v: number): string => `<c r="${ref}"><v>${v}</v></c>`;
  const sheet = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<sheetData>` +
      `<row r="1">${cell("A1", 5)}${cell("B1", 0)}${cell("C1", 10)}</row>` +
      `<row r="2">${cell("A2", 15)}${cell("B2", 50)}${cell("C2", 20)}</row>` +
      `<row r="3">${cell("A3", 15)}${cell("B3", 100)}${cell("C3", 40)}</row>` +
      `</sheetData>` +
      `<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule></conditionalFormatting>` +
      `<conditionalFormatting sqref="B1:B3"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>` +
      `<conditionalFormatting sqref="C1:C3"><cfRule type="dataBar" priority="3"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>` +
      `</worksheet>`,
  );
  const styles = strToU8(
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
      `</styleSheet>`,
  );
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": sheet,
    "xl/styles.xml": styles,
  });
}

describe("conditional formatting", () => {
  it("applies cellIs dxf, a colour scale and a data bar", () => {
    const wb = readWorkbook(xlsx());
    const vis = computeCondVisuals(wb.sheets[0]);
    // cellIs > 10: A2 and A3 match, A1 does not.
    expect(vis.get(key(2, 1))).toEqual({ bg: "#ffc7ce", color: "#9c0006" });
    expect(vis.has(key(1, 1))).toBe(false);
    // colour scale white->blue: B1 (min) is white, B3 (max) is blue.
    expect(vis.get(key(1, 2))?.bg).toBe("#ffffff");
    expect(vis.get(key(3, 2))?.bg).toBe("#0000ff");
    // data bar: C1 shortest, C3 full.
    expect(vis.get(key(1, 3))?.bar?.pct).toBeCloseTo(0, 5);
    expect(vis.get(key(3, 3))?.bar?.pct).toBeCloseTo(1, 5);
    expect(vis.get(key(2, 3))?.bar?.pct).toBeCloseTo((20 - 10) / (40 - 10), 5);
  });
});
