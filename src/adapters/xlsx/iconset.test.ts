import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../core/workbook";
import { computeCondVisuals } from "./condformat";
import { key } from "../../core/model";

// A column A1:A5 = 1..5 with a 3-icon set over 0/33/67 percent thresholds.
function book(): ReturnType<typeof readWorkbook> {
  const rows = [1, 2, 3, 4, 5].map((v, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${v}</v></c></row>`).join("");
  const cf = `<conditionalFormatting sqref="A1:A5"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>`;
  const sheet = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData>${cf}</worksheet>`;
  const xlsx = zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(sheet),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
  return readWorkbook(xlsx);
}

describe("icon set conditional formatting", () => {
  it("reads the iconSet rule", () => {
    const wb = book();
    const rule = wb.sheets[0].condFormats?.[0].rules[0];
    expect(rule?.type).toBe("iconSet");
    expect(rule?.iconSet?.set).toBe("3TrafficLights1");
    expect(rule?.iconSet?.cfvo.length).toBe(3);
  });

  it("renders an icon in each cell, low->high buckets", () => {
    const wb = book();
    const vis = computeCondVisuals(wb.sheets[0]);
    // Every cell gets an icon SVG.
    for (let r = 1; r <= 5; r++) expect(vis.get(key(r, 1))?.icon, `row ${r}`).to.contain("<svg");
    // Value 1 (min, red = hue 0) vs value 5 (max, green = hue 120) differ.
    expect(vis.get(key(1, 1))?.icon).to.contain("hsl(0,");
    expect(vis.get(key(5, 1))?.icon).to.contain("hsl(120,");
  });
});
