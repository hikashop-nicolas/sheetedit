import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { setXlsxComment, setXlsxCondFormat, setXlsxDataValidation, setXlsxHyperlink, writeXlsx } from "./write";
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

  it("authors a comment that round-trips (parts + rel + content type)", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxComment(wb, sheet, 2, 3, "Please review", "Ada");
    writeXlsx(wb); // serialize sheet.doc (legacyDrawing) into wb.files
    const re = readWorkbook(zipSync(wb.files));
    const cell = re.sheets[0].cells.get("2:3");
    expect(cell?.comments?.[0]).toEqual({ author: "Ada", text: "Please review" });
    // parts + registration exist
    expect(Object.keys(wb.files).some((f) => /comments\d+\.xml$/.test(f))).toBe(true);
    expect(Object.keys(wb.files).some((f) => /vmlDrawing\d+\.vml$/.test(f))).toBe(true);
    expect(new TextDecoder().decode(wb.files["[Content_Types].xml"])).toContain("comments+xml");
    const srels = new TextDecoder().decode(wb.files["xl/worksheets/_rels/sheet1.xml.rels"]);
    expect(srels).toMatch(/relationships\/comments/);
    expect(srels).toMatch(/vmlDrawing/);
  });

  it("authors a cellIs conditional format (dxf + cfRule) that round-trips", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxCondFormat(wb, sheet, [{ r1: 2, c1: 1, r2: 9, c2: 1 }], { kind: "cellIs", operator: "greaterThan", value: "5", fill: "#ffc7ce" });
    const xml = new TextDecoder().decode(serializeXml(sheet.doc!));
    expect(xml).toContain('<conditionalFormatting sqref="A2:A9">');
    expect(xml).toMatch(/<cfRule type="cellIs"[^>]*operator="greaterThan"/);
    expect(new TextDecoder().decode(serializeXml(wb.stylesDoc!))).toContain("FFffc7ce");
    // model updated so it renders
    expect(sheet.condFormats?.[0].rules[0].dxf?.bg).toBe("#ffc7ce");
  });

  it("authors the extended conditional-format rule types", () => {
    const wb = base(); const sheet = wb.sheets[0];
    const R = [{ r1: 1, c1: 1, r2: 9, c2: 1 }];
    const xml = () => new TextDecoder().decode(serializeXml(sheet.doc!));

    setXlsxCondFormat(wb, sheet, R, { kind: "cellIs", operator: "between", value: "1", value2: "5", fill: "#ffc7ce" });
    expect(xml()).toMatch(/<cfRule type="cellIs"[^>]*operator="between"[\s\S]*?<formula>1<\/formula><formula>5<\/formula>/);

    setXlsxCondFormat(wb, sheet, R, { kind: "text", operator: "containsText", text: "err", fill: "#ffc7ce" });
    expect(xml()).toMatch(/<cfRule type="containsText"[^>]*text="err"/);
    expect(xml()).toContain('SEARCH("err"');

    setXlsxCondFormat(wb, sheet, R, { kind: "top", rank: 3, bottom: true, fill: "#ffc7ce" });
    expect(xml()).toMatch(/<cfRule type="top10"[^>]*rank="3"[^>]*bottom="1"/);

    setXlsxCondFormat(wb, sheet, R, { kind: "average", below: true, fill: "#ffc7ce" });
    expect(xml()).toMatch(/<cfRule type="aboveAverage"[^>]*aboveAverage="0"/);

    setXlsxCondFormat(wb, sheet, R, { kind: "dupUnique", unique: true, fill: "#ffc7ce" });
    expect(xml()).toMatch(/<cfRule type="uniqueValues"/);

    setXlsxCondFormat(wb, sheet, R, { kind: "expression", formula: "A1>AVERAGE($A$1:$A$9)", fill: "#ffc7ce" });
    expect(xml()).toMatch(/<cfRule type="expression"[\s\S]*?<formula>A1&gt;AVERAGE/);

    setXlsxCondFormat(wb, sheet, R, { kind: "iconSet", set: "3Arrows", count: 3 });
    expect(xml()).toMatch(/<cfRule type="iconSet"[\s\S]*?<iconSet iconSet="3Arrows"><cfvo type="percent" val="0"\/><cfvo type="percent" val="33"\/><cfvo type="percent" val="67"\/>/);
    expect(sheet.condFormats?.[0].rules[0].iconSet?.set).toBe("3Arrows");

    setXlsxCondFormat(wb, sheet, R, { kind: "timePeriod", period: "last7Days", fill: "#ffc7ce" });
    expect(xml()).toMatch(/<cfRule type="timePeriod"[^>]*timePeriod="last7Days"/);
    expect(xml()).toContain("AND(TODAY()-FLOOR(A1,1)&lt;=6");
    expect(sheet.condFormats?.[0].rules[0].timePeriod).toBe("last7Days");
  });
});
