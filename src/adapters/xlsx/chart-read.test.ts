import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook } from "../../index";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function chartXml(): string {
  const ser = (i: number, name: string, vals: number[]) =>
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
    `<c:tx><c:strRef><c:f>Sheet1!$${String.fromCharCode(66 + i)}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
    `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>Sheet1!$${String.fromCharCode(66 + i)}$2:$${String.fromCharCode(66 + i)}$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>${vals[0]}</c:v></c:pt><c:pt idx="1"><c:v>${vals[1]}</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>`;
  return `<?xml version="1.0"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart>` +
    `<c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
    `<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, "North", [10, 20])}${ser(1, "South", [15, 25])}</c:barChart></c:plotArea>` +
    `<c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>`;
}

function drawingXml(): string {
  return `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}" xmlns:c="${C}">` +
    `<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame><a:graphic><a:graphicData uri="${C}"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
}

function xlsx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>`),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`),
    "xl/drawings/drawing1.xml": strToU8(drawingXml()),
    "xl/drawings/_rels/drawing1.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`),
    "xl/charts/chart1.xml": strToU8(chartXml()),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

function xlsxSchemeColor(): Uint8Array {
  const chart = `<?xml version="1.0"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>` +
    `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>` +
    `</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
  const theme = `<a:theme xmlns:a="${A}"><a:themeElements><a:clrScheme name="x"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="C0504D"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`;
  const base = new Map(Object.entries(unzipSync(xlsx())));
  base.set("xl/theme/theme1.xml", strToU8(theme));
  base.set("xl/drawings/drawing1.xml", strToU8(drawingXml()));
  base.set("xl/charts/chart1.xml", strToU8(chart));
  return zipSync(Object.fromEntries(base));
}

function xlsxMultiLevelCats(): Uint8Array {
  const chart = `<?xml version="1.0"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/>` +
    `<c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$3</c:f><c:multiLvlStrCache><c:ptCount val="2"/>` +
    `<c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:lvl>` +
    `<c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:lvl>` +
    `</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>` +
    `<c:val><c:numRef><c:f>Sheet1!$C$2:$C$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>` +
    `</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
  const base = new Map(Object.entries(unzipSync(xlsx())));
  base.set("xl/drawings/drawing1.xml", strToU8(drawingXml()));
  base.set("xl/charts/chart1.xml", strToU8(chart));
  return zipSync(Object.fromEntries(base));
}

describe("xlsx chart reader", () => {
  it("resolves a schemeClr series colour via the theme", () => {
    const wb = readWorkbook(xlsxSchemeColor());
    expect(wb.sheets[0].charts![0].series[0].color).toBe("#c0504d");
  });

  it("reads multi-level categories down to the innermost level", () => {
    const wb = readWorkbook(xlsxMultiLevelCats());
    expect(wb.sheets[0].charts![0].categories).toEqual({ ref: "Sheet1!$A$2:$B$3", cache: ["Jan", "Feb"] });
  });

  it("reads a base palette from the chart's colours part", () => {
    const cs = `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="${A}" meth="cycle" id="10"><a:srgbClr val="112233"/><a:schemeClr val="accent1"/><cs:variation/></cs:colorStyle>`;
    const theme = `<a:theme xmlns:a="${A}"><a:themeElements><a:clrScheme name="x"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="AABBCC"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`;
    const base = new Map(Object.entries(unzipSync(xlsx())));
    base.set("xl/theme/theme1.xml", strToU8(theme));
    base.set("xl/charts/colors1.xml", strToU8(cs));
    base.set("xl/charts/_rels/chart1.xml.rels", strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>`));
    const wb = readWorkbook(zipSync(Object.fromEntries(base)));
    expect(wb.sheets[0].charts![0].palette).toEqual(["#112233", "#aabbcc"]);
  });


  it("parses a bar chart with two series, categories, title and legend", () => {
    const wb = readWorkbook(xlsx());
    const charts = wb.sheets[0].charts!;
    expect(charts).toHaveLength(1);
    const ch = charts[0];
    expect(ch.kind).toBe("column");
    expect(ch.title).toBe("Sales");
    expect(ch.legend).toEqual({ show: true, pos: "bottom" });
    expect(ch.series).toHaveLength(2);
    expect(ch.series[0].name).toEqual({ ref: "Sheet1!$B$1", cache: ["North"] });
    expect(ch.series[0].values).toEqual({ ref: "Sheet1!$B$2:$B$3", cache: [10, 20] });
    expect(ch.categories).toEqual({ ref: "Sheet1!$A$2:$A$3", cache: ["Q1", "Q2"] });
    // twoCellAnchor from D2 (col 3,row 1 -> 1-based 4,2) to J17 (col 9,row 16 -> 10,17).
    expect(ch.anchor.fromCol).toBe(4);
    expect(ch.anchor.fromRow).toBe(2);
    expect(ch.anchor.toCol).toBe(10);
    expect(ch.anchor.toRow).toBe(17);
    expect(ch.original?.partPath).toBe("xl/charts/chart1.xml");
  });
});
