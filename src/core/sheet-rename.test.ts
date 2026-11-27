import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "./workbook";
import { renameSheet } from "./sheet-ops";
import { renameSheetInOdf } from "./sheet-rename";
import { setXlsxCondFormat, setXlsxDataValidation } from "../adapters/xlsx/write";

// Renaming a sheet must reach every formula that names it, not only the cell formulas:
// conditional formats, validations, charts, sparklines and pivot sources too.

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = (t: string) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${t}`;

function xlsx(): Uint8Array {
  const chart =
    `<?xml version="1.0"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>` +
    `</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
  const drawing =
    `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}" xmlns:c="${C}"><xdr:twoCellAnchor>` +
    `<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame><a:graphic><a:graphicData uri="${C}"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
  const rels = (items: [string, string][]) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([t, target], i) => `<Relationship Id="rId${i + 1}" Type="${REL(t)}" Target="${target}"/>`).join("")}</Relationships>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(rels([["officeDocument", "xl/workbook.xml"]])),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Report" sheetId="2" r:id="rId2"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(rels([["worksheet", "worksheets/sheet1.xml"], ["worksheet", "worksheets/sheet2.xml"]])),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c><c r="B1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row></sheetData></worksheet>`),
    "xl/worksheets/sheet2.xml": strToU8(
      `<worksheet xmlns="${S}" xmlns:r="${R}"><sheetData><row r="1"><c r="A1"><f>Data!B1*2</f><v>2</v></c></row></sheetData>` +
        `<conditionalFormatting sqref="A1:A5"><cfRule type="expression" priority="1"><formula>Data!$B$1&gt;0</formula></cfRule></conditionalFormatting>` +
        `<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B1:B5"><formula1>Data!$A$1:$A$2</formula1></dataValidation></dataValidations>` +
        `<hyperlinks><hyperlink ref="C1" location="Data!A1"/><hyperlink ref="C2" location="'Data'!B2"/><hyperlink ref="C3" location="Report!A1"/></hyperlinks>` +
        `<drawing r:id="rId1"/></worksheet>`,
    ),
    "xl/worksheets/_rels/sheet2.xml.rels": strToU8(rels([["drawing", "../drawings/drawing1.xml"]])),
    "xl/drawings/drawing1.xml": strToU8(drawing),
    "xl/drawings/_rels/drawing1.xml.rels": strToU8(rels([["chart", "../charts/chart1.xml"]])),
    "xl/charts/chart1.xml": strToU8(chart),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${S}"/>`),
  });
}

describe("renaming a sheet in an xlsx", () => {
  it("rewrites cell, conditional-format, validation and chart references", () => {
    const wb = readWorkbook(xlsx());
    const report = wb.sheets[1]!;
    expect(report.condFormats?.[0]?.rules[0]?.formulas).toEqual(["Data!$B$1>0"]);
    renameSheet(wb, 0, "Raw data");
    expect(report.condFormats?.[0]?.rules[0]?.formulas).toEqual(["'Raw data'!$B$1>0"]);
    expect(report.validations?.[0]?.rangeRef).toBe("'Raw data'!$A$1:$A$2");
    expect(report.charts?.[0]?.series[0]?.values.ref).toBe("'Raw data'!$B$1:$B$2");
    expect(report.charts?.[0]?.categories?.ref).toBe("'Raw data'!$A$1:$A$2");

    const files = unzipSync(writeWorkbook(wb));
    const sheet2 = strFromU8(files["xl/worksheets/sheet2.xml"]!);
    expect(sheet2).toContain("<f>'Raw data'!B1*2</f>");
    expect(sheet2).toContain("<formula>'Raw data'!$B$1&gt;0</formula>");
    expect(sheet2).toContain("<formula1>'Raw data'!$A$1:$A$2</formula1>");
    const chart = strFromU8(files["xl/charts/chart1.xml"]!);
    expect(chart).toContain("<c:f>'Raw data'!$B$1:$B$2</c:f>");
    expect(chart).not.toContain("Data!");
    expect(sheet2).toContain(`<hyperlink ref="C1" location="'Raw data'!A1"/>`);
    expect(sheet2).toContain(`<hyperlink ref="C2" location="'Raw data'!B2"/>`);
    expect(sheet2).toContain(`<hyperlink ref="C3" location="Report!A1"/>`);
    expect(report.cells.get("2:3")?.link).toEqual({ href: "'Raw data'!B2", internal: true, tip: undefined });

    const again = readWorkbook(writeWorkbook(wb));
    expect(again.sheets[1]!.validations?.[0]?.rangeRef).toBe("'Raw data'!$A$1:$A$2");
  });

  it("also reaches rules authored in this session", () => {
    const wb = readWorkbook(xlsx());
    const report = wb.sheets[1]!;
    setXlsxCondFormat(wb, report, [{ r1: 7, c1: 1, r2: 9, c2: 1 }], { kind: "expression", formula: "Data!$B$2=2", fill: "FFFF00" });
    setXlsxDataValidation(report, [{ r1: 7, c1: 2, r2: 9, c2: 2 }], { type: "custom", formula1: "Data!$B$2>1" });
    renameSheet(wb, 0, "Q1");
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet2.xml"]!);
    expect(xml).toContain("'Q1'!$B$2=2");
    expect(xml).toContain("'Q1'!$B$2&gt;1");
  });
});

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:xlink="http://www.w3.org/1999/xlink"`;

function ods(): Uint8Array {
  const content =
    `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
    `<table:content-validations><table:content-validation table:name="v1" table:condition="of:cell-content-is-in-list([$Data.$A$1:.$A$2])" table:allow-empty-cell="true" table:base-cell-address="Report.A1"/></table:content-validations>` +
    `<table:table table:name="Data"><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table>` +
    `<table:table table:name="Report"><table:table-row><table:table-cell table:content-validation-name="v1" table:formula="of:=[$Data.A1]*2" office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>` +
    `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="#Data.B1">go</text:a></text:p></table:table-cell>` +
    `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://data.test/Data.A1x">web</text:a></text:p></table:table-cell>` +
    `</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("renaming a sheet in an ods", () => {
  it("rewrites ODF references in place", () => {
    expect(renameSheetInOdf("of:cell-content-is-in-list([$Data.$A$1:.$A$2])", "data", "Raw data")).toBe("of:cell-content-is-in-list([$'Raw data'.$A$1:.$A$2])");
    expect(renameSheetInOdf("Data.A1:Data.B5", "Data", "New")).toBe("New.A1:New.B5");
    expect(renameSheetInOdf("$'My data'.A1 Database.A1 [.A1]", "My data", "Data2")).toBe("$Data2.A1 Database.A1 [.A1]");
  });

  it("rewrites the validation and the formula on save", () => {
    const wb = readWorkbook(ods());
    renameSheet(wb, 0, "Raw data");
    expect(wb.sheets[1]!.validations?.[0]?.rangeRef).toBe("'Raw data'!A1:A2");
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(xml).toContain("[$'Raw data'.$A$1:.$A$2]");
    expect(xml).toContain("of:=[$'Raw data'.A1]*2");
    expect(xml).not.toContain("$Data.");
    expect(wb.sheets[1]!.cells.get("1:2")?.link).toEqual({ href: "'Raw data'!B1", internal: true });
    expect(xml).toContain(`xlink:href="#'Raw data'.B1"`);
    expect(xml).toContain(`xlink:href="https://data.test/Data.A1x"`);
    expect(readWorkbook(writeWorkbook(wb)).sheets[1]!.cells.get("1:2")?.link?.href).toBe("'Raw data'!B1");
  });
});
