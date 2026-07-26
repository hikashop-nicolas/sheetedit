import { describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../index";
import type { PrintSetup } from "../core/print";

// End-to-end shape of an authored page setup in both formats. Set SHEETEDIT_WRITE_FIXTURES=1 to
// also drop the demo/c-print.* files, which are what gets handed to LibreOffice for the external
// check.

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";
const ROWS = [
  ["Region", "Q1", "Q2", "Q3"],
  ["North", "10", "12", "14"],
  ["South", "20", "22", "24"],
  ["East", "30", "32", "34"],
  ["West", "40", "42", "44"],
];

/** The setup both fixtures author, chosen so every modelled field is non-default. */
export const SETUP: PrintSetup = {
  orientation: "landscape",
  paperSize: 9,
  scale: 100,
  fitToWidth: 1,
  fitToHeight: 0,
  fitToPage: true,
  pageOrder: "overThenDown",
  firstPageNumber: 3,
  margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  gridLines: true,
  headings: true,
  horizontalCentered: true,
  verticalCentered: false,
  header: { left: "Quarterly", center: "Sales", right: "&D" },
  footer: { center: "Page &P of &N" },
  printArea: [{ r1: 1, c1: 1, r2: 5, c2: 4 }],
  titleRows: { from: 1, to: 1 },
  rowBreaks: [3],
  colBreaks: [3],
};

const emit = (name: string, bytes: Uint8Array): void => {
  if (process.env.SHEETEDIT_WRITE_FIXTURES) writeFileSync(`demo/${name}`, bytes);
};

function buildXlsx(): Uint8Array {
  const rows = ROWS.map((vals, i) => `<row r="${i + 1}">${vals.map((v, j) => `<c r="${String.fromCharCode(65 + j)}${i + 1}" t="inlineStr"><is><t>${v}</t></is></c>`).join("")}</row>`).join("");
  const src = zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="Sales" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}"><sheetData>${rows}</sheetData></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`),
  });
  const wb = readWorkbook(src);
  wb.sheets[0]!.printSetup = { ...SETUP };
  wb.sheets[0]!.printDirty = true;
  return writeWorkbook(wb);
}

const ODS_NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"` +
  ` xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
  ` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"`;

function buildOds(): Uint8Array {
  const tr = (vals: string[]) => `<table:table-row>${vals.map((v) => `<table:table-cell office:value-type="string"><text:p>${v}</text:p></table:table-cell>`).join("")}</table:table-row>`;
  const content =
    `<?xml version="1.0"?><office:document-content ${ODS_NS}><office:automatic-styles/><office:body><office:spreadsheet>` +
    `<table:table table:name="Sales"><table:table-column table:number-columns-repeated="4"/>${ROWS.map(tr).join("")}</table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const styles = `<?xml version="1.0"?><office:document-styles ${ODS_NS}><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>`;
  const src = zipSync({
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(content),
    "styles.xml": strToU8(styles),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>`),
  });
  const wb = readWorkbook(src);
  // ODF has no page-order-down equivalent worth asserting on, and titleRows is the header-rows
  // group, which this builder does not create; the rest is the same setup.
  wb.sheets[0]!.printSetup = { ...SETUP, titleRows: undefined };
  wb.sheets[0]!.printDirty = true;
  return writeWorkbook(wb);
}

describe("authored print setup, end to end", () => {
  it("an authored xlsx reads back with the same page setup", () => {
    const bytes = buildXlsx();
    emit("c-print.xlsx", bytes);
    const p = readWorkbook(bytes).sheets[0]!.printSetup!;
    expect(p.orientation).toBe("landscape");
    expect(p.paperSize).toBe(9);
    expect(p.fitToPage).toBe(true);
    expect(p.fitToWidth).toBe(1);
    expect(p.fitToHeight).toBe(0);
    expect(p.pageOrder).toBe("overThenDown");
    expect(p.firstPageNumber).toBe(3);
    expect(p.margins).toEqual(SETUP.margins);
    expect(p.gridLines).toBe(true);
    expect(p.headings).toBe(true);
    expect(p.horizontalCentered).toBe(true);
    expect(p.header).toEqual(SETUP.header);
    expect(p.footer).toEqual(SETUP.footer);
    expect(p.printArea).toEqual(SETUP.printArea);
    expect(p.titleRows).toEqual({ from: 1, to: 1 });
    expect(p.rowBreaks).toEqual([3]);
    expect(p.colBreaks).toEqual([3]);
  });

  it("an authored ods reads back with the same page setup", () => {
    const bytes = buildOds();
    emit("c-print.ods", bytes);
    const p = readWorkbook(bytes).sheets[0]!.printSetup!;
    expect(p.orientation).toBe("landscape");
    expect(p.paperSize).toBe(9);
    expect(p.fitToPage).toBe(true);
    expect(p.fitToWidth).toBe(1);
    expect(p.pageOrder).toBe("overThenDown");
    expect(p.firstPageNumber).toBe(3);
    // The margin split (page margin vs header block) has to survive the conversion both ways.
    expect(p.margins).toEqual(SETUP.margins);
    expect(p.gridLines).toBe(true);
    expect(p.headings).toBe(true);
    expect(p.horizontalCentered).toBe(true);
    expect(p.header).toEqual(SETUP.header);
    expect(p.footer).toEqual(SETUP.footer);
    expect(p.printArea).toEqual(SETUP.printArea);
    expect(p.rowBreaks).toEqual([3]);
    expect(p.colBreaks).toEqual([3]);
  });
});
