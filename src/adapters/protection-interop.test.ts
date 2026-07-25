import { describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../index";
import { canEditCell, isBlocked, isProtected } from "../core/protection";
import { setXlsxCellStyle } from "./xlsx/styles";
import { setOdsCellStyle } from "./ods/styles";

// End-to-end shape of an authored protected workbook, in both formats: a Stock sheet whose Qty
// column is unlocked, sheet protection granting one non-default allowance, and a locked workbook
// structure. These are the demo/c-protect.* fixtures; set SHEETEDIT_WRITE_FIXTURES=1 to rewrite
// them from this builder.
//
// Both outputs were checked against LibreOffice (26.2) by hand:
//   xlsx -> ods : table:protected="true", the select permissions, and style:cell-protect="none"
//                 on the unlocked Qty cells.
//   ods  -> xlsx: <sheetProtection sheet="true" objects="true" scenarios="true" insertRows="false"/>
//                 plus a cellXf carrying locked="false" that the Qty cells point at.
// LibreOffice drops workbook structure protection on its OOXML export (its own ods -> ods keeps
// it), so that flag is verified within ODF only.

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";
const ROWS = [["Item", "Qty", "Note"], ["Bolt", "12", "ok"], ["Nut", "40", "ok"]];

const emit = (name: string, bytes: Uint8Array): void => {
  if (process.env.SHEETEDIT_WRITE_FIXTURES) writeFileSync(`demo/${name}`, bytes);
};

function buildXlsx(): Uint8Array {
  const rows = ROWS.map((vals, i) => `<row r="${i + 1}">${vals.map((v, j) => `<c r="${String.fromCharCode(65 + j)}${i + 1}" t="inlineStr"><is><t>${v}</t></is></c>`).join("")}</row>`).join("");
  const src = zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="Stock" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    // The styles relationship is what makes a reader load styles.xml at all. Without it the
    // per-cell lock state is invisible to Excel and LibreOffice, which then treat every cell as
    // locked; verified against LibreOffice, which collapsed the whole style pool when it was absent.
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}"><sheetData>${rows}</sheetData></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`),
  });
  const wb = readWorkbook(src);
  const sheet = wb.sheets[0]!;
  for (let r = 2; r <= 3; r++) setXlsxCellStyle(wb, sheet, sheet.cells.get(`${r}:2`)!, { locked: false });
  sheet.protection = { sheet: true, locks: { objects: true, scenarios: true, sort: false } };
  sheet.protectionDirty = true;
  wb.protection = { structure: true };
  wb.protectionDirty = true;
  return writeWorkbook(wb);
}

function buildOds(): Uint8Array {
  const NS =
    `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"` +
    ` xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
    ` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"` +
    ` xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"`;
  const tr = (vals: string[]) => `<table:table-row>${vals.map((v) => `<table:table-cell office:value-type="string"><text:p>${v}</text:p></table:table-cell>`).join("")}</table:table-row>`;
  const content =
    `<?xml version="1.0"?><office:document-content ${NS}><office:automatic-styles/><office:body><office:spreadsheet>` +
    `<table:table table:name="Stock">${ROWS.map(tr).join("")}</table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const src = zipSync({
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`),
  });
  const wb = readWorkbook(src);
  const sheet = wb.sheets[0]!;
  for (let r = 2; r <= 3; r++) setOdsCellStyle(wb, sheet, sheet.cells.get(`${r}:2`)!, { locked: false });
  sheet.protection = { sheet: true, locks: { insertRows: false } };
  sheet.protectionDirty = true;
  wb.protection = { structure: true };
  wb.protectionDirty = true;
  return writeWorkbook(wb);
}

describe("authored protection, end to end", () => {
  it("an authored xlsx reads back with the same protection", () => {
    const bytes = buildXlsx();
    emit("c-protect.xlsx", bytes);
    const wb = readWorkbook(bytes);
    const s = wb.sheets[0]!;
    expect(isProtected(s)).toBe(true);
    expect(isBlocked(s, "sort")).toBe(false); // the one allowance granted
    expect(isBlocked(s, "insertRows")).toBe(true);
    expect(wb.protection?.structure).toBe(true);
    // Only the Qty column is editable while protected.
    expect(canEditCell(s, 2, 2)).toBe(true);
    expect(canEditCell(s, 3, 2)).toBe(true);
    expect(canEditCell(s, 2, 1)).toBe(false);
    expect(canEditCell(s, 1, 2)).toBe(false); // the header stays locked
  });

  it("the authored xlsx declares the styles relationship other readers need", () => {
    const rels = new TextDecoder().decode(readWorkbook(buildXlsx()).files["xl/_rels/workbook.xml.rels"]!);
    expect(rels).toContain("styles.xml");
  });

  it("an authored ods reads back with the same protection", () => {
    const bytes = buildOds();
    emit("c-protect.ods", bytes);
    const wb = readWorkbook(bytes);
    const s = wb.sheets[0]!;
    expect(isProtected(s)).toBe(true);
    expect(isBlocked(s, "insertRows")).toBe(false);
    expect(isBlocked(s, "deleteRows")).toBe(true);
    expect(wb.protection?.structure).toBe(true);
    expect(canEditCell(s, 2, 2)).toBe(true);
    expect(canEditCell(s, 2, 1)).toBe(false);
  });
});
