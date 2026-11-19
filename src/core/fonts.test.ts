import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "./workbook";
import { workbookFonts } from "./fonts";
import type { Workbook } from "./model";

// The picker's "In this document" group: every font the workbook names, from its styles, its
// rich-text runs and its theme, without parsing a sheet only to list fonts.

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const book = (): Uint8Array =>
  zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}"><sheetData><row r="1"><c r="A1" s="1" t="str"><v>x</v></c><c r="B1" t="inlineStr"><is><r><t>a</t></r><r><rPr><rFont val="Klee One"/></rPr><t>b</t></r></is></c></row></sheetData></worksheet>`),
    "xl/worksheets/sheet2.xml": strToU8(`<worksheet xmlns="${SS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><rFont val="Lobster"/></rPr><t>c</t></r><r><t>d</t></r></is></c></row></sheetData></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${SS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Meiryo"/></font></fonts><cellXfs count="2"><xf fontId="0"/><xf fontId="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`),
  });

describe("the fonts a workbook names", () => {
  it("lists the Normal font, the style pool and rich-text runs", () => {
    const fonts = workbookFonts(readWorkbook(book()));
    expect(fonts).toEqual(expect.arrayContaining(["Meiryo", "Klee One", "Lobster"]));
  });

  it("leaves an unparsed sheet alone", () => {
    const wb = readWorkbook(book(), { lazySheets: true });
    const fonts = workbookFonts(wb);
    expect(fonts).toContain("Meiryo"); // style pool, known up front
    expect(fonts).not.toContain("Lobster"); // only in sheet B, not parsed yet
  });

  it("adds the theme's heading and body fonts, once each", () => {
    const wb = { sheets: [], defaultFontName: "Aptos Narrow", theme: { majorFont: "Aptos Display", minorFont: "aptos narrow" } } as unknown as Workbook;
    expect(workbookFonts(wb)).toEqual(["Aptos Narrow", "Aptos Display"]);
  });
});
