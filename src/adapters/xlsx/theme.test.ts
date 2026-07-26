import { describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../core/workbook";
import { BUILTIN_THEMES, applyTint, resolveThemeRef, setWorkbookTheme, THEME_INDEX_ORDER, usesTheme } from "../../core/theme";
import { readXlsxTheme } from "./theme-read";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";

const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${A}" name="Office Theme"><a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements></a:theme>`;

// A1 uses the accent1 fill + a theme-coloured font, B1 an explicit red, C1 a tinted accent1.
const STYLES = `<styleSheet xmlns="${MAIN}">
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/><scheme val="minor"/></font>
<font><sz val="11"/><color theme="4"/><name val="Calibri"/><scheme val="minor"/></font>
<font><sz val="11"/><color rgb="FFFF0000"/><name val="Arial"/></font>
</fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor theme="5"/></patternFill></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

function book(opts: { theme?: string | null } = {}): Uint8Array {
  const data = `<sheetData><row r="1">` +
    `<c r="A1" s="1" t="inlineStr"><is><t>themed</t></is></c>` +
    `<c r="B1" s="2" t="inlineStr"><is><t>explicit</t></is></c>` +
    `<c r="C1" s="3" t="inlineStr"><is><t>plain</t></is></c>` +
    `</row></sheetData>`;
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(`<Types xmlns="${CT}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${opts.theme === null ? "" : `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`}</Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/>${opts.theme === null ? "" : `<Relationship Id="rId3" Type="${R}/theme" Target="theme/theme1.xml"/>`}</Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}">${data}</worksheet>`),
    "xl/styles.xml": strToU8(STYLES),
  };
  if (opts.theme !== null) files["xl/theme/theme1.xml"] = strToU8(opts.theme ?? THEME_XML);
  return zipSync(files);
}
const themeXml = (b: Uint8Array): string => strFromU8(unzipSync(b)["xl/theme/theme1.xml"]!);
const stylesXml = (b: Uint8Array): string => strFromU8(unzipSync(b)["xl/styles.xml"]!);

describe("theme model", () => {
  it("indexes the palette the way a <color theme=N> does, not in clrScheme order", () => {
    // The first two pairs are swapped in the index order; getting this wrong swaps black and white.
    expect(THEME_INDEX_ORDER[0]).toBe("lt1");
    expect(THEME_INDEX_ORDER[1]).toBe("dk1");
    const office = BUILTIN_THEMES[0]!;
    expect(resolveThemeRef({ index: 0 }, office)).toBe("#ffffff");
    expect(resolveThemeRef({ index: 1 }, office)).toBe("#000000");
    expect(resolveThemeRef({ index: 4 }, office)).toBe("#4472c4"); // accent1
  });

  it("applies a tint the way Excel does", () => {
    expect(applyTint("#808080", 0)).toBe("#808080");
    expect(applyTint("#808080", -1)).toBe("#000000"); // fully darkened
    expect(applyTint("#808080", 1)).toBe("#ffffff"); // fully lightened
  });
});

describe("xlsx theme", () => {
  it("reads the palette and the scheme fonts", () => {
    const theme = readXlsxTheme(strToU8(THEME_XML));
    expect(theme.name).toBe("Office");
    expect(theme.colors.accent1).toBe("#4472c4");
    expect(theme.colors.dk1).toBe("#000000"); // from the sysClr's lastClr
    expect(theme.colors.lt1).toBe("#ffffff");
    expect(theme.majorFont).toBe("Calibri Light");
    expect(theme.minorFont).toBe("Calibri");
  });

  it("falls back to Office when the workbook ships no theme", () => {
    expect(readXlsxTheme(undefined).colors.accent1).toBe(BUILTIN_THEMES[0]!.colors.accent1);
  });

  it("resolves theme colours onto the cells and remembers the reference", () => {
    const wb = readWorkbook(book());
    const a1 = wb.sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.color).toBe("#4472c4"); // font colour theme=4 -> accent1
    expect(a1.colorRef).toEqual({ index: 4 });
    expect(a1.bg).toBe("#ed7d31"); // fill theme=5 -> accent2
    expect(a1.bgRef).toEqual({ index: 5 });
    expect(a1.fontScheme).toBe("minor");
  });

  it("leaves an explicit colour with no theme reference", () => {
    const b1 = readWorkbook(book()).sheets[0]!.cells.get("1:2")!.cellStyle!;
    expect(b1.color).toBe("#ff0000");
    expect(b1.colorRef).toBeUndefined();
    expect(usesTheme(b1)).toBe(false);
  });

  it("recolours themed cells on a switch and leaves explicit ones alone", () => {
    const wb = readWorkbook(book());
    const sheet = wb.sheets[0]!;
    setWorkbookTheme(wb, BUILTIN_THEMES.find((t) => t.name === "Office 2007-2010")!);
    expect(sheet.cells.get("1:1")!.cellStyle!.color).toBe("#4f81bd"); // that theme's accent1
    expect(sheet.cells.get("1:1")!.cellStyle!.bg).toBe("#c0504d");
    expect(sheet.cells.get("1:2")!.cellStyle!.color).toBe("#ff0000"); // untouched
  });

  it("follows the theme's body font on a cell tagged with the minor scheme", () => {
    const wb = readWorkbook(book());
    setWorkbookTheme(wb, BUILTIN_THEMES.find((t) => t.name === "Berlin")!);
    expect(wb.sheets[0]!.cells.get("1:1")!.cellStyle!.fontFamily).toBe("Trebuchet MS");
  });

  it("writes the new palette into theme1.xml and leaves styles.xml alone", () => {
    const wb = readWorkbook(book());
    const before = stylesXml(book());
    setWorkbookTheme(wb, BUILTIN_THEMES.find((t) => t.name === "Berlin")!);
    const out = writeWorkbook(wb);
    const xml = themeXml(out);
    expect(xml).toMatch(/val="E97B1F"/i); // Berlin accent1
    expect(xml).toMatch(/typeface="Trebuchet MS"/);
    expect(xml).toMatch(/name="Berlin"/);
    // The cells still reference the palette by index, so the style pool must not change.
    expect(stylesXml(out)).toBe(before);
  });

  it("keeps the sysClr shape for dk1 / lt1 rather than freezing them to a literal", () => {
    const wb = readWorkbook(book());
    setWorkbookTheme(wb, BUILTIN_THEMES.find((t) => t.name === "Grayscale")!);
    const xml = themeXml(writeWorkbook(wb));
    expect(xml).toMatch(/<a:dk1><a:sysClr val="windowText" lastClr="000000"\/><\/a:dk1>/);
  });

  it("preserves the parts of the theme it does not model", () => {
    const wb = readWorkbook(book());
    setWorkbookTheme(wb, BUILTIN_THEMES[1]!);
    // fmtScheme drives chart and shape presets; regenerating the file would drop it.
    expect(themeXml(writeWorkbook(wb))).toContain("fmtScheme");
  });

  it("round-trips a switched theme", () => {
    const wb = readWorkbook(book());
    setWorkbookTheme(wb, BUILTIN_THEMES.find((t) => t.name === "Ion")!);
    const back = readWorkbook(writeWorkbook(wb));
    expect(back.theme!.colors.accent1).toBe("#b01513");
    expect(back.sheets[0]!.cells.get("1:1")!.cellStyle!.color).toBe("#b01513");
  });

  it("creates a theme part for a workbook that had none", () => {
    const wb = readWorkbook(book({ theme: null }));
    expect(wb.files["xl/theme/theme1.xml"]).toBeUndefined();
    setWorkbookTheme(wb, BUILTIN_THEMES.find((t) => t.name === "Slice")!);
    const out = writeWorkbook(wb);
    expect(themeXml(out)).toMatch(/val="052F61"/i);
    // It has to be reachable, or Excel ignores it.
    expect(strFromU8(unzipSync(out)["[Content_Types].xml"]!)).toContain("theme+xml");
    expect(strFromU8(unzipSync(out)["xl/_rels/workbook.xml.rels"]!)).toContain("theme/theme1.xml");
    expect(readWorkbook(out).theme!.colors.accent1).toBe("#052f61");
  });

  it("leaves the theme part untouched when nothing switched it", () => {
    const src = book();
    expect(themeXml(writeWorkbook(readWorkbook(src)))).toBe(themeXml(src));
  });

  it("writes the demo fixture", () => {
    if (!process.env.SHEETEDIT_WRITE_FIXTURES) return;
    writeFileSync("demo/c-theme.xlsx", book());
  });
});
