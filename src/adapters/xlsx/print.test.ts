import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../core/workbook";
import { parseHeaderFooter, formatHeaderFooter, paperSizeFor, toggleBreak } from "../../core/print";
import { parseTitles } from "./print-read";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";

function book(sheetExtra = "", wbInner = "", sheetName = "S"): Uint8Array {
  const data = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c></row></sheetData>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}">${wbInner}<sheets><sheet name="${sheetName}" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}">${data}${sheetExtra}</worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${MAIN}"/>`),
  });
}
const sheetXml = (b: Uint8Array): string => strFromU8(unzipSync(b)["xl/worksheets/sheet1.xml"]!);
const wbXml = (b: Uint8Array): string => strFromU8(unzipSync(b)["xl/workbook.xml"]!);

describe("header/footer field syntax", () => {
  it("splits the three regions", () => {
    expect(parseHeaderFooter("&LLeft&CMiddle&RRight")).toEqual({ left: "Left", center: "Middle", right: "Right" });
  });

  it("treats a string with no region code as the centre, which is how Excel reads it", () => {
    expect(parseHeaderFooter("Just this")).toEqual({ center: "Just this" });
  });

  it("keeps field and formatting codes in the region text so they round-trip", () => {
    const hf = parseHeaderFooter(`&C&"Arial,Bold"&12Page &P of &N`);
    expect(hf).toEqual({ center: `&"Arial,Bold"&12Page &P of &N` });
    expect(formatHeaderFooter(hf)).toBe(`&C&"Arial,Bold"&12Page &P of &N`);
  });

  it("does not mistake a doubled ampersand for a region switch", () => {
    expect(parseHeaderFooter("&LR&&D")).toEqual({ left: "R&&D" });
  });

  it("round-trips an empty setting as undefined", () => {
    expect(parseHeaderFooter("")).toBeUndefined();
    expect(formatHeaderFooter(undefined)).toBeUndefined();
  });
});

describe("print helpers", () => {
  it("matches a paper size from its dimensions either way round", () => {
    expect(paperSizeFor(210, 297)).toBe(9);
    expect(paperSizeFor(297, 210)).toBe(9); // landscape states the same paper turned round
    expect(paperSizeFor(999, 999)).toBeUndefined();
  });

  it("toggles a break and keeps the list sorted", () => {
    expect(toggleBreak([5], 3)).toEqual([3, 5]);
    expect(toggleBreak([3, 5], 3)).toEqual([5]);
  });

  it("parses print titles for rows, columns, or both", () => {
    expect(parseTitles("S!$1:$2")).toEqual({ rows: { from: 1, to: 2 } });
    expect(parseTitles("S!$A:$B")).toEqual({ cols: { from: 1, to: 2 } });
    expect(parseTitles("S!$A:$A,S!$1:$1")).toEqual({ cols: { from: 1, to: 1 }, rows: { from: 1, to: 1 } });
  });
});

describe("xlsx print setup", () => {
  it("reads the page setup elements", () => {
    const wb = readWorkbook(book(
      `<printOptions gridLines="true" headings="true" horizontalCentered="true"/>` +
      `<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
      `<pageSetup paperSize="9" orientation="landscape" scale="80" fitToWidth="1" fitToHeight="0" pageOrder="overThenDown"/>`));
    const p = wb.sheets[0]!.printSetup!;
    expect(p.orientation).toBe("landscape");
    expect(p.paperSize).toBe(9);
    expect(p.scale).toBe(80);
    expect(p.pageOrder).toBe("overThenDown");
    expect(p.gridLines).toBe(true);
    expect(p.margins).toEqual({ left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });
  });

  it("ignores a first page number the sheet does not opt into", () => {
    const on = readWorkbook(book(`<pageSetup firstPageNumber="4" useFirstPageNumber="true"/>`)).sheets[0]!.printSetup!;
    const off = readWorkbook(book(`<pageSetup firstPageNumber="4"/>`)).sheets[0]!.printSetup;
    expect(on.firstPageNumber).toBe(4);
    expect(off?.firstPageNumber).toBeUndefined();
  });

  it("counts a break as the line that starts the new page", () => {
    // The file stores the break 0-based, so id=2 is the break above row 3.
    const p = readWorkbook(book(`<rowBreaks count="1"><brk id="2" man="1"/></rowBreaks>`)).sheets[0]!.printSetup!;
    expect(p.rowBreaks).toEqual([3]);
  });

  it("skips a break the application computed rather than the user", () => {
    const p = readWorkbook(book(`<rowBreaks count="2"><brk id="2" man="1"/><brk id="9" man="0"/></rowBreaks>`)).sheets[0]!.printSetup!;
    expect(p.rowBreaks).toEqual([3]);
  });

  it("reads the print area and titles from the sheet-scoped names", () => {
    const names = `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">S!$A$1:$C$5</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">S!$1:$2</definedName></definedNames>`;
    const p = readWorkbook(book("", names)).sheets[0]!.printSetup!;
    expect(p.printArea).toEqual([{ r1: 1, c1: 1, r2: 5, c2: 3 }]);
    expect(p.titleRows).toEqual({ from: 1, to: 2 });
  });

  it("a sheet with no print elements has no print setup", () => {
    expect(readWorkbook(book()).sheets[0]!.printSetup).toBeUndefined();
  });

  it("writes the elements in the schema's order", () => {
    const wb = readWorkbook(book());
    Object.assign(wb.sheets[0]!, {
      printSetup: { orientation: "landscape", margins: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 }, gridLines: true, header: { center: "Hi" }, rowBreaks: [3] },
      printDirty: true,
    });
    const xml = sheetXml(writeWorkbook(wb));
    const at = (tag: string) => xml.indexOf(`<${tag}`);
    expect(at("printOptions")).toBeGreaterThan(xml.indexOf("</sheetData>"));
    expect(at("pageMargins")).toBeGreaterThan(at("printOptions"));
    expect(at("pageSetup")).toBeGreaterThan(at("pageMargins"));
    expect(at("headerFooter")).toBeGreaterThan(at("pageSetup"));
    expect(at("rowBreaks")).toBeGreaterThan(at("headerFooter"));
  });

  it("writes fitToPage onto sheetPr, where the fit-to counts take effect", () => {
    const wb = readWorkbook(book());
    Object.assign(wb.sheets[0]!, { printSetup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0 }, printDirty: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/<sheetPr><pageSetUpPr fitToPage="true"\/><\/sheetPr>/);
    expect(xml).toMatch(/fitToWidth="1"/);
  });

  it("writes the print area and titles as sheet-scoped names", () => {
    const wb = readWorkbook(book());
    Object.assign(wb.sheets[0]!, { printSetup: { printArea: [{ r1: 1, c1: 1, r2: 5, c2: 3 }], titleRows: { from: 1, to: 1 } }, printDirty: true });
    const xml = wbXml(writeWorkbook(wb));
    expect(xml).toContain(`name="_xlnm.Print_Area" localSheetId="0"`);
    expect(xml).toContain("S!$A$1:$C$5");
    expect(xml).toContain("S!$1:$1");
  });

  it("quotes a sheet name that is not a bare identifier", () => {
    const wb = readWorkbook(book("", "", "My Sheet"));
    Object.assign(wb.sheets[0]!, { printSetup: { printArea: [{ r1: 1, c1: 1, r2: 2, c2: 2 }] }, printDirty: true });
    expect(wbXml(writeWorkbook(wb))).toContain("'My Sheet'!$A$1:$B$2");
  });

  it("drops the print names when the area is cleared", () => {
    const names = `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">S!$A$1:$C$5</definedName></definedNames>`;
    const wb = readWorkbook(book("", names));
    Object.assign(wb.sheets[0]!, { printSetup: {}, printDirty: true });
    expect(wbXml(writeWorkbook(wb))).not.toContain("Print_Area");
  });

  it("leaves a sheet untouched when the print setup was not changed", () => {
    const src = book(`<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="landscape"/>`);
    expect(sheetXml(writeWorkbook(readWorkbook(src)))).toBe(sheetXml(src));
  });

  it("keeps attributes it does not model when patching pageSetup", () => {
    const wb = readWorkbook(book(`<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>`));
    Object.assign(wb.sheets[0]!, { printSetup: { ...wb.sheets[0]!.printSetup, orientation: "landscape" }, printDirty: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/horizontalDpi="600"/);
    expect(xml).toMatch(/copies="3"/);
    expect(xml).toMatch(/orientation="landscape"/);
  });

  it("round-trips everything it wrote", () => {
    const wb = readWorkbook(book());
    const setup = {
      orientation: "portrait" as const, paperSize: 11, scale: 75,
      margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
      gridLines: true, headings: false, horizontalCentered: true, verticalCentered: true,
      header: { left: "L", right: "&P" }, footer: { center: "F" },
      printArea: [{ r1: 2, c1: 2, r2: 8, c2: 4 }], titleCols: { from: 1, to: 2 },
      rowBreaks: [4, 9], colBreaks: [3],
      firstPageNumber: 7,
    };
    Object.assign(wb.sheets[0]!, { printSetup: setup, printDirty: true });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.printSetup!;
    expect(back).toMatchObject(setup);
  });
});
