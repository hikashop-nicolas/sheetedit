import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../core/workbook";
import { canEditCell, isBlocked, isProtected } from "../../core/protection";
import { setXlsxCellStyle } from "./styles";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";

function book(sheetInner: string, styles?: string, wbInner?: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}">${wbInner ?? ""}<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}">${sheetInner}</worksheet>`),
    "xl/styles.xml": strToU8(styles ?? `<styleSheet xmlns="${MAIN}"/>`),
  });
}

const DATA = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>u</t></is></c></row></sheetData>`;
// xf 0 is the default (locked); xf 1 unlocks the cell that points at it.
const STYLES = `<styleSheet xmlns="${MAIN}"><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0"/></xf></cellXfs></styleSheet>`;
const sheetXml = (b: Uint8Array): string => strFromU8(unzipSync(b)["xl/worksheets/sheet1.xml"]!);
const wbXml = (b: Uint8Array): string => strFromU8(unzipSync(b)["xl/workbook.xml"]!);
const stylesXml = (b: Uint8Array): string => strFromU8(unzipSync(b)["xl/styles.xml"]!);

describe("xlsx sheet protection", () => {
  it("reads the protected flag", () => {
    const wb = readWorkbook(book(`${DATA}<sheetProtection sheet="1" objects="1" scenarios="1"/>`));
    expect(isProtected(wb.sheets[0]!)).toBe(true);
  });

  it("an unprotected sheet has no protection state", () => {
    expect(readWorkbook(book(DATA)).sheets[0]!.protection).toBeUndefined();
  });

  it("only records the flags the file states, so the rest keep their spec defaults", () => {
    const s = readWorkbook(book(`${DATA}<sheetProtection sheet="1" insertRows="0"/>`)).sheets[0]!;
    expect(s.protection!.locks).toEqual({ insertRows: false });
    // Stated: inserting rows is allowed. Unstated: formatting defaults to blocked.
    expect(isBlocked(s, "insertRows")).toBe(false);
    expect(isBlocked(s, "formatCells")).toBe(true);
    // selectLockedCells defaults the other way: selecting is allowed unless stated.
    expect(isBlocked(s, "selectLockedCells")).toBe(false);
  });

  it("blocks nothing while the sheet is not protected, whatever the flags say", () => {
    const s = readWorkbook(book(`${DATA}<sheetProtection sheet="0" insertRows="1"/>`)).sheets[0]!;
    expect(isBlocked(s, "insertRows")).toBe(false);
  });

  it("reads per-cell lock state from the style, defaulting to locked", () => {
    const wb = readWorkbook(book(DATA, STYLES));
    const s = wb.sheets[0]!;
    expect(s.cells.get("1:1")!.cellStyle?.unlocked).toBeUndefined();
    expect(s.cells.get("1:2")!.cellStyle?.unlocked).toBe(true);
  });

  it("only locked cells are read-only, and only while the sheet is protected", () => {
    const wb = readWorkbook(book(`${DATA}<sheetProtection sheet="1"/>`, STYLES));
    const s = wb.sheets[0]!;
    expect(canEditCell(s, 1, 1)).toBe(false); // locked (the default)
    expect(canEditCell(s, 1, 2)).toBe(true); // explicitly unlocked
    expect(canEditCell(s, 5, 5)).toBe(false); // a blank cell inherits the locked default
    s.protection = undefined;
    expect(canEditCell(s, 1, 1)).toBe(true);
  });

  it("writes only the flags that differ from the spec default", () => {
    const wb = readWorkbook(book(DATA));
    Object.assign(wb.sheets[0]!, { protection: { sheet: true, locks: { objects: true, scenarios: true, insertRows: false } }, protectionDirty: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/<sheetProtection[^>]*sheet="1"/);
    expect(xml).toMatch(/objects="1"/);
    expect(xml).toMatch(/insertRows="0"/);
    expect(xml).not.toContain("formatCells"); // default (blocked) stays implicit
    expect(xml).not.toContain("sort=");
  });

  it("places sheetProtection after sheetData, where the schema requires it", () => {
    const wb = readWorkbook(book(DATA));
    Object.assign(wb.sheets[0]!, { protection: { sheet: true }, protectionDirty: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml.indexOf("<sheetProtection")).toBeGreaterThan(xml.indexOf("</sheetData>"));
  });

  it("removes the element when protection is lifted", () => {
    const wb = readWorkbook(book(`${DATA}<sheetProtection sheet="1" objects="1"/>`));
    Object.assign(wb.sheets[0]!, { protection: undefined, protectionDirty: true });
    expect(sheetXml(writeWorkbook(wb))).not.toContain("sheetProtection");
  });

  it("keeps a password hash it did not author", () => {
    const src = `${DATA}<sheetProtection sheet="1" algorithmName="SHA-512" hashValue="AbC=" saltValue="XyZ=" spinCount="100000"/>`;
    const wb = readWorkbook(book(src));
    const prot = wb.sheets[0]!.protection!;
    expect(prot.password).toEqual({ hash: "AbC=", algorithmName: "SHA-512", saltValue: "XyZ=", spinCount: "100000" });
    // Re-protecting with different allowances must not silently drop the hash.
    Object.assign(wb.sheets[0]!, { protection: { ...prot, locks: { sort: false } }, protectionDirty: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/hashValue="AbC="/);
    expect(xml).toMatch(/spinCount="100000"/);
  });

  it("leaves the sheet untouched when protection was not changed", () => {
    const src = book(`${DATA}<sheetProtection sheet="1" objects="1" scenarios="1"/>`);
    expect(sheetXml(writeWorkbook(readWorkbook(src)))).toBe(sheetXml(src));
  });

  it("round-trips what it wrote", () => {
    const wb = readWorkbook(book(DATA));
    Object.assign(wb.sheets[0]!, { protection: { sheet: true, locks: { insertRows: false, sort: false } }, protectionDirty: true });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect(back.protection!.sheet).toBe(true);
    expect(isBlocked(back, "insertRows")).toBe(false);
    expect(isBlocked(back, "sort")).toBe(false);
    expect(isBlocked(back, "deleteRows")).toBe(true);
  });
});

describe("xlsx cell locking", () => {
  it("unlocks a cell through the style pool", () => {
    const wb = readWorkbook(book(DATA));
    const sheet = wb.sheets[0]!;
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:1")!, { locked: false });
    const out = writeWorkbook(wb);
    expect(stylesXml(out)).toMatch(/<protection locked="0"\/>/);
    expect(readWorkbook(out).sheets[0]!.cells.get("1:1")!.cellStyle?.unlocked).toBe(true);
  });

  it("an unrelated style change does not silently re-lock an unlocked cell", () => {
    // The xf is rebuilt from scratch on every style change, so the lock state has to be carried
    // across or protecting the sheet would suddenly freeze a cell the user had unlocked.
    const wb = readWorkbook(book(DATA, STYLES));
    const sheet = wb.sheets[0]!;
    const cell = sheet.cells.get("1:2")!;
    expect(cell.cellStyle?.unlocked).toBe(true);
    setXlsxCellStyle(wb, sheet, cell, { bold: true });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:2")!;
    expect(back.cellStyle?.bold).toBe(true);
    expect(back.cellStyle?.unlocked).toBe(true);
  });

  it("re-locks a cell that was unlocked", () => {
    const wb = readWorkbook(book(DATA, STYLES));
    const sheet = wb.sheets[0]!;
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:2")!, { locked: true });
    expect(readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:2")!.cellStyle?.unlocked).toBeUndefined();
  });

  it("preserves a hidden-formula flag across an unrelated style change", () => {
    const styles = `<styleSheet xmlns="${MAIN}"><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection hidden="1"/></xf></cellXfs></styleSheet>`;
    const wb = readWorkbook(book(DATA, styles));
    const sheet = wb.sheets[0]!;
    expect(sheet.cells.get("1:2")!.cellStyle?.formulaHidden).toBe(true);
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:2")!, { italic: true });
    expect(readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:2")!.cellStyle?.formulaHidden).toBe(true);
  });
});

describe("xlsx workbook protection", () => {
  it("reads a locked structure", () => {
    const wb = readWorkbook(book(DATA, undefined, `<workbookProtection lockStructure="1"/>`));
    expect(wb.protection).toEqual({ structure: true });
  });

  it("ignores an empty element, which locks nothing", () => {
    expect(readWorkbook(book(DATA, undefined, `<workbookProtection/>`)).protection).toBeUndefined();
  });

  it("writes the element before bookViews, where the schema requires it", () => {
    const wb = readWorkbook(book(DATA, undefined, `<bookViews><workbookView/></bookViews>`));
    wb.protection = { structure: true };
    wb.protectionDirty = true;
    const xml = wbXml(writeWorkbook(wb));
    expect(xml.indexOf("<workbookProtection")).toBeLessThan(xml.indexOf("<bookViews>"));
    expect(xml).toMatch(/lockStructure="1"/);
  });

  it("round-trips a structure lock it wrote", () => {
    const wb = readWorkbook(book(DATA));
    wb.protection = { structure: true };
    wb.protectionDirty = true;
    expect(readWorkbook(writeWorkbook(wb)).protection).toEqual({ structure: true });
  });

  it("clears an existing lock", () => {
    const wb = readWorkbook(book(DATA, undefined, `<workbookProtection lockStructure="1"/>`));
    wb.protection = { structure: undefined };
    wb.protectionDirty = true;
    expect(wbXml(writeWorkbook(wb))).not.toContain("workbookProtection");
  });
});
