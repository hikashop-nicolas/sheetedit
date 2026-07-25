import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { canEditCell, isBlocked, isProtected } from "../../core/protection";
import { setOdsCellStyle } from "./styles";

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"` +
  ` xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
  ` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"` +
  ` xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"`;

const cell = (t: string, style?: string): string =>
  `<table:table-cell office:value-type="string"${style ? ` table:style-name="${style}"` : ""}><text:p>${t}</text:p></table:table-cell>`;

/** A one-sheet ods; `tableAttrs` / `tableHead` go on the table, `bodyAttrs` on office:spreadsheet. */
function ods(opts: { tableAttrs?: string; tableHead?: string; bodyAttrs?: string; styles?: string } = {}): Record<string, Uint8Array> {
  const content =
    `<?xml version="1.0"?><office:document-content ${NS}>` +
    `<office:automatic-styles>${opts.styles ?? ""}</office:automatic-styles>` +
    `<office:body><office:spreadsheet${opts.bodyAttrs ? " " + opts.bodyAttrs : ""}>` +
    `<table:table table:name="Sheet1"${opts.tableAttrs ? " " + opts.tableAttrs : ""}>${opts.tableHead ?? ""}` +
    `<table:table-row>${cell("a")}${cell("b", "ceOpen")}</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return {
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  };
}

/** An automatic cell style that leaves the cell unlocked. */
const OPEN_STYLE = `<style:style style:name="ceOpen" style:family="table-cell"><style:table-cell-properties style:cell-protect="none"/></style:style>`;
const content = (b: Uint8Array): string => strFromU8(unzipSync(b)["content.xml"]!);

describe("ods sheet protection", () => {
  it("reads table:protected", () => {
    const wb = readWorkbook(zipSync(ods({ tableAttrs: `table:protected="true"` })));
    expect(isProtected(wb.sheets[0]!)).toBe(true);
  });

  it("an unprotected table has no protection state", () => {
    expect(readWorkbook(zipSync(ods())).sheets[0]!.protection).toBeUndefined();
  });

  it("inverts the loext permission flags into blocked-action flags", () => {
    // ODF states what is ALLOWED; the model states what is BLOCKED.
    const head = `<loext:table-protection loext:select-unprotected-cells="true" loext:insert-rows="true"/>`;
    const s = readWorkbook(zipSync(ods({ tableAttrs: `table:protected="true"`, tableHead: head }))).sheets[0]!;
    expect(isBlocked(s, "insertRows")).toBe(false); // permission granted
    expect(isBlocked(s, "selectUnlockedCells")).toBe(false);
    expect(isBlocked(s, "deleteRows")).toBe(true); // not granted -> default (blocked)
  });

  it("reads per-cell lock state from style:cell-protect", () => {
    const s = readWorkbook(zipSync(ods({ styles: OPEN_STYLE }))).sheets[0]!;
    expect(s.cells.get("1:1")!.cellStyle?.unlocked).toBeUndefined(); // no style = locked default
    expect(s.cells.get("1:2")!.cellStyle?.unlocked).toBe(true);
  });

  it("reads a hidden formula from the compound cell-protect value", () => {
    const styles = `<style:style style:name="ceOpen" style:family="table-cell"><style:table-cell-properties style:cell-protect="protected formula-hidden"/></style:style>`;
    const st = readWorkbook(zipSync(ods({ styles }))).sheets[0]!.cells.get("1:2")!.cellStyle;
    expect(st?.unlocked).toBeUndefined();
    expect(st?.formulaHidden).toBe(true);
  });

  it("only locked cells are read-only under protection", () => {
    const s = readWorkbook(zipSync(ods({ tableAttrs: `table:protected="true"`, styles: OPEN_STYLE }))).sheets[0]!;
    expect(canEditCell(s, 1, 1)).toBe(false);
    expect(canEditCell(s, 1, 2)).toBe(true);
  });

  it("writes table:protected and the granted permissions", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { protection: { sheet: true, locks: { insertRows: false } }, protectionDirty: true });
    const xml = content(writeWorkbook(wb));
    expect(xml).toMatch(/table:protected="true"/);
    expect(xml).toMatch(/loext:insert-rows="true"/);
    expect(xml).not.toContain("delete-rows"); // blocked -> no permission attribute
  });

  it("puts the extension element first, the only position ODF allows", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { protection: { sheet: true, locks: { insertRows: false } }, protectionDirty: true });
    const xml = content(writeWorkbook(wb));
    expect(xml).toMatch(/<table:table [^>]*>\s*<loext:table-protection/);
  });

  it("clears every protection attribute when protection is lifted", () => {
    const wb = readWorkbook(zipSync(ods({ tableAttrs: `table:protected="true" table:protection-key="Zm9v"`, tableHead: `<loext:table-protection loext:insert-rows="true"/>` })));
    Object.assign(wb.sheets[0]!, { protection: undefined, protectionDirty: true });
    const xml = content(writeWorkbook(wb));
    expect(xml).not.toContain("table:protected");
    expect(xml).not.toContain("protection-key");
    expect(xml).not.toContain("table-protection");
  });

  it("keeps a protection key it did not author", () => {
    const wb = readWorkbook(zipSync(ods({ tableAttrs: `table:protected="true" table:protection-key="Zm9vYmFy" table:protection-key-digest-algorithm="http://www.w3.org/2000/09/xmldsig#sha256"` })));
    const prot = wb.sheets[0]!.protection!;
    expect(prot.password?.hash).toBe("Zm9vYmFy");
    Object.assign(wb.sheets[0]!, { protection: { ...prot, locks: { insertRows: false } }, protectionDirty: true });
    expect(content(writeWorkbook(wb))).toMatch(/table:protection-key="Zm9vYmFy"/);
  });

  it("round-trips what it wrote", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { protection: { sheet: true, locks: { insertRows: false, deleteColumns: false } }, protectionDirty: true });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect(isProtected(back)).toBe(true);
    expect(isBlocked(back, "insertRows")).toBe(false);
    expect(isBlocked(back, "deleteColumns")).toBe(false);
    expect(isBlocked(back, "insertColumns")).toBe(true);
  });

  it("does not touch the protection state of a sheet nobody changed", () => {
    // writeOds always re-serializes content.xml, so this checks the protection markup itself is
    // carried over unchanged rather than comparing bytes.
    const src = zipSync(ods({ tableAttrs: `table:protected="true"`, tableHead: `<loext:table-protection loext:insert-rows="true"/>` }));
    const out = content(writeWorkbook(readWorkbook(src)));
    expect(out).toMatch(/table:protected="true"/);
    expect(out).toMatch(/loext:insert-rows="true"/);
  });
});

describe("ods cell locking", () => {
  it("unlocks a cell and round-trips it", () => {
    const wb = readWorkbook(zipSync(ods()));
    const sheet = wb.sheets[0]!;
    setOdsCellStyle(wb, sheet, sheet.cells.get("1:1")!, { locked: false });
    const out = writeWorkbook(wb);
    expect(content(out)).toMatch(/style:cell-protect="none"/);
    expect(readWorkbook(out).sheets[0]!.cells.get("1:1")!.cellStyle?.unlocked).toBe(true);
  });

  it("an unrelated style change keeps the cell unlocked", () => {
    const wb = readWorkbook(zipSync(ods({ styles: OPEN_STYLE })));
    const sheet = wb.sheets[0]!;
    setOdsCellStyle(wb, sheet, sheet.cells.get("1:2")!, { bold: true });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:2")!;
    expect(back.cellStyle?.bold).toBe(true);
    expect(back.cellStyle?.unlocked).toBe(true);
  });
});

describe("ods workbook protection", () => {
  it("reads table:structure-protected", () => {
    const wb = readWorkbook(zipSync(ods({ bodyAttrs: `table:structure-protected="true"` })));
    expect(wb.protection?.structure).toBe(true);
  });

  it("writes and round-trips a structure lock", () => {
    const wb = readWorkbook(zipSync(ods()));
    wb.protection = { structure: true };
    wb.protectionDirty = true;
    const out = writeWorkbook(wb);
    expect(content(out)).toMatch(/table:structure-protected="true"/);
    expect(readWorkbook(out).protection?.structure).toBe(true);
  });

  it("clears an existing structure lock", () => {
    const wb = readWorkbook(zipSync(ods({ bodyAttrs: `table:structure-protected="true"` })));
    wb.protection = { structure: undefined };
    wb.protectionDirty = true;
    expect(content(writeWorkbook(wb))).not.toContain("structure-protected");
  });
});
