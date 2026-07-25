import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";
import { canEditCell, canEditRange, isBlocked, isProtected, isStructureLocked, SHEET_LOCK_DEFAULTS } from "./protection";
import type { Sheet, Workbook } from "./model";

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

const frame = () => new Promise<void>((res) => setTimeout(res, 30));

const sheetOf = (over: Partial<Sheet> = {}): Sheet => ({ name: "S", cells: new Map(), maxRow: 0, maxCol: 0, ...over });

describe("protection gating", () => {
  it("blocks nothing when the sheet is not protected", () => {
    const s = sheetOf();
    expect(isProtected(s)).toBe(false);
    expect(canEditCell(s, 1, 1)).toBe(true);
    for (const flag of Object.keys(SHEET_LOCK_DEFAULTS) as (keyof typeof SHEET_LOCK_DEFAULTS)[]) expect(isBlocked(s, flag)).toBe(false);
  });

  it("a protected sheet locks its cells by default", () => {
    const s = sheetOf({ protection: { sheet: true } });
    expect(canEditCell(s, 1, 1)).toBe(false);
    expect(canEditCell(s, 999, 999)).toBe(false); // blank cells inherit the default too
  });

  it("an explicitly unlocked cell stays editable under protection", () => {
    const s = sheetOf({ protection: { sheet: true } });
    s.cells.set("1:1", { row: 1, col: 1, value: "", kind: "blank", cellStyle: { unlocked: true } });
    expect(canEditCell(s, 1, 1)).toBe(true);
    expect(canEditCell(s, 1, 2)).toBe(false);
  });

  it("a range is editable only when every cell in it is", () => {
    const s = sheetOf({ protection: { sheet: true } });
    for (let c = 1; c <= 3; c++) s.cells.set(`1:${c}`, { row: 1, col: c, value: "", kind: "blank", cellStyle: { unlocked: true } });
    expect(canEditRange(s, { r1: 1, c1: 1, r2: 1, c2: 3 })).toBe(true);
    expect(canEditRange(s, { r1: 1, c1: 1, r2: 2, c2: 3 })).toBe(false); // row 2 is locked
  });

  it("a stated flag wins over its default, in both directions", () => {
    const s = sheetOf({ protection: { sheet: true, locks: { insertRows: false, selectLockedCells: true } } });
    expect(isBlocked(s, "insertRows")).toBe(false); // default true, stated false
    expect(isBlocked(s, "selectLockedCells")).toBe(true); // default false, stated true
    expect(isBlocked(s, "deleteRows")).toBe(true); // unstated, default true
  });

  it("workbook structure locking is independent of any sheet", () => {
    const wb = { kind: "xlsx", sheets: [], files: {} } as Workbook;
    expect(isStructureLocked(wb)).toBe(false);
    wb.protection = { structure: true };
    expect(isStructureLocked(wb)).toBe(true);
  });
});

// --- editor enforcement -------------------------------------------------------

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** A two-cell sheet where B1 is unlocked (xf 1) and A1 keeps the locked default. */
function protectedBook(protection = `<sheetProtection sheet="1"/>`): Uint8Array {
  const data = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>b</t></is></c></row></sheetData>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}">${data}${protection}</worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${MAIN}"><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0"/></xf></cellXfs></styleSheet>`),
  });
}

const mount = async (bytes: Uint8Array) => {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const ed = createSheetEditor(host, bytes, { fileName: "p.xlsx" });
  await frame();
  return { host, ed };
};
const inputAt = (host: HTMLElement, rc: string): HTMLInputElement =>
  host.querySelector<HTMLInputElement>(`td[data-rc="${rc}"] input`)!;

describe("editor enforcement of protection", () => {
  it("makes a locked cell read-only but leaves an unlocked one editable", async () => {
    const { host, ed } = await mount(protectedBook());
    expect(inputAt(host, "1:1").readOnly).toBe(true);
    expect(inputAt(host, "1:2").readOnly).toBe(false);
    ed.destroy();
    host.remove();
  });

  it("refuses a programmatic write to a locked cell and explains why", async () => {
    const { host, ed } = await mount(protectedBook());
    ed.setCellValue("A1", "nope");
    await frame();
    expect(ed.getCellValue("A1")).toBe("a");
    const notice = host.querySelector<HTMLElement>(".sheetedit-notice")!;
    expect(notice.hidden).toBe(false);
    expect(notice.textContent).toBeTruthy();
    ed.destroy();
    host.remove();
  });

  it("still writes to an unlocked cell on a protected sheet", async () => {
    const { host, ed } = await mount(protectedBook());
    ed.setCellValue("B1", "yes");
    await frame();
    expect(ed.getCellValue("B1")).toBe("yes");
    ed.destroy();
    host.remove();
  });

  it("leaves every cell editable when the sheet carries no protection", async () => {
    const { host, ed } = await mount(protectedBook(""));
    expect(inputAt(host, "1:1").readOnly).toBe(false);
    ed.setCellValue("A1", "fine");
    await frame();
    expect(ed.getCellValue("A1")).toBe("fine");
    ed.destroy();
    host.remove();
  });
});
