import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";
import { SHAPE_GALLERY } from "./shape-geom";

// The toolbar's dropdowns. Font and size used to be native <select>s, whose list the browser draws
// in its own layer: they looked nothing like the menus beside them, and they were the only two that
// never collided with the cell float bar - which is why the collision on the others went unnoticed.

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

const frame = () => new Promise<void>((res) => setTimeout(res, 40));
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** A minimal xlsx, so the full styling toolbar is built (a csv has none). */
const book = (): Uint8Array =>
  zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}"><sheetData><row r="1"><c r="A1" t="str"><v>x</v></c></row></sheetData></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${SS}"/>`),
  });

async function mount(): Promise<HTMLElement> {
  const host = document.createElement("div");
  document.body.appendChild(host);
  createSheetEditor(host, book(), { filename: "s.xlsx" });
  await frame();
  return host;
}

const menus = (host: HTMLElement): HTMLElement[] =>
  [...host.querySelectorAll<HTMLElement>(".sheetedit-tb-groupmenu")];
const openOne = (host: HTMLElement): HTMLElement | undefined => menus(host).find((m) => !m.hidden);
const button = (host: HTMLElement, label: string): HTMLElement =>
  [...host.querySelectorAll<HTMLElement>(".sheetedit-btn")].find((b) => b.textContent?.startsWith(label))!;

describe("the toolbar's dropdowns", () => {
  it("are all menus of our own, so they all look and stack the same", async () => {
    const host = await mount();
    expect(host.querySelectorAll("select").length, "no native dropdown left").toBe(0);
    // The font, size, number-format and formatting buttons each own one.
    expect(menus(host).length).toBeGreaterThanOrEqual(4);
    host.remove();
  });

  it("opens one at a time", async () => {
    const host = await mount();
    button(host, "Aa").click();
    expect(openOne(host), "the font menu opened").toBeTruthy();
    const first = openOne(host);
    button(host, "123").click();
    expect(openOne(host), "the format menu opened").toBeTruthy();
    expect(openOne(host), "and the font menu closed").not.toBe(first);
    host.remove();
  });

  it("lists the fonts and sizes as items, not as options", async () => {
    const host = await mount();
    button(host, "Aa").click();
    const items = [...openOne(host)!.querySelectorAll("button")];
    expect(items.map((b) => b.textContent)).toContain("Georgia");
    host.remove();
  });
});

describe("the shape gallery", () => {
  it("offers every geometry as a preview, under a translated heading", () => {
    const total = SHAPE_GALLERY.reduce((n, c) => n + c.geoms.length, 0);
    expect(total).toBeGreaterThan(50);
    expect(SHAPE_GALLERY.every((c) => c.label.startsWith("shapeCat"))).toBe(true);
  });
});
