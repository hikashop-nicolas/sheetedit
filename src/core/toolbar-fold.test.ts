import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";

// How the toolbar folds on a narrow screen. Measured in jsdom by giving every visible control a
// fixed width, since jsdom lays nothing out.
//
// The bug: when the authoring controls had all folded into "⋯" and the row still overflowed, the
// style cluster collapsed into "Aa", which freed most of the row, but the authoring controls stayed
// folded. On a phone the toolbar showed a handful of buttons, "Aa", "⋯" and empty space.

const observers: (() => void)[] = [];
beforeAll(() => {
  (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
    constructor(private cb: () => void) { observers.push(() => this.cb()); }
    observe() {}
    unobserve() {}
    disconnect() {}
  };
});

const frame = () => new Promise<void>((res) => setTimeout(res, 40));
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const book = (): Uint8Array =>
  zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}"><sheetData><row r="1"><c r="A1" t="str"><v>x</v></c></row></sheetData></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${SS}"/>`),
  });

const W = 40; // every visible control is this wide
const shown = (el: Element): boolean => (el as HTMLElement).style.display !== "none" && !(el as HTMLElement).hidden;
/** A control's width: a group (the style slot) is as wide as its visible children. */
const widthOf = (el: Element): number => {
  if (!shown(el)) return 0;
  const kids = [...el.children].filter(shown);
  return kids.length && el.tagName !== "BUTTON" ? kids.reduce((a, k) => a + widthOf(k), 0) : W;
};

async function mountAt(width: number): Promise<HTMLElement> {
  const host = document.createElement("div");
  document.body.appendChild(host);
  createSheetEditor(host, book(), { filename: "s.xlsx" });
  await frame();
  const toolbar = host.querySelector<HTMLElement>(".sheetedit-toolbar")!;
  Object.defineProperty(toolbar, "clientWidth", { configurable: true, get: () => width });
  Object.defineProperty(toolbar, "scrollWidth", { configurable: true, get: () => [...toolbar.children].reduce((a, c) => a + widthOf(c), 0) });
  for (const run of observers) run();
  return host;
}

const visibleButtons = (host: HTMLElement): number =>
  [...host.querySelector<HTMLElement>(".sheetedit-toolbar")!.children].reduce((a, c) => a + widthOf(c), 0) / W;
const folded = (host: HTMLElement): number => host.querySelectorAll(".sheetedit-tb-moremenu .sheetedit-more-item").length;

describe("folding the toolbar on a narrow screen", () => {
  it("uses the room the collapsed style buttons leave for the authoring buttons", async () => {
    const wide = await mountAt(100000);
    const all = visibleButtons(wide);
    expect(folded(wide), "nothing folds when everything fits").toBe(0);
    wide.remove();

    // Narrow enough that the style cluster must collapse, wide enough for several more buttons.
    const host = await mountAt(W * 12);
    expect(visibleButtons(host), "the row is filled, not left half empty").toBeGreaterThanOrEqual(11);
    expect(visibleButtons(host), "and never overflows").toBeLessThanOrEqual(12);
    expect(folded(host), "only what could not fit is in the menu").toBeLessThan(all);
    host.remove();
  });
});
