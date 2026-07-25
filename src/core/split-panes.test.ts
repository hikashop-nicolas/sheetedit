import { describe, expect, it, beforeEach } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** A workbook with enough rows that a split has something to scroll. */
function book(pane = ""): Uint8Array {
  const rows = Array.from({ length: 40 }, (_, i) =>
    `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>r${i + 1}</t></is></c></row>`).join("");
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}">${pane}<sheetData>${rows}</sheetData></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${MAIN}"/>`),
  });
}

const SPLIT = `<sheetViews><sheetView workbookViewId="0"><pane ySplit="600" topLeftCell="A4" state="split"/></sheetView></sheetViews>`;
const COLSPLIT = `<sheetViews><sheetView workbookViewId="0"><pane xSplit="960" topLeftCell="B1" state="split"/></sheetView></sheetViews>`;
const BOTH = `<sheetViews><sheetView workbookViewId="0"><pane xSplit="960" ySplit="600" topLeftCell="B4" state="split"/></sheetView></sheetViews>`;
const FROZEN = `<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" state="frozen"/></sheetView></sheetViews>`;

// jsdom lacks ResizeObserver (the toolbar overflow logic uses it).
if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
  (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
}

let host: HTMLElement;
beforeEach(() => {
  document.body.innerHTML = "";
  host = document.createElement("div");
  document.body.appendChild(host);
});

/** The viewports actually on screen; the unused quadrants stay in the DOM hidden. */
const panes = (): HTMLElement[] =>
  ([...host.querySelectorAll(".sheetedit-grid")] as HTMLElement[]).filter((p) => p.style.display !== "none");
const splitPane = (): HTMLElement | null => host.querySelector(".sheetedit-grid-split");

describe("split panes", () => {
  it("gives a row split a second scrolling viewport", () => {
    createSheetEditor(host, book(SPLIT));
    expect(panes().length).toBe(2);
    expect(splitPane()!.style.display).not.toBe("none");
  });

  it("keeps one viewport for a frozen boundary", () => {
    createSheetEditor(host, book(FROZEN));
    expect(panes().length).toBe(1);
    expect(splitPane()!.style.display).toBe("none");
  });

  it("keeps one viewport when there is no boundary at all", () => {
    createSheetEditor(host, book());
    expect(panes().length).toBe(1);
    expect(splitPane()!.style.display).toBe("none");
  });

  it("draws the column header once, in the top pane only", () => {
    createSheetEditor(host, book(SPLIT));
    const [top, bottom] = panes();
    expect(top!.querySelectorAll("th.colhead").length).toBeGreaterThan(0);
    expect(bottom!.querySelectorAll("th.colhead").length).toBe(0);
  });

  it("renders cells in both viewports, so a row can be on screen twice", () => {
    createSheetEditor(host, book(SPLIT));
    const [top, bottom] = panes();
    expect(top!.querySelectorAll("td[data-rc]").length).toBeGreaterThan(0);
    expect(bottom!.querySelectorAll("td[data-rc]").length).toBeGreaterThan(0);
  });

  it("makes no row sticky inside a split pane, unlike a freeze", () => {
    createSheetEditor(host, book(SPLIT));
    expect(host.querySelectorAll("tr th.rownum.frz").length).toBe(0);
    document.body.innerHTML = "";
    const host2 = document.createElement("div");
    document.body.appendChild(host2);
    createSheetEditor(host2, book(FROZEN));
    expect(host2.querySelectorAll("th.rownum.frz").length).toBeGreaterThan(0);
  });

  it("gives a column split a second viewport beside the first", () => {
    createSheetEditor(host, book(COLSPLIT));
    expect(panes().length).toBe(2);
    expect(host.querySelector(".sheetedit-grid-right")!.getAttribute("style")).not.toContain("display: none");
  });

  it("gives a split on both axes all four viewports", () => {
    createSheetEditor(host, book(BOTH));
    expect(panes().length).toBe(4);
  });

  it("draws the row numbers in the left band only", () => {
    createSheetEditor(host, book(COLSPLIT));
    const [left, right] = panes();
    expect(left!.querySelectorAll("th.rownum").length).toBeGreaterThan(0);
    expect(right!.querySelectorAll("th.rownum").length).toBe(0);
    // The right band still draws the column header, being in the top row band.
    expect(right!.querySelectorAll("th.colhead").length).toBeGreaterThan(0);
  });

  it("makes no column sticky inside a split pane, unlike a freeze", () => {
    createSheetEditor(host, book(COLSPLIT));
    expect(host.querySelectorAll("th.colhead.frz").length).toBe(0);
  });

  it("tears the second viewport down with the editor", () => {
    const ed = createSheetEditor(host, book(SPLIT));
    expect(panes().length).toBe(2);
    ed.destroy();
    expect(host.querySelectorAll(".sheetedit-grid").length).toBe(0);
  });
});
