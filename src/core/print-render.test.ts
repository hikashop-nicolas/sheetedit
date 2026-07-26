import { describe, expect, it } from "vitest";
import { buildPrintDocument, buildPrintJob, pageGeom, pageOrder, paginate, resolveFields, samePaper } from "./print-render";
import type { PrintSetup } from "./print";
import type { Sheet, Workbook } from "./model";

/** A sheet of `rows` x `cols` cells, every line a uniform size. */
function sheetOf(rows: number, cols: number, over: Partial<Sheet> = {}): Sheet {
  const sheet: Sheet = { name: "S", cells: new Map(), maxRow: rows, maxCol: cols, ...over };
  for (let r = 1; r <= rows; r++)
    for (let c = 1; c <= cols; c++) sheet.cells.set(`${r}:${c}`, { row: r, col: c, value: `${r}-${c}`, kind: "s" });
  return sheet;
}
const A4_PORTRAIT: PrintSetup = { paperSize: 9, orientation: "portrait", margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 } };

describe("page geometry", () => {
  it("turns paper size and margins into a body box", () => {
    const g = pageGeom(A4_PORTRAIT);
    expect(Math.round(g.pageW)).toBe(794); // A4 210mm at 96dpi
    expect(Math.round(g.pageH)).toBe(1123);
    expect(Math.round(g.bodyW)).toBe(698); // less the half-inch margins
  });

  it("turns the paper round for landscape", () => {
    const g = pageGeom({ ...A4_PORTRAIT, orientation: "landscape" });
    expect(Math.round(g.pageW)).toBe(1123);
    expect(Math.round(g.pageH)).toBe(794);
  });
});

describe("pagination", () => {
  it("breaks rows when they no longer fit the page", () => {
    // 1000 rows of 24px against ~1027px of body: about 42 rows a page.
    const p = paginate(sheetOf(1000, 3), A4_PORTRAIT, pageGeom(A4_PORTRAIT)).areas[0]!;
    expect(p.rowPages.length).toBeGreaterThan(20);
    expect(p.rowPages[0]!.length).toBeGreaterThan(30);
    // Every row appears exactly once, in order.
    const all = p.rowPages.flat();
    expect(all.length).toBe(1000);
    expect(all[0]).toBe(1);
    expect(all[999]).toBe(1000);
  });

  it("honours a manual break even when the page is not full", () => {
    const p = paginate(sheetOf(20, 2), { ...A4_PORTRAIT, rowBreaks: [5] }, pageGeom(A4_PORTRAIT)).areas[0]!;
    expect(p.rowPages[0]).toEqual([1, 2, 3, 4]);
    expect(p.rowPages[1]![0]).toBe(5);
  });

  it("repeats the title rows on every page instead of printing them once", () => {
    const p = paginate(sheetOf(200, 2), { ...A4_PORTRAIT, titleRows: { from: 1, to: 2 } }, pageGeom(A4_PORTRAIT));
    expect(p.titleRows).toEqual([1, 2]);
    // The repeated rows are not in the body run, or they would print twice on page one.
    expect(p.areas[0]!.rowPages.flat()).not.toContain(1);
    expect(p.areas[0]!.rowPages[0]![0]).toBe(3);
  });

  it("only prints the print area", () => {
    const sheet = sheetOf(50, 20, { printSetup: {} });
    const p = paginate(sheet, { ...A4_PORTRAIT, printArea: [{ r1: 2, c1: 2, r2: 4, c2: 3 }] }, pageGeom(A4_PORTRAIT)).areas[0]!;
    expect(p.rowPages.flat()).toEqual([2, 3, 4]);
    expect(p.colPages.flat()).toEqual([2, 3]);
  });

  it("skips hidden lines, which take no space on the page either", () => {
    const sheet = sheetOf(10, 3, { hiddenRows: new Set([2, 3]) });
    expect(paginate(sheet, A4_PORTRAIT, pageGeom(A4_PORTRAIT)).areas[0]!.rowPages.flat()).toEqual([1, 4, 5, 6, 7, 8, 9, 10]);
  });

  it("shrinks to fit the requested page width, and never enlarges", () => {
    const wide = sheetOf(5, 40); // 40 default columns is far wider than a page
    const geom = pageGeom(A4_PORTRAIT);
    const fitted = paginate(wide, { ...A4_PORTRAIT, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }, geom);
    expect(fitted.scale).toBeLessThan(1);
    expect(fitted.areas[0]!.colPages.length).toBe(1); // the whole width now fits one page
    // A sheet narrower than the page is left alone rather than blown up.
    const narrow = paginate(sheetOf(5, 2), { ...A4_PORTRAIT, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }, geom);
    expect(narrow.scale).toBe(1);
  });

  it("a percentage scale changes how much fits on a page", () => {
    const geom = pageGeom(A4_PORTRAIT);
    const full = paginate(sheetOf(300, 2), A4_PORTRAIT, geom).areas[0]!;
    const half = paginate(sheetOf(300, 2), { ...A4_PORTRAIT, scale: 50 }, geom).areas[0]!;
    expect(half.rowPages[0]!.length).toBeGreaterThan(full.rowPages[0]!.length);
  });

  it("orders the pages the way the setup asks", () => {
    const p = { areas: [{ colPages: [[1], [2]], rowPages: [[1], [2]] }], scale: 1, titleRows: [], titleCols: [] };
    expect(pageOrder(p, "downThenOver").map((x) => `${x.cols[0]}/${x.rows[0]}`)).toEqual(["1/1", "1/2", "2/1", "2/2"]);
    expect(pageOrder(p, "overThenDown").map((x) => `${x.cols[0]}/${x.rows[0]}`)).toEqual(["1/1", "2/1", "1/2", "2/2"]);
  });
});

describe("header and footer fields", () => {
  const ctx = { page: 2, pages: 7, sheet: "Sales", file: "book.xlsx" };
  it("resolves the page, count, sheet and file codes", () => {
    expect(resolveFields("Page &P of &N", ctx)).toBe("Page 2 of 7");
    expect(resolveFields("&A - &F", ctx)).toBe("Sales - book.xlsx");
  });
  it("keeps a doubled ampersand as a literal", () => {
    expect(resolveFields("R&&D", ctx)).toBe("R&D");
  });
  it("drops the formatting codes, which have no plain-text meaning", () => {
    expect(resolveFields('&"Arial,Bold"&12Total', ctx)).toBe("Total");
  });
});

describe("print document", () => {
  const wbOf = (sheet: Sheet): Workbook => ({ kind: "xlsx", sheets: [sheet], files: {} });

  it("builds one page box per page", () => {
    const sheet = sheetOf(200, 2, { printSetup: A4_PORTRAIT });
    const root = buildPrintDocument(wbOf(sheet), 0, "book.xlsx")!;
    const pages = root.querySelectorAll(".sheetedit-print-page");
    expect(pages.length).toBeGreaterThan(1);
    expect(parseFloat((pages[0] as HTMLElement).style.width)).toBeCloseTo(793.7, 1); // A4 210mm
  });

  it("repeats the title row on the second page", () => {
    const sheet = sheetOf(200, 2, { printSetup: { ...A4_PORTRAIT, titleRows: { from: 1, to: 1 } } });
    const root = buildPrintDocument(wbOf(sheet), 0, "book.xlsx")!;
    const second = root.querySelectorAll(".sheetedit-print-page")[1]!;
    expect(second.textContent).toContain("1-1"); // the title row's own cell
  });

  it("numbers the pages through the footer, resolving &P and &N", () => {
    const sheet = sheetOf(200, 2, { printSetup: { ...A4_PORTRAIT, footer: { center: "Page &P of &N" } } });
    const root = buildPrintDocument(wbOf(sheet), 0, "book.xlsx")!;
    const feet = [...root.querySelectorAll(".is-footer")].map((f) => f.textContent);
    expect(feet[0]).toContain("Page 1 of ");
    expect(feet[1]).toContain("Page 2 of ");
    expect(feet[0]).toContain(`of ${feet.length}`);
  });

  it("starts at the first page number the setup gives", () => {
    const sheet = sheetOf(50, 2, { printSetup: { ...A4_PORTRAIT, firstPageNumber: 5, footer: { center: "&P" } } });
    const root = buildPrintDocument(wbOf(sheet), 0, "book.xlsx")!;
    expect(root.querySelector(".is-footer")!.textContent).toContain("5");
  });

  it("adds row and column headings only when asked", () => {
    const plain = buildPrintDocument(wbOf(sheetOf(3, 2, { printSetup: A4_PORTRAIT })), 0, "b")!;
    expect(plain.querySelector(".sheetedit-print-head")).toBeNull();
    const withHeads = buildPrintDocument(wbOf(sheetOf(3, 2, { printSetup: { ...A4_PORTRAIT, headings: true } })), 0, "b")!;
    expect(withHeads.querySelector(".sheetedit-print-head")!.textContent).toBe("");
    expect(withHeads.textContent).toContain("A");
  });

  it("carries the cell's own formatting onto the paper", () => {
    const sheet = sheetOf(2, 2, { printSetup: A4_PORTRAIT });
    sheet.cells.get("1:1")!.cellStyle = { bold: true, bg: "#ff0000" };
    const root = buildPrintDocument(wbOf(sheet), 0, "b")!;
    const td = root.querySelector("td") as HTMLElement;
    expect(td.style.fontWeight).toBe("700");
    expect(td.style.background).toBe("rgb(255, 0, 0)");
  });

  it("spans a merged cell instead of printing it as separate ones", () => {
    const sheet = sheetOf(3, 3, { printSetup: A4_PORTRAIT, merges: [{ r1: 1, c1: 1, r2: 1, c2: 3 }] });
    const root = buildPrintDocument(wbOf(sheet), 0, "b")!;
    const first = root.querySelector("tr td") as HTMLTableCellElement;
    expect(first.colSpan).toBe(3);
    expect(root.querySelectorAll("tr")[0]!.querySelectorAll("td").length).toBe(1);
  });

  it("sets a @page rule matching the paper, so the browser does not impose its own", () => {
    const root = buildPrintDocument(wbOf(sheetOf(2, 2, { printSetup: { ...A4_PORTRAIT, orientation: "landscape" } })), 0, "b")!;
    const css = root.querySelector("style")!.textContent!;
    expect(css).toContain("margin: 0");
    expect(css).toMatch(/size: 11\.69in 8\.27in/);
  });

  it("returns nothing for a sheet with nothing in it", () => {
    expect(buildPrintDocument(wbOf({ name: "S", cells: new Map(), maxRow: 0, maxCol: 0 }), 0, "b")).toBeNull();
  });
});

describe("multi-range print areas", () => {
  const wbOf = (...sheets: Sheet[]): Workbook => ({ kind: "xlsx", sheets, files: {} });

  it("paginates each range separately", () => {
    const setup = { ...A4_PORTRAIT, printArea: [{ r1: 1, c1: 1, r2: 2, c2: 2 }, { r1: 5, c1: 1, r2: 6, c2: 2 }] };
    const p = paginate(sheetOf(10, 4), setup, pageGeom(A4_PORTRAIT));
    expect(p.areas.length).toBe(2);
    expect(p.areas[0]!.rowPages.flat()).toEqual([1, 2]);
    expect(p.areas[1]!.rowPages.flat()).toEqual([5, 6]);
  });

  it("starts a new page for each range rather than running them together", () => {
    // Both ranges are tiny and would share a page if they were one continuous run.
    const setup = { ...A4_PORTRAIT, printArea: [{ r1: 1, c1: 1, r2: 2, c2: 2 }, { r1: 5, c1: 1, r2: 6, c2: 2 }] };
    const sheet = sheetOf(10, 4, { printSetup: setup });
    const root = buildPrintJob(wbOf(sheet), 0, "b")!.root;
    const pages = root.querySelectorAll(".sheetedit-print-page");
    expect(pages.length).toBe(2);
    expect(pages[0]!.textContent).toContain("1-1");
    expect(pages[0]!.textContent).not.toContain("5-1");
    expect(pages[1]!.textContent).toContain("5-1");
  });

  it("sizes fit-to from the widest range, so the ranges stay comparable", () => {
    const setup = { ...A4_PORTRAIT, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      printArea: [{ r1: 1, c1: 1, r2: 2, c2: 2 }, { r1: 5, c1: 1, r2: 6, c2: 40 }] };
    const p = paginate(sheetOf(10, 40), setup, pageGeom(A4_PORTRAIT));
    expect(p.scale).toBeLessThan(1); // the wide range decides it
    expect(p.areas[1]!.colPages.length).toBe(1);
  });
});

describe("print scope", () => {
  const wbOf = (...sheets: Sheet[]): Workbook => ({ kind: "xlsx", sheets, files: {} });
  const twoSheets = () => [
    sheetOf(3, 2, { printSetup: A4_PORTRAIT }),
    { ...sheetOf(3, 2, { printSetup: A4_PORTRAIT }), name: "Second" },
  ];

  it("prints only the active sheet by default", () => {
    const [a, b] = twoSheets();
    const root = buildPrintJob(wbOf(a!, b!), 0, "f")!.root;
    const sheets = new Set([...root.querySelectorAll(".sheetedit-print-page")].map((p) => (p as HTMLElement).dataset.sheet));
    expect([...sheets]).toEqual(["S"]);
  });

  it("prints every sheet when the job says so", () => {
    const [a, b] = twoSheets();
    const root = buildPrintJob(wbOf(a!, b!), 0, "f", { scope: "all" })!.root;
    const sheets = [...root.querySelectorAll(".sheetedit-print-page")].map((p) => (p as HTMLElement).dataset.sheet);
    expect(sheets).toEqual(["S", "Second"]);
  });

  it("numbers the pages through the whole job, not per sheet", () => {
    const setup = { ...A4_PORTRAIT, footer: { center: "&P/&N" } };
    const a = sheetOf(3, 2, { printSetup: setup });
    const b = { ...sheetOf(3, 2, { printSetup: setup }), name: "Second" };
    const root = buildPrintJob(wbOf(a, b), 0, "f", { scope: "all" })!.root;
    expect([...root.querySelectorAll(".is-footer")].map((f) => f.textContent)).toEqual(["1/2", "2/2"]);
  });

  it("skips a sheet with nothing to print instead of emitting a blank page", () => {
    const a = sheetOf(3, 2, { printSetup: A4_PORTRAIT });
    const empty: Sheet = { name: "Empty", cells: new Map(), maxRow: 0, maxCol: 0 };
    const root = buildPrintJob(wbOf(a, empty), 0, "f", { scope: "all" })!.root;
    expect(root.querySelectorAll(".sheetedit-print-page").length).toBe(1);
  });

  it("prints just the selection, without touching the sheet's own print area", () => {
    const sheet = sheetOf(10, 4, { printSetup: { ...A4_PORTRAIT, printArea: [{ r1: 1, c1: 1, r2: 10, c2: 4 }] } });
    const root = buildPrintJob(wbOf(sheet), 0, "f", { scope: "sheet", selection: { r1: 2, c1: 1, r2: 3, c2: 2 } })!.root;
    const text = root.textContent!;
    expect(text).toContain("2-1");
    expect(text).not.toContain("9-1");
    // The job is transient: the sheet keeps the area it had.
    expect(sheet.printSetup!.printArea).toEqual([{ r1: 1, c1: 1, r2: 10, c2: 4 }]);
  });

  it("flags a job whose sheets disagree on paper, which one @page cannot express", () => {
    const a = sheetOf(3, 2, { printSetup: { ...A4_PORTRAIT, paperSize: 9 } });
    const b = { ...sheetOf(3, 2, { printSetup: { ...A4_PORTRAIT, paperSize: 8 } }), name: "A3" };
    expect(buildPrintJob(wbOf(a, b), 0, "f", { scope: "all" })!.mixedPaper).toBe(true);
    expect(buildPrintJob(wbOf(a), 0, "f", { scope: "all" })!.mixedPaper).toBe(false);
    expect(samePaper([{ geom: pageGeom(A4_PORTRAIT) }, { geom: pageGeom(A4_PORTRAIT) }])).toBe(true);
  });

  it("returns nothing when no sheet in the job has anything", () => {
    const empty: Sheet = { name: "Empty", cells: new Map(), maxRow: 0, maxCol: 0 };
    expect(buildPrintJob(wbOf(empty), 0, "f", { scope: "all" })).toBeNull();
    expect(buildPrintDocument(wbOf(empty), 0, "f")).toBeNull();
  });
});
