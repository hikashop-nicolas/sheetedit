import { cellDisplay, colToLetters, getCell, key, type CellStyle, type Sheet, type Workbook } from "./model";
import { DEFAULT_MARGINS, DEFAULT_PAPER, MM_PER_INCH, PAPER_SIZES, effectivePrintArea, type HeaderFooter, type PrintSetup } from "./print";

// ---------------------------------------------------------------------------
// Printing
// ---------------------------------------------------------------------------
// The browser can only print what is in the DOM, so printing means laying the print area out as
// pages ourselves and handing those to window.print(). Doing it properly is what makes the page
// setup mean something: paper size, margins, scaling, manual breaks, repeated title rows and page
// order all decide where the breaks land, and none of that survives simply printing the grid.
//
// Each page is a fixed-size box with the paper's dimensions and @page margin 0, so the margins,
// header and footer are ours to place rather than the browser's. What we cannot control is the
// browser's own header/footer (the URL and date strip); that is a print-dialog setting.

const CSS_DPI = 96;
const inchToPx = (n: number): number => n * CSS_DPI;
const mmToPx = (n: number): number => (n / MM_PER_INCH) * CSS_DPI;

/** A page's geometry, in CSS pixels. */
export interface PageGeom {
  pageW: number;
  pageH: number;
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
  /** The body box a page's cells are laid out in, before scaling. */
  bodyW: number;
  bodyH: number;
}

export function pageGeom(setup: PrintSetup): PageGeom {
  const paper = PAPER_SIZES[setup.paperSize ?? DEFAULT_PAPER] ?? PAPER_SIZES[DEFAULT_PAPER]!;
  const landscape = setup.orientation === "landscape";
  const pageW = mmToPx(landscape ? paper.h : paper.w);
  const pageH = mmToPx(landscape ? paper.w : paper.h);
  const m = setup.margins ?? DEFAULT_MARGINS;
  const left = inchToPx(m.left), right = inchToPx(m.right), top = inchToPx(m.top), bottom = inchToPx(m.bottom);
  return { pageW, pageH, left, right, top, bottom, header: inchToPx(m.header), footer: inchToPx(m.footer), bodyW: pageW - left - right, bodyH: pageH - top - bottom };
}

const colWidth = (sheet: Sheet, c: number): number =>
  sheet.hiddenCols?.has(c) ? 0 : (sheet.colWidths?.get(c) ?? sheet.defaultColWidth ?? 96);
const rowHeight = (sheet: Sheet, r: number): number =>
  sheet.hiddenRows?.has(r) || sheet.filterHidden?.has(r) ? 0 : (sheet.rowHeights?.get(r) ?? sheet.defaultRowHeight ?? 24);

/** Split a run of lines into pages: at a manual break, or when the next line would overflow. */
function splitLines(lines: number[], sizeOf: (n: number) => number, capacity: number, breaks: Set<number>, repeatSize: number): number[][] {
  const pages: number[][] = [];
  let cur: number[] = [];
  let used = repeatSize;
  for (const line of lines) {
    const size = sizeOf(line);
    if (size === 0) continue; // hidden lines are not printed and take no space
    const forced = breaks.has(line) && cur.length > 0;
    if (forced || (cur.length > 0 && used + size > capacity)) {
      pages.push(cur);
      cur = [];
      used = repeatSize;
    }
    cur.push(line);
    used += size;
  }
  if (cur.length) pages.push(cur);
  return pages.length ? pages : [[]];
}

/** One print area's page grid: which columns and which rows each of its pages carries. */
export interface AreaPages {
  colPages: number[][];
  rowPages: number[][];
}

export interface Pagination {
  /** One entry per range of the print area, in the order the file lists them. */
  areas: AreaPages[];
  /** The scale actually applied, as a fraction (fit-to may override the stated percentage). */
  scale: number;
  /** Rows / columns repeated on every page. */
  titleRows: number[];
  titleCols: number[];
}

/** Work out where the pages break for one sheet's print area. */
export function paginate(sheet: Sheet, setup: PrintSetup, geom: PageGeom): Pagination {
  const areas = effectivePrintArea(sheet, setup);
  const titleRows: number[] = [];
  const titleCols: number[] = [];
  if (setup.titleRows) for (let r = setup.titleRows.from; r <= setup.titleRows.to; r++) if (rowHeight(sheet, r) > 0) titleRows.push(r);
  if (setup.titleCols) for (let c = setup.titleCols.from; c <= setup.titleCols.to; c++) if (colWidth(sheet, c) > 0) titleCols.push(c);
  // No area at all (an empty sheet) means there is genuinely nothing to lay out.
  if (!areas.length) return { areas: [], scale: 1, titleRows, titleCols };

  const repeatW = titleCols.reduce((a, c) => a + colWidth(sheet, c), 0);
  const repeatH = titleRows.reduce((a, r) => a + rowHeight(sheet, r), 0);
  // The lines each range contributes, with the repeated title lines taken out of the body run.
  const runs = areas.map((area) => {
    const cols: number[] = [];
    for (let c = area.c1; c <= area.c2; c++) if (!titleCols.includes(c)) cols.push(c);
    const rows: number[] = [];
    for (let r = area.r1; r <= area.r2; r++) if (!titleRows.includes(r)) rows.push(r);
    return { cols, rows };
  });

  // Fit-to overrides the stated scale, and applies to the sheet as a whole: one scale for every
  // range, sized by the largest, so the ranges stay comparable on the paper.
  let scale = (setup.fitToPage ? 100 : setup.scale ?? 100) / 100;
  if (setup.fitToPage) {
    const totalW = repeatW + Math.max(...runs.map((r) => r.cols.reduce((a, c) => a + colWidth(sheet, c), 0)));
    const totalH = repeatH + Math.max(...runs.map((r) => r.rows.reduce((a, x) => a + rowHeight(sheet, x), 0)));
    const fits: number[] = [];
    if (setup.fitToWidth) fits.push((geom.bodyW * setup.fitToWidth) / Math.max(1, totalW));
    if (setup.fitToHeight) fits.push((geom.bodyH * setup.fitToHeight) / Math.max(1, totalH));
    // Never scale up: Excel's fit-to only shrinks.
    if (fits.length) scale = Math.min(1, ...fits);
  }
  const capW = geom.bodyW / scale;
  const capH = geom.bodyH / scale;

  return {
    areas: runs.map((run) => ({
      colPages: splitLines(run.cols, (c) => colWidth(sheet, c), capW, new Set(setup.colBreaks ?? []), repeatW),
      rowPages: splitLines(run.rows, (r) => rowHeight(sheet, r), capH, new Set(setup.rowBreaks ?? []), repeatH),
    })),
    scale,
    titleRows,
    titleCols,
  };
}

/**
 * The pages, in the order the setup asks for. Each range of a multi-range print area starts its own
 * page, which is how Excel prints them: the ranges are separate blocks, not one continuous grid.
 */
export function pageOrder(p: Pagination, order: PrintSetup["pageOrder"]): { cols: number[]; rows: number[] }[] {
  const out: { cols: number[]; rows: number[] }[] = [];
  for (const area of p.areas) {
    if (order === "overThenDown") {
      for (const rows of area.rowPages) for (const cols of area.colPages) out.push({ cols, rows });
    } else {
      for (const cols of area.colPages) for (const rows of area.rowPages) out.push({ cols, rows });
    }
  }
  // A range that contributes no lines at all contributes no page.
  return out.filter((page) => page.cols.length && page.rows.length);
}

/** Substitute the field codes a header or footer region can carry. */
export function resolveFields(text: string, ctx: { page: number; pages: number; sheet: string; file: string }): string {
  const now = new Date();
  return text.replace(/&(&|[A-Za-z])/g, (_m, c: string) => {
    switch (c) {
      case "&": return "&";
      case "P": return String(ctx.page);
      case "N": return String(ctx.pages);
      case "D": return now.toLocaleDateString();
      case "T": return now.toLocaleTimeString();
      case "A": return ctx.sheet;
      case "F": case "Z": return ctx.file;
      // Formatting codes (&B bold, &12 size, &"font") have no plain-text meaning here.
      default: return "";
    }
  }).replace(/&"[^"]*"/g, "").replace(/&\d+/g, "");
}

const applyStyle = (el: HTMLElement, st: CellStyle | undefined): void => {
  if (!st) return;
  // The workbook's own formatting is data, so it is set per element rather than by a class.
  if (st.bg) el.style.background = st.bg;
  if (st.color) el.style.color = st.color;
  if (st.bold) el.style.fontWeight = "700";
  if (st.italic) el.style.fontStyle = "italic";
  if (st.underline || st.strike) el.style.textDecoration = `${st.underline ? "underline" : ""} ${st.strike ? "line-through" : ""}`.trim();
  if (st.fontSize) el.style.fontSize = `${st.fontSize}pt`;
  if (st.fontFamily) el.style.fontFamily = st.fontFamily;
  if (st.align) el.style.textAlign = st.align;
  if (st.valign) el.style.verticalAlign = st.valign === "middle" ? "middle" : st.valign;
  if (st.wrap) el.style.whiteSpace = "pre-wrap";
  const b = st.borders;
  if (b) {
    if (b.top) el.style.borderTop = `1px solid ${b.top}`;
    if (b.right) el.style.borderRight = `1px solid ${b.right}`;
    if (b.bottom) el.style.borderBottom = `1px solid ${b.bottom}`;
    if (b.left) el.style.borderLeft = `1px solid ${b.left}`;
  }
};

/** Build the table for one page: the repeated title lines plus this page's own rows and columns. */
function buildPageTable(sheet: Sheet, p: Pagination, cols: number[], rows: number[], setup: PrintSetup): HTMLElement {
  const allCols = [...p.titleCols, ...cols];
  const allRows = [...p.titleRows, ...rows];
  const table = document.createElement("table");
  table.className = "sheetedit-print-table";
  if (setup.gridLines) table.classList.add("has-grid");

  // Merges: the top-left cell spans, the covered ones are skipped.
  const covered = new Set<string>();
  const spanAt = new Map<string, { rs: number; cs: number }>();
  for (const m of sheet.merges ?? []) {
    spanAt.set(key(m.r1, m.c1), { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(key(r, c));
  }

  const colgroup = document.createElement("colgroup");
  if (setup.headings) colgroup.appendChild(document.createElement("col")).style.width = "36px";
  for (const c of allCols) colgroup.appendChild(document.createElement("col")).style.width = `${colWidth(sheet, c)}px`;
  table.appendChild(colgroup);

  if (setup.headings) {
    const tr = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "sheetedit-print-head";
    tr.appendChild(corner);
    for (const c of allCols) {
      const th = document.createElement("th");
      th.className = "sheetedit-print-head";
      th.textContent = colToLetters(c);
      tr.appendChild(th);
    }
    table.appendChild(tr);
  }

  for (const r of allRows) {
    const tr = document.createElement("tr");
    tr.style.height = `${rowHeight(sheet, r)}px`;
    if (setup.headings) {
      const th = document.createElement("th");
      th.className = "sheetedit-print-head";
      th.textContent = String(r);
      tr.appendChild(th);
    }
    for (const c of allCols) {
      if (covered.has(key(r, c))) continue;
      const td = document.createElement("td");
      const span = spanAt.get(key(r, c));
      if (span) {
        if (span.rs > 1) td.rowSpan = span.rs;
        if (span.cs > 1) td.colSpan = span.cs;
      }
      const cell = getCell(sheet, r, c);
      td.textContent = cellDisplay(cell);
      if (cell?.kind === "n") td.classList.add("num");
      applyStyle(td, cell?.cellStyle);
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
  return table;
}

/** One header/footer band, with its three regions. */
function buildBand(hf: HeaderFooter | undefined, cls: string, ctx: { page: number; pages: number; sheet: string; file: string }): HTMLElement | null {
  if (!hf || (!hf.left && !hf.center && !hf.right)) return null;
  const band = document.createElement("div");
  band.className = cls;
  for (const region of ["left", "center", "right"] as const) {
    const span = document.createElement("span");
    span.className = `sheetedit-print-${region}`;
    span.textContent = hf[region] ? resolveFields(hf[region]!, ctx) : "";
    band.appendChild(span);
  }
  return band;
}

/** What a print job covers. */
export interface PrintJob {
  /** Which sheets: the active one, or every sheet in the workbook. */
  scope: "sheet" | "all";
  /** Restrict to this range instead of the sheet's print area (used for "print the selection"). */
  selection?: { r1: number; c1: number; r2: number; c2: number };
}

/** One sheet's contribution to a job, already laid out. */
interface SheetPlan {
  sheet: Sheet;
  setup: PrintSetup;
  geom: PageGeom;
  pagination: Pagination;
  pages: { cols: number[]; rows: number[] }[];
}

/** Lay out each sheet in the job, dropping the ones with nothing to print. */
function planJob(wb: Workbook, sheetIndex: number, job: PrintJob): SheetPlan[] {
  const indexes = job.scope === "all" ? wb.sheets.map((_s, i) => i) : [sheetIndex];
  const plans: SheetPlan[] = [];
  for (const i of indexes) {
    const sheet = wb.sheets[i];
    if (!sheet) continue;
    // A selection overrides the sheet's own print area for this job only; it is never committed.
    const setup: PrintSetup = job.selection ? { ...(sheet.printSetup ?? {}), printArea: [job.selection] } : sheet.printSetup ?? {};
    const geom = pageGeom(setup);
    const pagination = paginate(sheet, setup, geom);
    const pages = pageOrder(pagination, setup.pageOrder);
    if (pages.length) plans.push({ sheet, setup, geom, pagination, pages });
  }
  return plans;
}

/** Whether every sheet in the job agrees on the paper, which one @page rule can express. */
export function samePaper(plans: { geom: PageGeom }[]): boolean {
  return plans.every((p) => Math.abs(p.geom.pageW - plans[0]!.geom.pageW) < 1 && Math.abs(p.geom.pageH - plans[0]!.geom.pageH) < 1);
}

export interface PrintResult {
  root: HTMLElement;
  /** Set when the job spans sheets whose paper differs, which one @page rule cannot express. */
  mixedPaper: boolean;
}

/**
 * Build the printable document: one absolutely-sized box per page, each carrying its own header,
 * body and footer. Returns null when the job has nothing to print.
 */
export function buildPrintJob(wb: Workbook, sheetIndex: number, fileName: string, job: PrintJob = { scope: "sheet" }): PrintResult | null {
  const plans = planJob(wb, sheetIndex, job);
  if (!plans.length) return null;

  const root = document.createElement("div");
  root.className = "sheetedit-print";
  // @page margin 0 plus a page box of exactly the paper size: the margins, header and footer are
  // then ours to place, which is the only way to honour the file's own values. A browser applies
  // one @page to the whole job, so a mixed-paper job takes the first sheet's and says so.
  const first = plans[0]!.geom;
  const style = document.createElement("style");
  style.textContent = `@page { size: ${Math.round((first.pageW / CSS_DPI) * 100) / 100}in ${Math.round((first.pageH / CSS_DPI) * 100) / 100}in; margin: 0; }`;
  root.appendChild(style);

  // Page numbers run through the whole job, so &N is the job's total rather than one sheet's.
  const total = plans.reduce((a, p) => a + p.pages.length, 0);
  let n = 0;
  for (const plan of plans) {
    const { sheet, setup, geom, pagination } = plan;
    plan.pages.forEach((page) => {
      const ctx = { page: (setup.firstPageNumber ?? 1) + n, pages: total, sheet: sheet.name, file: fileName };
      n++;
      const box = document.createElement("div");
      box.className = "sheetedit-print-page";
      box.dataset.sheet = sheet.name;
      Object.assign(box.style, { width: `${geom.pageW}px`, height: `${geom.pageH}px` });

      const header = buildBand(setup.header, "sheetedit-print-band is-header", ctx);
      if (header) {
        Object.assign(header.style, { top: `${geom.header}px`, left: `${geom.left}px`, right: `${geom.right}px` });
        box.appendChild(header);
      }
      const body = document.createElement("div");
      body.className = "sheetedit-print-body";
      Object.assign(body.style, {
        left: `${geom.left}px`, top: `${geom.top}px`,
        width: `${geom.bodyW}px`, height: `${geom.bodyH}px`,
        justifyContent: setup.horizontalCentered ? "center" : "flex-start",
        alignItems: setup.verticalCentered ? "center" : "flex-start",
      });
      const scaler = document.createElement("div");
      // The scale is a zoom of the whole page body, exactly as the setting means it.
      Object.assign(scaler.style, { transform: `scale(${pagination.scale})`, transformOrigin: setup.horizontalCentered ? "top center" : "top left" });
      scaler.appendChild(buildPageTable(sheet, pagination, page.cols, page.rows, setup));
      body.appendChild(scaler);
      box.appendChild(body);

      const footer = buildBand(setup.footer, "sheetedit-print-band is-footer", ctx);
      if (footer) {
        Object.assign(footer.style, { bottom: `${geom.footer}px`, left: `${geom.left}px`, right: `${geom.right}px` });
        box.appendChild(footer);
      }
      root.appendChild(box);
    });
  }
  return { root, mixedPaper: !samePaper(plans) };
}

/** The single-sheet job, kept as the simple entry point. */
export function buildPrintDocument(wb: Workbook, sheetIndex: number, fileName: string): HTMLElement | null {
  return buildPrintJob(wb, sheetIndex, fileName)?.root ?? null;
}
