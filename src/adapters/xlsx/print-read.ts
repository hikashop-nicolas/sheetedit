import { firstByLocal, parseA1Ref, type Sheet, type Workbook } from "../../core/model";
import { parseHeaderFooter, type PrintSetup } from "../../core/print";

// xlsx print settings live in five worksheet elements (printOptions, pageMargins, pageSetup,
// headerFooter, row/colBreaks), one flag on sheetPr, and two sheet-scoped defined names in
// workbook.xml (_xlnm.Print_Area and _xlnm.Print_Titles).

const bool = (el: Element | undefined, name: string): boolean | undefined => {
  const v = el?.getAttribute(name);
  if (v == null) return undefined;
  return v === "1" || v === "true";
};
const num = (el: Element | undefined, name: string): number | undefined => {
  const v = el?.getAttribute(name);
  if (v == null || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
};

/** "Sheet1!$A$1:$C$5" (or "$A$1:$C$5") -> a 1-based inclusive range. */
export function parseAreaRef(ref: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const body = (ref.includes("!") ? ref.slice(ref.lastIndexOf("!") + 1) : ref).replace(/\$/g, "");
  const [a, b] = body.split(":");
  const p1 = parseA1Ref(a ?? "");
  const p2 = b ? parseA1Ref(b) : p1;
  if (!p1 || !p2) return null;
  return { r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) };
}

/**
 * A Print_Titles value: "Sheet!$1:$2" for repeated rows, "Sheet!$A:$B" for columns, or both joined
 * by a comma. Each part names whole lines, so it is parsed on its own rather than as a cell range.
 */
export function parseTitles(value: string): { rows?: { from: number; to: number }; cols?: { from: number; to: number } } {
  const out: { rows?: { from: number; to: number }; cols?: { from: number; to: number } } = {};
  for (const part of value.split(",")) {
    const body = (part.includes("!") ? part.slice(part.lastIndexOf("!") + 1) : part).replace(/\$/g, "").trim();
    const m = /^([A-Za-z]+|\d+):([A-Za-z]+|\d+)$/.exec(body);
    if (!m) continue;
    const [, a, b] = m;
    if (/^\d+$/.test(a!)) out.rows = { from: Number(a), to: Number(b) };
    else {
      const c1 = parseA1Ref(`${a}1`), c2 = parseA1Ref(`${b}1`);
      if (c1 && c2) out.cols = { from: c1.col, to: c2.col };
    }
  }
  return out;
}

/** The manual break lines from a <rowBreaks> / <colBreaks> element. */
function readBreaks(el: Element | undefined): number[] | undefined {
  if (!el) return undefined;
  const out: number[] = [];
  for (const brk of Array.from(el.children)) {
    if (brk.localName !== "brk") continue;
    // @man distinguishes a user break from one the app computed; only manual ones are ours to keep.
    const manual = brk.getAttribute("man");
    if (manual != null && manual !== "1" && manual !== "true") continue;
    const id = Number(brk.getAttribute("id") || "0");
    // The break starts a new page AT that line; xlsx counts it 0-based from the sheet origin.
    if (id > 0) out.push(id + 1);
  }
  return out.length ? out.sort((a, b) => a - b) : undefined;
}

/** Read one worksheet's print settings (everything except the workbook-level defined names). */
export function readXlsxPrintSetup(sheet: Sheet, doc: Document): void {
  const ws = doc.documentElement;
  const el = (name: string): Element | undefined => firstByLocal(ws, name);
  const opts = el("printOptions"), margins = el("pageMargins"), setup = el("pageSetup"), hf = el("headerFooter");
  const sheetPr = el("sheetPr");
  const setUpPr = sheetPr ? firstByLocal(sheetPr, "pageSetUpPr") : undefined;
  const p: PrintSetup = {};

  const orient = setup?.getAttribute("orientation");
  if (orient === "landscape" || orient === "portrait") p.orientation = orient;
  p.paperSize = num(setup, "paperSize");
  p.scale = num(setup, "scale");
  p.fitToWidth = num(setup, "fitToWidth");
  p.fitToHeight = num(setup, "fitToHeight");
  p.fitToPage = bool(setUpPr, "fitToPage");
  const order = setup?.getAttribute("pageOrder");
  if (order === "overThenDown" || order === "downThenOver") p.pageOrder = order;
  // A first page number only means anything when the sheet opts into using it.
  if (bool(setup, "useFirstPageNumber")) {
    const n = num(setup, "firstPageNumber");
    if (n != null && n !== 1) p.firstPageNumber = n;
  }

  if (margins) {
    const m = (n: string): number => num(margins, n) ?? 0;
    p.margins = { left: m("left"), right: m("right"), top: m("top"), bottom: m("bottom"), header: m("header"), footer: m("footer") };
  }
  p.gridLines = bool(opts, "gridLines");
  p.headings = bool(opts, "headings");
  p.horizontalCentered = bool(opts, "horizontalCentered");
  p.verticalCentered = bool(opts, "verticalCentered");

  if (hf) {
    p.header = parseHeaderFooter(firstByLocal(hf, "oddHeader")?.textContent ?? undefined);
    p.footer = parseHeaderFooter(firstByLocal(hf, "oddFooter")?.textContent ?? undefined);
  }
  p.rowBreaks = readBreaks(el("rowBreaks"));
  p.colBreaks = readBreaks(el("colBreaks"));

  for (const k of Object.keys(p) as (keyof PrintSetup)[]) if (p[k] === undefined) delete p[k];
  if (Object.keys(p).length) sheet.printSetup = { ...sheet.printSetup, ...p };
}

/** Apply the sheet-scoped _xlnm.Print_Area / Print_Titles names onto their sheets. */
export function readXlsxPrintNames(wb: Workbook, wbDoc: Document): void {
  for (const dn of Array.from(wbDoc.getElementsByTagName("definedName"))) {
    const name = dn.getAttribute("name");
    if (name !== "_xlnm.Print_Area" && name !== "_xlnm.Print_Titles") continue;
    // These are sheet-scoped: localSheetId indexes the workbook's sheet order.
    const idx = Number(dn.getAttribute("localSheetId") ?? "-1");
    const sheet = wb.sheets[idx];
    const value = dn.textContent?.trim();
    if (!sheet || !value) continue;
    const p: PrintSetup = sheet.printSetup ?? {};
    if (name === "_xlnm.Print_Area") {
      const areas = value.split(",").map(parseAreaRef).filter((r): r is NonNullable<typeof r> => !!r);
      if (areas.length) p.printArea = areas;
    } else {
      const t = parseTitles(value);
      if (t.rows) p.titleRows = t.rows;
      if (t.cols) p.titleCols = t.cols;
    }
    sheet.printSetup = p;
  }
}
