import { parseXmlOpt, type Sheet, type Workbook } from "../../core/model";
import type { ChartAnchor, ChartKind, ChartModel, ChartRef, ChartSeries } from "../../core/chart-model";

// Read ODS charts: each is an embedded OpenDocument object (its own content.xml with chart:chart)
// referenced from a draw:frame in the sheet's table. We resolve the object, parse the chart into
// the shared ChartModel, and anchor it to the frame's cell. Data ranges are converted from ODS
// addresses ("Sheet1.B2:Sheet1.B4") to the model's "Sheet1!B2:B4" form. Best-effort anchor: the
// containing cell + svg offsets, with table:end-cell-address for the far corner when present.

const kids = (el: Element, local: string): Element[] => Array.from(el.children).filter((c) => c.localName === local);
const kid = (el: Element | undefined, local: string): Element | undefined => (el ? kids(el, local)[0] : undefined);
export const descend = (root: Element, local: string): Element[] => Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);
export const attrByLocal = (el: Element, local: string): string | null => {
  for (const at of Array.from(el.attributes)) if (at.localName === local) return at.value;
  return null;
};
const A = attrByLocal;

/** Parse an ODS length ("5cm", "127pt", "2in", "48px") to pixels at 96 dpi. */
function lenPx(v: string | null): number {
  if (!v) return 0;
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(v.trim());
  if (!m) return 0;
  const n = Number(m[1]);
  switch (m[2]) {
    case "cm": return n * 37.795;
    case "mm": return n * 3.7795;
    case "in": return n * 96;
    case "pt": return n * (96 / 72);
    case "pc": return n * 16;
    default: return n; // px or unitless
  }
}

/** "Sheet1.B2:Sheet1.B4" / ".$B$2:.$B$4" / "$Sheet1.$B$2:$Sheet1.$B$4" -> "Sheet1!B2:B4". */
function odsRef(addr: string | null): string | undefined {
  if (!addr) return undefined;
  const part = (s: string): { sheet?: string; cell: string } => {
    const m = /^\$?([^.]*)\.(.+)$/.exec(s);
    return m ? { sheet: m[1] ? m[1].replace(/^\$/, "") : undefined, cell: m[2].replace(/\$/g, "") } : { cell: s.replace(/\$/g, "") };
  };
  const [a, b] = addr.split(":");
  const pa = part(a);
  const pb = b ? part(b) : pa;
  const sheet = pa.sheet ?? pb.sheet;
  const range = b ? `${pa.cell}:${pb.cell}` : pa.cell;
  return sheet ? `${sheet}!${range}` : range;
}
const asRef = (addr: string | null): ChartRef | undefined => { const r = odsRef(addr); return r ? { ref: r } : undefined; };

const CLASS_KIND: Record<string, ChartKind> = {
  bar: "column", line: "line", area: "area", circle: "pie", ring: "doughnut",
  scatter: "scatter", bubble: "bubble", radar: "radar", "filled-radar": "radar",
};

const A1 = (addr: string): { row: number; col: number } | null => {
  const m = /([A-Z]+)(\d+)/.exec(addr.split("!").pop() ?? addr);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]), col };
};

function colIndexOf(rowEl: Element, cellEl: Element): number {
  let idx = 0;
  for (const ch of Array.from(rowEl.children)) {
    if (ch === cellEl) break;
    if (ch.localName === "table-cell" || ch.localName === "covered-table-cell") idx += Number(A(ch, "number-columns-repeated") || "1");
  }
  return idx + 1;
}
function rowIndexOf(table: Element, rowEl: Element): number {
  let idx = 0;
  let found = false;
  const walk = (node: Element): void => {
    for (const ch of Array.from(node.children)) {
      if (found) return;
      if (ch.localName === "table-row") { if (ch === rowEl) { found = true; return; } idx += Number(A(ch, "number-rows-repeated") || "1"); }
      else if (ch.localName === "table-header-rows" || ch.localName === "table-rows") walk(ch);
    }
  };
  walk(table);
  return idx + 1;
}

export function anchorOf(frame: Element, table: Element): ChartAnchor {
  const cell = frame.closest ? (frame.closest("*") as Element) : frame.parentElement;
  let fromCol = 1;
  let fromRow = 1;
  // Find the containing table-cell and its row.
  let el: Element | null = frame.parentElement;
  while (el && el.localName !== "table-cell" && el.localName !== "covered-table-cell") el = el.parentElement;
  const rowEl = el?.parentElement;
  if (el && rowEl && rowEl.localName === "table-row") { fromCol = colIndexOf(rowEl, el); fromRow = rowIndexOf(table, rowEl); }
  void cell;
  const fromColOff = lenPx(A(frame, "x"));
  const fromRowOff = lenPx(A(frame, "y"));
  const endAddr = A(frame, "end-cell-address");
  const end = endAddr ? A1(endAddr.replace(/\$/g, "")) : null;
  if (end) return { fromCol, fromRow, fromColOff, fromRowOff, toCol: end.col, toRow: end.row, toColOff: lenPx(A(frame, "end-x")), toRowOff: lenPx(A(frame, "end-y")) };
  const w = lenPx(A(frame, "width"));
  const h = lenPx(A(frame, "height"));
  return { fromCol, fromRow, fromColOff, fromRowOff, toCol: fromCol, toRow: fromRow, toColOff: fromColOff + w, toRowOff: fromRowOff + h };
}

function parseOdsChart(objectDoc: Document, anchor: ChartAnchor, id: string, objectDir: string): ChartModel | null {
  // Both office:chart (wrapper) and chart:chart share the local name "chart"; pick the inner one
  // (it carries the class / plot-area).
  const chart = descend(objectDoc.documentElement, "chart").find((c) => A(c, "class") || kid(c, "plot-area"));
  if (!chart) return null;
  const cls = (A(chart, "class") || "").replace(/^chart:/, "");
  let kind = CLASS_KIND[cls] ?? "column";
  const plot = kid(chart, "plot-area");
  if (!plot) return null;
  // Resolve the plot-area's chart-properties style: bar orientation (vertical) + stacked/percentage.
  const styleProps = new Map<string, Element>();
  for (const st of descend(objectDoc.documentElement, "style")) { const nm = A(st, "name"); const cp = kid(st, "chart-properties"); if (nm && cp) styleProps.set(nm, cp); }
  const props = styleProps.get(A(plot, "style-name") ?? A(chart, "style-name") ?? "");
  let stacked = false;
  let percent = false;
  let labels: ChartModel["labels"];
  if (props) {
    const v = A(props, "vertical");
    if (cls === "bar" && v === "false") kind = "bar"; // horizontal bars
    else if (cls === "bar" && v === "true") kind = "column";
    if (A(props, "stacked") === "true") stacked = true;
    if (A(props, "percentage") === "true") { stacked = true; percent = true; }
    const num = A(props, "data-label-number");
    const txt = A(props, "data-label-text");
    const spec: ChartModel["labels"] = {};
    if (num === "value" || num === "value-and-percentage") spec!.value = true;
    if (num === "percentage" || num === "value-and-percentage") spec!.percent = true;
    if (txt === "true") spec!.category = true;
    if (Object.keys(spec!).length) labels = spec;
  }
  // Axes: a secondary value axis is a second chart:axis of dimension "y"; series attach to one by name.
  const yAxisNames = kids(plot, "axis").filter((ax) => A(ax, "dimension") === "y").map((ax) => A(ax, "name"));
  const primaryY = yAxisNames[0] ?? null;
  const series = kids(plot, "series").map((s) => {
    const nameRef = asRef(A(s, "label-cell-address"));
    const sCls = (A(s, "class") || "").replace(/^chart:/, "");
    const sKind = sCls ? CLASS_KIND[sCls] : undefined;
    const out: ChartSeries = { name: nameRef, values: asRef(A(s, "values-cell-range-address")) ?? { cache: [] }, xValues: undefined };
    if (sKind && sKind !== kind) out.type = sKind; // combo: series overrides the chart class
    const attached = A(s, "attached-axis");
    if (attached && primaryY && attached !== primaryY) out.secondaryAxis = true;
    return out;
  });
  const catsEl = kid(plot, "categories");
  const categories = catsEl ? asRef(A(catsEl, "cell-range-address")) : undefined;
  const titleEl = kid(chart, "title");
  const title = titleEl ? (descend(titleEl, "p").map((p) => p.textContent ?? "").join("").trim() || undefined) : undefined;
  const legendEl = kid(chart, "legend");
  const posMap: Record<string, "top" | "bottom" | "left" | "right"> = { top: "top", bottom: "bottom", start: "left", end: "right" };
  return {
    id,
    kind,
    stacked: stacked || undefined,
    percent: percent || undefined,
    title,
    legend: { show: !!legendEl, pos: posMap[A(legendEl ?? chart, "legend-position") ?? "end"] ?? "right" },
    dataLabels: labels?.value || undefined,
    labels,
    categories,
    series,
    anchor,
    original: { objectDir },
  };
}

/** Populate each sheet's charts from the ODS content.xml draw:frame chart objects. */
export function readOdsCharts(wb: Workbook, files: Record<string, Uint8Array>): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  const frames = descend(doc.documentElement, "frame");
  let n = 0;
  for (const frame of frames) {
    const obj = kids(frame, "object").find((o) => (A(o, "href") || "").length);
    if (!obj) continue;
    const href = (A(obj, "href") || "").replace(/^\.\//, "").replace(/\/$/, "");
    // The frame's sheet: nearest ancestor table:table.
    let t: Element | null = frame.parentElement;
    while (t && t.localName !== "table") t = t.parentElement;
    const sheet: Sheet | undefined = t ? wb.sheets.find((s) => s.tableEl === t) : undefined;
    if (!sheet || !t) continue;
    const objPath = `${href}/content.xml`;
    const objDoc = files[objPath] ? parseXmlOpt(files[objPath]) : undefined;
    if (!objDoc) continue;
    const model = parseOdsChart(objDoc, anchorOf(frame, t), `chart-${++n}`, href);
    if (model) (sheet.charts ??= []).push(model);
  }
}
