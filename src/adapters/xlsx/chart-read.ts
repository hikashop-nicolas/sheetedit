import { parseXmlOpt, type Sheet } from "../../core/model";
import { emuToPx, type ChartAnchor, type ChartKind, type ChartModel, type ChartRef, type ChartSeries } from "../../core/chart-model";

// Read the charts anchored on a worksheet: sheet rels -> drawingN.xml (the anchors) -> chartN.xml
// (the DrawingML chart) -> ChartModel. Namespace-prefix-agnostic (elements are matched by local
// name). Only the anchor + chart parts are needed; both are preserved verbatim on save unless the
// chart is edited.

const kids = (el: Element, local: string): Element[] => Array.from(el.children).filter((c) => c.localName === local);
const kid = (el: Element | undefined, local: string): Element | undefined => (el ? kids(el, local)[0] : undefined);
const descend = (root: Element | Document, local: string): Element[] => Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);
const textOf = (el: Element | undefined): string => el?.textContent ?? "";
const attr = (el: Element | undefined, name: string): string | null => el?.getAttribute(name) ?? null;

/** Resolve a package part path relative to a base directory, collapsing "../" and "./". */
function resolvePart(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts: string[] = [];
  for (const seg of `${base}/${target}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
  return parts.join("/");
}

function relMap(files: Record<string, Uint8Array>, relsPath: string): { byId: Map<string, string>; byType: { id: string; type: string; target: string }[] } {
  const byId = new Map<string, string>();
  const byType: { id: string; type: string; target: string }[] = [];
  const doc = files[relsPath] ? parseXmlOpt(files[relsPath]) : undefined;
  if (doc) for (const r of Array.from(doc.getElementsByTagName("Relationship"))) {
    const id = r.getAttribute("Id");
    const target = r.getAttribute("Target");
    const type = r.getAttribute("Type") ?? "";
    if (id && target) { byId.set(id, target); byType.push({ id, type, target }); }
  }
  return { byId, byType };
}

function ptsOf(cacheEl: Element): (string | number | null)[] {
  const n = Number(attr(kid(cacheEl, "ptCount"), "val") || "0");
  const arr: (string | number | null)[] = new Array(Math.max(0, n)).fill(null);
  for (const pt of kids(cacheEl, "pt")) {
    const idx = Number(pt.getAttribute("idx") || "0");
    const v = textOf(kid(pt, "v"));
    if (idx >= arr.length) arr.length = idx + 1;
    arr[idx] = /^-?\d+(?:\.\d+)?$/.test(v) ? Number(v) : v;
  }
  return arr;
}

/** A c:cat / c:val / c:tx / c:xVal container -> a ChartRef (formula + cached points, or a literal). */
function refOf(container: Element | undefined): ChartRef | undefined {
  if (!container) return undefined;
  const numRef = kid(container, "numRef") ?? kid(container, "strRef");
  if (numRef) {
    const f = textOf(kid(numRef, "f")).trim();
    const cacheEl = kid(numRef, "numCache") ?? kid(numRef, "strCache");
    return { ref: f || undefined, cache: cacheEl ? ptsOf(cacheEl) : undefined };
  }
  const lit = kid(container, "v");
  return lit ? { cache: [textOf(lit)] } : undefined;
}

const CHART_ELEMS: { local: string; kind: ChartKind }[] = [
  { local: "barChart", kind: "column" }, { local: "bar3DChart", kind: "column" },
  { local: "lineChart", kind: "line" }, { local: "line3DChart", kind: "line" },
  { local: "areaChart", kind: "area" }, { local: "area3DChart", kind: "area" },
  { local: "pieChart", kind: "pie" }, { local: "pie3DChart", kind: "pie" },
  { local: "doughnutChart", kind: "doughnut" },
  { local: "scatterChart", kind: "scatter" }, { local: "bubbleChart", kind: "bubble" },
  { local: "radarChart", kind: "radar" },
];

function colorOf(ser: Element): string | undefined {
  const srgb = descend(ser, "srgbClr")[0];
  return srgb ? `#${srgb.getAttribute("val")}` : undefined;
}

function titleText(chart: Element): string | undefined {
  const title = kid(chart, "title");
  if (!title) return undefined;
  const runs = descend(title, "t").map((t) => t.textContent ?? "");
  if (runs.length) return runs.join("").trim() || undefined;
  const cachePts = descend(title, "pt").map((p) => textOf(kid(p, "v")));
  return cachePts.join("").trim() || undefined;
}

const LEGEND_POS: Record<string, "top" | "bottom" | "left" | "right"> = { t: "top", b: "bottom", l: "left", r: "right", tr: "right" };

function parseChart(chartDoc: Document, anchor: ChartAnchor, id: string, original: ChartModel["original"]): ChartModel | null {
  const space = chartDoc.documentElement;
  const chart = kid(space, "chart");
  const plot = chart && kid(chart, "plotArea");
  if (!plot) return null;
  // A combo chart has several chart-type elements (e.g. barChart + lineChart) in one plotArea.
  const typeEls = Array.from(plot.children).filter((e) => CHART_ELEMS.some((c) => c.local === e.localName));
  if (!typeEls.length) return null;
  const kindOf = (e: Element): ChartKind => {
    const base = CHART_ELEMS.find((c) => c.local === e.localName)!.kind;
    return e.localName.startsWith("bar") ? (attr(kid(e, "barDir"), "val") === "bar" ? "bar" : "column") : base;
  };
  // Secondary value axis: a second valAx, or one anchored on the right.
  const valAxes = kids(plot, "valAx").map((ax) => ({ id: attr(kid(ax, "axId"), "val"), pos: attr(kid(ax, "axPos"), "val") }));
  const secondaryValId = valAxes.length > 1 ? (valAxes.find((a) => a.pos === "r")?.id ?? valAxes[1].id) : null;

  const kind = kindOf(typeEls[0]);
  const grouping = attr(kid(typeEls[0], "grouping"), "val");
  const stacked = grouping === "stacked" || grouping === "percentStacked";
  let categories: ChartRef | undefined;
  const collected: { idx: number; s: ChartSeries }[] = [];
  const el0 = typeEls[0]; // for data-label detection
  typeEls.forEach((el, ti) => {
    const k = kindOf(el);
    const axIds = kids(el, "axId").map((a) => attr(a, "val"));
    const isSecondary = !!secondaryValId && axIds[axIds.length - 1] === secondaryValId;
    for (const ser of kids(el, "ser")) {
      const nameRef = refOf(kid(ser, "tx"));
      const s: ChartSeries = {
        name: nameRef?.ref ? nameRef : (nameRef?.cache ? String(nameRef.cache[0] ?? "") : undefined),
        values: refOf(kid(ser, "val")) ?? refOf(kid(ser, "yVal")) ?? { cache: [] },
        color: colorOf(ser),
      };
      if (k === "scatter" || k === "bubble") { s.xValues = refOf(kid(ser, "xVal")); s.values = refOf(kid(ser, "yVal")) ?? s.values; }
      if (k === "bubble") s.sizes = refOf(kid(ser, "bubbleSize"));
      if (ti > 0) s.type = k; // combo: series from a non-base type element carry their kind
      if (isSecondary) s.secondaryAxis = true;
      if (!categories) categories = refOf(kid(ser, "cat"));
      collected.push({ idx: Number(attr(kid(ser, "idx"), "val") || String(collected.length)), s });
    }
  });
  const series = collected.sort((a, b) => a.idx - b.idx).map((x) => x.s);
  const el = el0;
  const legendEl = kid(chart, "legend");
  const model: ChartModel = {
    id,
    kind,
    stacked: stacked || undefined,
    title: titleText(chart),
    legend: { show: !!legendEl, pos: LEGEND_POS[attr(kid(legendEl, "legendPos"), "val") ?? "r"] ?? "right" },
    categories,
    series,
    anchor,
    original,
  };
  const catAxTitle = titleText(kid(plot, "catAx") ?? space);
  const valAxTitle = titleText(kid(plot, "valAx") ?? space);
  if (catAxTitle || valAxTitle) model.axes = { x: catAxTitle ? { title: catAxTitle } : undefined, y: valAxTitle ? { title: valAxTitle } : undefined };
  if (descend(el, "showVal").some((v) => attr(v, "val") === "1")) model.dataLabels = true;
  return model;
}

function anchorOf(anchorEl: Element): ChartAnchor | null {
  const from = kid(anchorEl, "from");
  if (from) {
    const to = kid(anchorEl, "to");
    const val = (parent: Element | undefined, name: string): number => Number(textOf(kid(parent, name)) || "0");
    const fromCol = val(from, "col") + 1;
    const fromRow = val(from, "row") + 1;
    const fromColOff = emuToPx(val(from, "colOff"));
    const fromRowOff = emuToPx(val(from, "rowOff"));
    if (to) {
      return { fromCol, fromRow, fromColOff, fromRowOff, toCol: val(to, "col") + 1, toRow: val(to, "row") + 1, toColOff: emuToPx(val(to, "colOff")), toRowOff: emuToPx(val(to, "rowOff")) };
    }
    // oneCellAnchor: from + ext (size). Express the size via the "to" offsets on the same cell.
    const ext = kid(anchorEl, "ext");
    const w = emuToPx(Number(attr(ext, "cx") || "0"));
    const h = emuToPx(Number(attr(ext, "cy") || "0"));
    return { fromCol, fromRow, fromColOff, fromRowOff, toCol: fromCol, toRow: fromRow, toColOff: fromColOff + w, toRowOff: fromRowOff + h };
  }
  // absoluteAnchor: pos + ext (EMU from the sheet origin).
  const pos = kid(anchorEl, "pos");
  const ext = kid(anchorEl, "ext");
  if (pos && ext) {
    const x = emuToPx(Number(attr(pos, "x") || "0"));
    const y = emuToPx(Number(attr(pos, "y") || "0"));
    const w = emuToPx(Number(attr(ext, "cx") || "0"));
    const h = emuToPx(Number(attr(ext, "cy") || "0"));
    return { fromCol: 1, fromRow: 1, fromColOff: x, fromRowOff: y, toCol: 1, toRow: 1, toColOff: x + w, toRowOff: y + h };
  }
  return null;
}

/** Populate sheet.charts from the worksheet's drawing + chart parts. */
export function readCharts(sheet: Sheet, files: Record<string, Uint8Array>, path: string): void {
  const relsPath = path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const { byType } = relMap(files, relsPath);
  const drawings = byType.filter((r) => /drawing/i.test(r.type) && /drawings\//i.test(r.target)).map((r) => resolvePart("xl/worksheets", r.target));
  const out: ChartModel[] = [];
  for (const drawPath of drawings) {
    const drawDoc = files[drawPath] ? parseXmlOpt(files[drawPath]) : undefined;
    if (!drawDoc) continue;
    const drawRels = relMap(files, drawPath.replace(/drawings\/(drawing[^/]+\.xml)$/i, "drawings/_rels/$1.rels")).byId;
    const drawBase = drawPath.replace(/\/[^/]+$/, "");
    // Every anchor kind that carries a graphicFrame -> a chart reference.
    for (const anchorEl of Array.from(drawDoc.documentElement.children)) {
      if (!/Anchor$/.test(anchorEl.localName)) continue;
      const chartRefEl = descend(anchorEl, "chart").find((c) => c.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || c.getAttribute("r:id"));
      const rid = chartRefEl ? (chartRefEl.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ?? chartRefEl.getAttribute("r:id")) : null;
      if (!rid || !drawRels.has(rid)) continue;
      const chartPath = resolvePart(drawBase, drawRels.get(rid)!);
      const chartDoc = files[chartPath] ? parseXmlOpt(files[chartPath]) : undefined;
      const anchor = anchorOf(anchorEl);
      if (!chartDoc || !anchor) continue;
      const model = parseChart(chartDoc, anchor, `chart-${out.length + 1}`, { partPath: chartPath, drawingPath: drawPath });
      if (model) out.push(model);
    }
  }
  if (out.length) sheet.charts = out;
}
