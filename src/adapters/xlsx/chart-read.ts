import { parseXmlOpt, type Sheet } from "../../core/model";
import { emuToPx, type ChartAnchor, type ChartDataLabels, type ChartErrorBars, type ChartKind, type ChartModel, type ChartRef, type ChartSeries, type ChartTextStyle, type ChartTrendline } from "../../core/chart-model";

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
export function resolvePart(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts: string[] = [];
  for (const seg of `${base}/${target}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
  return parts.join("/");
}

export function relMap(files: Record<string, Uint8Array>, relsPath: string): { byId: Map<string, string>; byType: { id: string; type: string; target: string }[] } {
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
  // Multi-level categories (c:multiLvlStrRef): render the innermost level; write flattens it.
  const multi = kid(container, "multiLvlStrRef");
  if (multi) {
    const f = textOf(kid(multi, "f")).trim();
    const cacheEl = kid(multi, "multiLvlStrCache");
    const lvl = cacheEl ? kids(cacheEl, "lvl")[0] : undefined;
    return { ref: f || undefined, cache: lvl ? ptsOf(lvl) : undefined };
  }
  const lit = kid(container, "v");
  return lit ? { cache: [textOf(lit)] } : undefined;
}

const TREND_TYPE: Record<string, ChartTrendline["type"]> = { linear: "linear", exp: "exp", log: "log", poly: "poly", power: "power", movingAvg: "movingAvg" };
/** A c:trendline on a series -> ChartTrendline. */
function readTrendline(ser: Element): ChartTrendline | undefined {
  const t = kid(ser, "trendline");
  if (!t) return undefined;
  const ty = attr(kid(t, "trendlineType"), "val");
  const type = ty ? TREND_TYPE[ty] : undefined;
  if (!type) return undefined;
  const num = (n: string): number | undefined => { const v = attr(kid(t, n), "val"); return v != null ? Number(v) : undefined; };
  const out: ChartTrendline = { type };
  const order = num("order"); if (order != null) out.order = order;
  const period = num("period"); if (period != null) out.order = period; // movingAvg reuses order for the period
  const fw = num("forward"); if (fw != null) out.forward = fw;
  const bw = num("backward"); if (bw != null) out.backward = bw;
  const ic = num("intercept"); if (ic != null) out.intercept = ic;
  if (attr(kid(t, "dispEq"), "val") === "1") out.dispEq = true;
  if (attr(kid(t, "dispRSqr"), "val") === "1") out.dispRSqr = true;
  const nm = textOf(kid(t, "name")).trim(); if (nm) out.name = nm;
  const col = colorOf(t); if (col) out.color = col;
  return out;
}

/** A c:errBars on a series -> ChartErrorBars. */
function readErrorBars(ser: Element): ChartErrorBars | undefined {
  const eb = kid(ser, "errBars");
  if (!eb) return undefined;
  const valueType = (attr(kid(eb, "errValType"), "val") as ChartErrorBars["valueType"]) ?? "fixedVal";
  const out: ChartErrorBars = { valueType };
  const dir = attr(kid(eb, "errBarType"), "val"); if (dir === "both" || dir === "plus" || dir === "minus") out.direction = dir;
  const v = attr(kid(eb, "val"), "val"); if (v != null) out.value = Number(v);
  if (attr(kid(eb, "noEndCap"), "val") === "1") out.noEndCap = true;
  const side = (name: string): (number | null)[] | undefined => {
    const el = kid(eb, name);
    const cacheEl = el ? (kid(el, "numLit") ?? (kid(el, "numRef") ? kid(kid(el, "numRef"), "numCache") : undefined)) : undefined;
    return cacheEl ? ptsOf(cacheEl).map((x) => (typeof x === "number" ? x : x == null ? null : Number(x))) : undefined;
  };
  if (valueType === "cust") { out.plus = side("plus"); out.minus = side("minus"); }
  return out;
}

/** A c:dLbls (chart-level or per-series) -> the content/position flags that are set. */
function readDLbls(parent: Element | undefined): ChartDataLabels | undefined {
  const d = kid(parent, "dLbls");
  if (!d) return undefined;
  const on = (n: string): boolean => attr(kid(d, n), "val") === "1";
  const spec: ChartDataLabels = {};
  if (on("showVal")) spec.value = true;
  if (on("showCatName")) spec.category = true;
  if (on("showSerName")) spec.seriesName = true;
  if (on("showPercent")) spec.percent = true;
  if (on("showLegendKey")) spec.legendKey = true;
  const pos = attr(kid(d, "dLblPos"), "val");
  if (pos) spec.position = pos;
  return Object.keys(spec).length ? spec : undefined;
}

const CHART_ELEMS: { local: string; kind: ChartKind }[] = [
  { local: "barChart", kind: "column" }, { local: "bar3DChart", kind: "column" },
  { local: "lineChart", kind: "line" }, { local: "line3DChart", kind: "line" },
  { local: "areaChart", kind: "area" }, { local: "area3DChart", kind: "area" },
  { local: "pieChart", kind: "pie" }, { local: "pie3DChart", kind: "pie" },
  { local: "ofPieChart", kind: "pie" },
  { local: "doughnutChart", kind: "doughnut" },
  { local: "scatterChart", kind: "scatter" }, { local: "bubbleChart", kind: "bubble" },
  { local: "radarChart", kind: "radar" },
  { local: "stockChart", kind: "stock" },
  { local: "surfaceChart", kind: "surface" }, { local: "surface3DChart", kind: "surface" },
];

/** The fill colour of an element's OWN spPr (not a descendant's), as CSS. Resolves srgbClr,
    schemeClr (via the theme), and the first gradient stop; applies a lumMod/lumOff tint. */
function colorOf(el: Element, theme: Record<string, string> = {}): string | undefined {
  const spPr = kid(el, "spPr");
  if (!spPr) return undefined;
  // Fill colour, or (for line series) the line colour in a:ln/solidFill.
  const fill = kid(spPr, "solidFill") ?? kid(spPr, "gradFill") ?? (kid(spPr, "ln") ? kid(kid(spPr, "ln"), "solidFill") : undefined);
  if (!fill) return undefined;
  const clr = descend(fill, "srgbClr")[0] ?? descend(fill, "schemeClr")[0];
  if (!clr) return undefined;
  let hex = clr.localName === "srgbClr" ? `#${clr.getAttribute("val")}` : theme[clr.getAttribute("val") ?? ""];
  if (!hex) return undefined;
  // Apply a luminance modulation/offset tint if present (common on theme colours).
  const lm = descend(clr, "lumMod")[0]; const lo = descend(clr, "lumOff")[0];
  if (lm || lo) hex = applyLum(hex, lm ? Number(lm.getAttribute("val")) / 100000 : 1, lo ? Number(lo.getAttribute("val")) / 100000 : 0);
  return hex;
}
function applyLum(hex: string, mod: number, off: number): string {
  const h = hex.replace("#", "");
  const ch = (i: number): number => Math.max(0, Math.min(255, Math.round(parseInt(h.slice(i, i + 2), 16) * mod + off * 255)));
  return `#${[0, 2, 4].map((i) => ch(i).toString(16).padStart(2, "0")).join("")}`;
}

/** Text styling (a:rPr on a run, or a:defRPr in a txPr) within a title/txPr container. */
function readTextStyle(container: Element | undefined, theme: Record<string, string>): ChartTextStyle | undefined {
  if (!container) return undefined;
  const rpr = descend(container, "rPr")[0] ?? descend(container, "defRPr")[0];
  if (!rpr) return undefined;
  const out: ChartTextStyle = {};
  const sz = attr(rpr, "sz"); if (sz != null) out.size = Number(sz) / 100;
  if (attr(rpr, "b") === "1") out.bold = true;
  if (attr(rpr, "i") === "1") out.italic = true;
  const fill = kid(rpr, "solidFill");
  const clr = fill ? (descend(fill, "srgbClr")[0] ?? descend(fill, "schemeClr")[0]) : undefined;
  if (clr) { const hex = clr.localName === "srgbClr" ? `#${clr.getAttribute("val")}` : theme[clr.getAttribute("val") ?? ""]; if (hex) out.color = hex; }
  const tf = descend(rpr, "latin")[0]?.getAttribute("typeface"); if (tf) out.font = tf;
  return Object.keys(out).length ? out : undefined;
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

function parseChart(chartDoc: Document, anchor: ChartAnchor, id: string, original: ChartModel["original"], theme: Record<string, string>): ChartModel | null {
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
  const percent = grouping === "percentStacked";
  let categories: ChartRef | undefined;
  let categoryLevels: (string | number | null)[][] | undefined;
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
        color: colorOf(ser, theme),
      };
      if (k === "scatter" || k === "bubble") { s.xValues = refOf(kid(ser, "xVal")); s.values = refOf(kid(ser, "yVal")) ?? s.values; }
      if (k === "bubble") s.sizes = refOf(kid(ser, "bubbleSize"));
      if (attr(kid(ser, "smooth"), "val") === "1") s.smooth = true;
      const mk = kid(ser, "marker");
      if (mk) { const sym = attr(kid(mk, "symbol"), "val"); const sz = attr(kid(mk, "size"), "val"); if (sym || sz) s.marker = { symbol: sym ?? undefined, size: sz != null ? Number(sz) : undefined }; }
      const dpts = kids(ser, "dPt");
      if (dpts.length) {
        const pc: (string | undefined)[] = [];
        const expl: (number | undefined)[] = [];
        for (const dp of dpts) {
          const j = Number(attr(kid(dp, "idx"), "val") || "0");
          const col = colorOf(dp, theme); if (col) pc[j] = col;
          const e = attr(kid(dp, "explosion"), "val"); if (e != null) expl[j] = Number(e);
        }
        if (pc.some(Boolean)) s.pointColors = pc;
        if (expl.some((v) => v != null)) s.explosion = expl;
      }
      const sl = readDLbls(ser); if (sl) s.labels = sl;
      const tl = readTrendline(ser); if (tl) s.trendline = tl;
      const eb = readErrorBars(ser); if (eb) s.errorBars = eb;
      const ln = kid(kid(ser, "spPr"), "ln");
      if (ln) {
        const w = attr(ln, "w"); if (w != null) s.lineWidth = Math.round(Number(w) / 12700 * 100) / 100;
        const pd = attr(kid(ln, "prstDash"), "val"); if (pd) s.dash = pd;
      }
      if (ti > 0) s.type = k; // combo: series from a non-base type element carry their kind
      if (isSecondary) s.secondaryAxis = true;
      if (!categories) categories = refOf(kid(ser, "cat"));
      if (!categoryLevels) {
        const ml = kid(kid(ser, "cat"), "multiLvlStrRef");
        const ce = ml ? kid(ml, "multiLvlStrCache") : undefined;
        if (ce) { const lv = kids(ce, "lvl").map((l) => ptsOf(l)); if (lv.length > 1) categoryLevels = lv; }
      }
      collected.push({ idx: Number(attr(kid(ser, "idx"), "val") || String(collected.length)), s });
    }
  });
  const series = collected.sort((a, b) => a.idx - b.idx).map((x) => x.s);
  const el = el0;
  const doughEl = typeEls.find((e) => e.localName === "doughnutChart");
  const barEl = typeEls.find((e) => e.localName === "barChart" || e.localName === "bar3DChart");
  const pieEl = typeEls.find((e) => e.localName === "pieChart" || e.localName === "doughnutChart");
  const numAttr = (e: Element | undefined, name: string): number | undefined => { const v = attr(kid(e, name), "val"); return v != null ? Number(v) : undefined; };
  const legendEl = kid(chart, "legend");
  const model: ChartModel = {
    id,
    kind,
    stacked: stacked || undefined,
    percent: percent || undefined,
    blanksAs: (attr(kid(chart, "dispBlanksAs"), "val") as ChartModel["blanksAs"]) || undefined,
    holeSize: numAttr(doughEl, "holeSize"),
    gapWidth: numAttr(barEl, "gapWidth"),
    overlap: numAttr(barEl, "overlap"),
    rotation: numAttr(pieEl, "firstSliceAng"),
    titleStyle: readTextStyle(kid(chart, "title"), theme),
    legendStyle: readTextStyle(kid(legendEl, "txPr"), theme),
    plotFill: colorOf(plot, theme),
    areaFill: colorOf(space, theme),
    threeD: (kind !== "surface" && /3DChart$/.test(el0.localName)) || undefined,
    ofPie: ((): ChartModel["ofPie"] => {
      const op = typeEls.find((e) => e.localName === "ofPieChart");
      if (!op) return undefined;
      const type = attr(kid(op, "ofPieType"), "val") === "bar" ? "bar" : "pie";
      const splitType = attr(kid(op, "splitType"), "val");
      return { type, splitCount: splitType === "pos" ? numAttr(op, "splitPos") : undefined, secondSize: numAttr(op, "secondPieSize"), gapWidth: numAttr(op, "gapWidth") };
    })(),
    title: titleText(chart),
    legend: { show: !!legendEl, pos: LEGEND_POS[attr(kid(legendEl, "legendPos"), "val") ?? "r"] ?? "right" },
    categories,
    categoryLevels,
    series,
    anchor,
    original,
  };
  if (legendEl) {
    const del = kids(legendEl, "legendEntry").filter((le) => attr(kid(le, "delete"), "val") === "1").map((le) => Number(attr(kid(le, "idx"), "val") || "0"));
    if (del.length) model.legend!.deleted = del;
    if (attr(kid(legendEl, "overlay"), "val") === "1") model.legend!.overlay = true;
  }
  const catAxTitle = titleText(kid(plot, "catAx") ?? space);
  const valAxTitle = titleText(kid(plot, "valAx") ?? space);
  const bounds = (ax: Element | undefined): { min?: number; max?: number } | undefined => {
    const sc = ax ? kid(ax, "scaling") : undefined;
    if (!sc) return undefined;
    const mn = attr(kid(sc, "min"), "val");
    const mx = attr(kid(sc, "max"), "val");
    return mn != null || mx != null ? { min: mn != null ? Number(mn) : undefined, max: mx != null ? Number(mx) : undefined } : undefined;
  };
  const yb = bounds(kid(plot, "valAx"));
  const yFmt = attr(kid(kid(plot, "valAx"), "numFmt"), "formatCode") ?? undefined;
  const xDate = !!kid(plot, "dateAx");
  const catAxEl = kid(plot, "catAx") ?? kid(plot, "dateAx");
  const valAxEl = kid(plot, "valAx");
  const xLabelStyle = readTextStyle(kid(catAxEl, "txPr"), theme);
  const xTitleStyle = readTextStyle(kid(catAxEl, "title"), theme);
  const yLabelStyle = readTextStyle(kid(valAxEl, "txPr"), theme);
  const yTitleStyle = readTextStyle(kid(valAxEl, "title"), theme);
  const anyX = catAxTitle || xDate || xLabelStyle || xTitleStyle;
  const anyY = valAxTitle || yb || yFmt || yLabelStyle || yTitleStyle;
  if (anyX || anyY) model.axes = {
    x: anyX ? { title: catAxTitle, date: xDate || undefined, labelStyle: xLabelStyle, titleStyle: xTitleStyle } : undefined,
    y: anyY ? { title: valAxTitle, min: yb?.min, max: yb?.max, numFmt: yFmt, labelStyle: yLabelStyle, titleStyle: yTitleStyle } : undefined,
  };
  const chartLabels = readDLbls(el);
  if (chartLabels) model.labels = chartLabels;
  if (chartLabels?.value || descend(el, "showVal").some((v) => attr(v, "val") === "1")) model.dataLabels = true;
  return model;
}

export function anchorOf(anchorEl: Element): ChartAnchor | null {
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

/** The chart's colour-style part (colors1.xml) base colours, resolved via the theme; falls back to
    the theme accent colours. Used as the default series palette when series have no explicit colour. */
function readPalette(files: Record<string, Uint8Array>, chartPath: string, theme: Record<string, string>): string[] | undefined {
  const relsPath = chartPath.replace(/charts\/(chart[^/]+\.xml)$/i, "charts/_rels/$1.rels");
  const csRel = relMap(files, relsPath).byType.find((r) => /chartColorStyle|colors/i.test(r.type) || /colors\d+\.xml$/i.test(r.target));
  if (csRel) {
    const csPath = resolvePart(chartPath.replace(/\/[^/]+$/, ""), csRel.target);
    const doc = files[csPath] ? parseXmlOpt(files[csPath]) : undefined;
    if (doc) {
      const colors = Array.from(doc.documentElement.children).filter((e) => e.localName === "srgbClr" || e.localName === "schemeClr");
      const out = colors.map((clr) => (clr.localName === "srgbClr" ? `#${clr.getAttribute("val")}` : theme[clr.getAttribute("val") ?? ""])).filter(Boolean) as string[];
      if (out.length) return out;
    }
  }
  const accents = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((k) => theme[k]).filter(Boolean) as string[];
  return accents.length ? accents : undefined;
}

/** Populate sheet.charts from the worksheet's drawing + chart parts. */
export function readCharts(sheet: Sheet, files: Record<string, Uint8Array>, path: string, theme: Record<string, string> = {}): void {
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
      const model = parseChart(chartDoc, anchor, `chart-${out.length + 1}`, { partPath: chartPath, drawingPath: drawPath }, theme);
      if (model) { const pal = readPalette(files, chartPath, theme); if (pal) model.palette = pal; out.push(model); }
    }
  }
  if (out.length) sheet.charts = out;
}
