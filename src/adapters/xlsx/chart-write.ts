import { parseXmlOpt, serializeXml, type Sheet, type Workbook } from "../../core/model";
import { pxToEmu, type ChartAxis, type ChartDataLabels, type ChartErrorBars, type ChartModel, type ChartSeries, type ChartTextStyle, type ChartTrendline } from "../../core/chart-model";
import { resolveLabels, resolveNumbers, seriesName } from "../../core/chart-data";

// Write created / edited charts to xlsx DrawingML. Only dirty charts are emitted; a chart read
// from the file and left untouched keeps its original parts verbatim. A created chart gets a new
// chart part + a two-cell anchor in the sheet's drawing (creating the drawing + registrations if
// needed); an edited chart's part is rewritten and its anchor updated in place.

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const strCache = (vals: string[]): string => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("")}</c:strCache>`;
const numCache = (vals: (number | null)[]): string => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (v == null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:numCache>`;
const strRef = (ref: string | undefined, cache: string[]): string => `<c:strRef><c:f>${esc(ref ?? "")}</c:f>${strCache(cache)}</c:strRef>`;
const numRef = (ref: string | undefined, cache: (number | null)[]): string => `<c:numRef><c:f>${esc(ref ?? "")}</c:f>${numCache(cache)}</c:numRef>`;
/** A c:multiLvlStrRef for multi-level category labels (levels innermost-first). */
const multiLvlStrRef = (ref: string | undefined, levels: (string | number | null)[][]): string => {
  const n = Math.max(0, ...levels.map((l) => l.length));
  const lvls = levels.map((l) => `<c:lvl>${l.map((v, i) => (v == null ? "" : `<c:pt idx="${i}"><c:v>${esc(String(v))}</c:v></c:pt>`)).join("")}</c:lvl>`).join("");
  return `<c:multiLvlStrRef><c:f>${esc(ref ?? "")}</c:f><c:multiLvlStrCache><c:ptCount val="${n}"/>${lvls}</c:multiLvlStrCache></c:multiLvlStrRef>`;
};
const catXml = (catRef: string | undefined, catLabels: string[], catLevels?: (string | number | null)[][]): string =>
  catLevels && catLevels.length > 1 ? `<c:cat>${multiLvlStrRef(catRef, catLevels)}</c:cat>` : (catRef ? `<c:cat>${strRef(catRef, catLabels)}</c:cat>` : "");

const fillPr = (hex: string): string => `<c:spPr><a:solidFill><a:srgbClr val="${hex.replace("#", "")}"/></a:solidFill></c:spPr>`;
const srgb = (hex: string): string => `<a:srgbClr val="${hex.replace("#", "")}"/>`;
/** A series c:spPr: solid fill and, when a line width/dash is set, an a:ln (with the colour). */
function serSpPr(s: ChartSeries): string {
  const fill = s.color ? `<a:solidFill>${srgb(s.color)}</a:solidFill>` : "";
  const hasLn = s.lineWidth != null || s.dash;
  const ln = hasLn ? `<a:ln${s.lineWidth != null ? ` w="${Math.round(s.lineWidth * 12700)}"` : ""}>${s.color ? `<a:solidFill>${srgb(s.color)}</a:solidFill>` : ""}${s.dash ? `<a:prstDash val="${s.dash}"/>` : ""}</a:ln>` : "";
  return fill || ln ? `<c:spPr>${fill}${ln}</c:spPr>` : "";
}
const markerXml = (m?: { symbol?: string; size?: number }): string => (m ? `<c:marker><c:symbol val="${m.symbol ?? "circle"}"/>${m.size != null ? `<c:size val="${m.size}"/>` : ""}</c:marker>` : "");
/** A c:dLbls block (chart-level or per-series) from the label content/position flags. */
function dLblsXml(spec?: ChartDataLabels): string {
  if (!spec) return "";
  const b = (v?: boolean): number => (v ? 1 : 0);
  const pos = spec.position ? `<c:dLblPos val="${spec.position}"/>` : "";
  return `<c:dLbls>${pos}<c:showLegendKey val="${b(spec.legendKey)}"/><c:showVal val="${b(spec.value)}"/><c:showCatName val="${b(spec.category)}"/><c:showSerName val="${b(spec.seriesName)}"/><c:showPercent val="${b(spec.percent)}"/><c:showBubbleSize val="0"/></c:dLbls>`;
}
/** c:dPt elements carrying per-point colour and/or slice explosion. */
function dPtsXml(s: ChartSeries): string {
  const n = Math.max(s.pointColors?.length ?? 0, s.explosion?.length ?? 0);
  let out = "";
  for (let j = 0; j < n; j++) {
    const c = s.pointColors?.[j];
    const e = s.explosion?.[j];
    if (c == null && e == null) continue;
    out += `<c:dPt><c:idx val="${j}"/><c:bubble3D val="0"/>${e != null ? `<c:explosion val="${e}"/>` : ""}${c ? fillPr(c) : ""}</c:dPt>`;
  }
  return out;
}
/** A c:trendline (regression overlay) inside a series. */
function trendlineXml(t?: ChartTrendline): string {
  if (!t) return "";
  const name = t.name ? `<c:name>${esc(t.name)}</c:name>` : "";
  const spPr = t.color ? `<c:spPr><a:ln><a:solidFill><a:srgbClr val="${t.color.replace("#", "")}"/></a:solidFill></a:ln></c:spPr>` : "";
  const order = t.type === "poly" && t.order != null ? `<c:order val="${t.order}"/>` : "";
  const period = t.type === "movingAvg" && t.order != null ? `<c:period val="${t.order}"/>` : "";
  const fw = t.forward != null ? `<c:forward val="${t.forward}"/>` : "";
  const bw = t.backward != null ? `<c:backward val="${t.backward}"/>` : "";
  const ic = t.intercept != null ? `<c:intercept val="${t.intercept}"/>` : "";
  const rsq = t.dispRSqr ? `<c:dispRSqr val="1"/>` : "";
  const eq = t.dispEq ? `<c:dispEq val="1"/>` : "";
  return `<c:trendline>${name}${spPr}<c:trendlineType val="${t.type}"/>${order}${period}${fw}${bw}${ic}${rsq}${eq}</c:trendline>`;
}
/** A c:errBars (error bars) inside a series. */
function errBarsXml(eb?: ChartErrorBars): string {
  if (!eb) return "";
  const numLit = (vals?: (number | null)[]): string => (vals ? `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (v == null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:numLit>` : "");
  const cust = eb.valueType === "cust";
  const plus = cust ? `<c:plus>${numLit(eb.plus)}</c:plus>` : "";
  const minus = cust ? `<c:minus>${numLit(eb.minus)}</c:minus>` : "";
  const val = !cust && eb.value != null ? `<c:val val="${eb.value}"/>` : "";
  return `<c:errBars><c:errDir val="y"/><c:errBarType val="${eb.direction ?? "both"}"/><c:errValType val="${eb.valueType}"/><c:noEndCap val="${eb.noEndCap ? 1 : 0}"/>${plus}${minus}${val}</c:errBars>`;
}
function serCategory(wb: Workbook, s: ChartSeries, i: number, catRef: string | undefined, catLabels: string[], catLevels?: (string | number | null)[][]): string {
  const name = seriesName(wb, s.name) ?? `Series ${i + 1}`;
  const nameRef = typeof s.name === "object" ? s.name : undefined;
  const tx = nameRef?.ref ? `<c:tx>${strRef(nameRef.ref, [name])}</c:tx>` : `<c:tx><c:v>${esc(name)}</c:v></c:tx>`;
  const spPr = serSpPr(s);
  const cat = catXml(catRef, catLabels, catLevels);
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${markerXml(s.marker)}${dPtsXml(s)}${dLblsXml(s.labels)}${trendlineXml(s.trendline)}${errBarsXml(s.errorBars)}${cat}<c:val>${numRef(s.values.ref, resolveNumbers(wb, s.values))}</c:val>${s.smooth ? '<c:smooth val="1"/>' : ""}</c:ser>`;
}
function serXY(wb: Workbook, s: ChartSeries, i: number): string {
  const name = seriesName(wb, s.name) ?? `Series ${i + 1}`;
  const nameRef = typeof s.name === "object" ? s.name : undefined;
  const tx = nameRef?.ref ? `<c:tx>${strRef(nameRef.ref, [name])}</c:tx>` : `<c:tx><c:v>${esc(name)}</c:v></c:tx>`;
  const spPr = serSpPr(s);
  const x = `<c:xVal>${numRef(s.xValues?.ref, resolveNumbers(wb, s.xValues))}</c:xVal>`;
  const y = `<c:yVal>${numRef(s.values.ref, resolveNumbers(wb, s.values))}</c:yVal>`;
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${markerXml(s.marker)}${dLblsXml(s.labels)}${trendlineXml(s.trendline)}${errBarsXml(s.errorBars)}${x}${y}<c:smooth val="${s.smooth ? 1 : 0}"/></c:ser>`;
}

/** Run-property attributes shared by a:rPr and a:defRPr. */
const rPrAttrs = (s: ChartTextStyle): string => `${s.size != null ? ` sz="${Math.round(s.size * 100)}"` : ""}${s.bold ? ` b="1"` : ""}${s.italic ? ` i="1"` : ""}`;
const rPrInner = (s: ChartTextStyle): string => `${s.color ? `<a:solidFill><a:srgbClr val="${s.color.replace("#", "")}"/></a:solidFill>` : ""}${s.font ? `<a:latin typeface="${esc(s.font)}"/>` : ""}`;
/** A c:txPr (default run properties) for tick labels / legend text. */
const txPrXml = (s?: ChartTextStyle): string => (s ? `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${rPrAttrs(s)}>${rPrInner(s)}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>` : "");
/** A c:title (rich text) for an axis or the chart, with an optional run style. */
const richTitleXml = (title: string, s?: ChartTextStyle): string => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${s ? rPrAttrs(s) : ""}>${s ? rPrInner(s) : ""}</a:defRPr></a:pPr><a:r><a:rPr lang="en-US"${s ? rPrAttrs(s) : ""}>${s ? rPrInner(s) : ""}</a:rPr><a:t>${esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
const axTitle = (ax?: ChartAxis): string => (ax?.title ? richTitleXml(ax.title, ax.titleStyle) : "");

const catAx = (id: number, cross: number, pos: string, del = false, date = false, ax?: ChartAxis): string => {
  const tag = date ? "dateAx" : "catAx";
  const extra = date ? "<c:baseTimeUnit val=\"days\"/>" : "";
  return `<c:${tag}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${del ? 1 : 0}"/><c:axPos val="${pos}"/>${axTitle(ax)}${txPrXml(ax?.labelStyle)}<c:crossAx val="${cross}"/>${extra}</c:${tag}>`;
};
const valAx = (id: number, cross: number, pos: string, crossesMax = false, ax?: ChartAxis): string => {
  const scale = `<c:orientation val="minMax"/>${ax?.max != null ? `<c:max val="${ax.max}"/>` : ""}${ax?.min != null ? `<c:min val="${ax.min}"/>` : ""}`;
  const fmt = ax?.numFmt ? `<c:numFmt formatCode="${esc(ax.numFmt)}" sourceLinked="0"/>` : "";
  return `<c:valAx><c:axId val="${id}"/><c:scaling>${scale}</c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${axTitle(ax)}${fmt}${txPrXml(ax?.labelStyle)}${crossesMax ? '<c:crosses val="max"/>' : ""}<c:crossAx val="${cross}"/></c:valAx>`;
};
const localOf = (k: string): string => (k === "column" || k === "bar" ? "barChart" : k === "line" ? "lineChart" : "areaChart");
const serAx = (id: number, cross: number): string => `<c:serAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="${cross}"/></c:serAx>`;

/** Body for a 3D chart (bar3D / line3D / area3D / pie3D): the flat model re-emitted as its 3D
    element plus a series axis and a minimal view3D so it stays 3D in Excel. */
function threeDBody(model: ChartModel, wb: Workbook, catRef: string | undefined, catLabels: string[], catLevels: (string | number | null)[][] | undefined, dLbls: string): string {
  const AX1 = 111111111, AX2 = 222222222, AX3 = 333333333;
  const sers = model.series.map((s, i) => serCategory(wb, s, i, catRef, catLabels, catLevels)).join("");
  if (model.kind === "pie") return `<c:pie3DChart><c:varyColors val="1"/>${sers}${dLbls}</c:pie3DChart>`;
  const group = model.percent ? "percentStacked" : model.stacked ? "stacked" : model.kind === "line" || model.kind === "area" ? "standard" : "clustered";
  const axes = catAx(AX1, AX2, model.kind === "bar" ? "l" : "b", false, false, model.axes?.x) + valAx(AX2, AX1, model.kind === "bar" ? "b" : "l", false, model.axes?.y) + serAx(AX3, AX2);
  const ids = `<c:axId val="${AX1}"/><c:axId val="${AX2}"/><c:axId val="${AX3}"/>`;
  if (model.kind === "line") return `<c:line3DChart><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}${dLbls}${ids}</c:line3DChart>${axes}`;
  if (model.kind === "area") return `<c:area3DChart><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}${dLbls}${ids}</c:area3DChart>${axes}`;
  const dir = model.kind === "bar" ? "bar" : "col";
  const spacing = `${model.gapWidth != null ? `<c:gapWidth val="${model.gapWidth}"/>` : ""}`;
  return `<c:bar3DChart><c:barDir val="${dir}"/><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}${dLbls}${spacing}<c:shape val="box"/>${ids}</c:bar3DChart>${axes}`;
}

/** Body for a combo chart (mixed series types and/or a secondary axis): one chart-type element per
    (kind, axis) group, plus the primary axes and a secondary value axis when needed. */
function comboBody(model: ChartModel, wb: Workbook, catRef: string | undefined, catLabels: string[], dLbls: string, catLevels?: (string | number | null)[][]): string {
  const AX1 = 111111111, AX2 = 222222222, AX3 = 333333333, AX4 = 444444444;
  const groups = new Map<string, { kind: string; secondary: boolean; series: ChartSeries[] }>();
  model.series.forEach((s) => {
    const kind = s.type ?? model.kind;
    const secondary = !!s.secondaryAxis;
    const key = `${kind}|${secondary}`;
    (groups.get(key) ?? groups.set(key, { kind, secondary, series: [] }).get(key)!).series.push(s);
  });
  let idx = 0;
  const parts: string[] = [];
  for (const g of groups.values()) {
    const sers = g.series.map((s, j) => serCategory(wb, s, idx + j, catRef, catLabels, catLevels)).join("");
    idx += g.series.length;
    const dir = g.kind === "bar" ? '<c:barDir val="bar"/>' : g.kind === "column" ? '<c:barDir val="col"/>' : "";
    const grp = `<c:grouping val="${model.percent ? "percentStacked" : model.stacked ? "stacked" : g.kind === "line" || g.kind === "area" ? "standard" : "clustered"}"/>`;
    const marker = g.kind === "line" ? '<c:marker val="1"/>' : "";
    const local = localOf(g.kind);
    parts.push(`<c:${local}>${dir}${grp}<c:varyColors val="0"/>${sers}${dLbls}${marker}<c:axId val="${g.secondary ? AX4 : AX1}"/><c:axId val="${g.secondary ? AX3 : AX2}"/></c:${local}>`);
  }
  const hasSecondary = model.series.some((s) => s.secondaryAxis);
  const axes = catAx(AX1, AX2, "b", false, model.axes?.x?.date === true, model.axes?.x) + valAx(AX2, AX1, "l", false, model.axes?.y) + (hasSecondary ? valAx(AX3, AX4, "r", true) + catAx(AX4, AX3, "b", true) : "");
  return parts.join("") + axes;
}

/** Generate the DrawingML chart part for a model, embedding resolved values as caches. */
export function chartXml(model: ChartModel, wb: Workbook): string {
  const catRef = model.categories?.ref;
  const catLabels = resolveLabels(wb, model.categories);
  const catLevels = model.categoryLevels;
  const AX1 = 111111111;
  const AX2 = 222222222;
  const xDate = model.axes?.x?.date === true;
  const dLbls = dLblsXml(model.labels ?? (model.dataLabels ? { value: true } : undefined));
  const isCombo = ["column", "bar", "line", "area"].includes(model.kind) && model.series.some((s) => (s.type && s.type !== model.kind) || s.secondaryAxis);
  const d3 = model.threeD === true && ["column", "bar", "line", "area", "pie"].includes(model.kind) && !isCombo;
  let body: string;
  if (d3) {
    body = threeDBody(model, wb, catRef, catLabels, catLevels, dLbls);
  } else if (isCombo) {
    body = comboBody(model, wb, catRef, catLabels, dLbls, catLevels);
  } else if (model.kind === "scatter" || model.kind === "bubble") {
    const sers = model.series.map((s, i) => serXY(wb, s, i)).join("");
    body = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${dLbls}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:scatterChart>${valAx(AX1, AX2, "b")}${valAx(AX2, AX1, "l", false, model.axes?.y)}`;
  } else if (model.kind === "stock") {
    // High-low-close (3 series) or open-high-low-close (4); the extra open series adds up/down bars.
    const sers = model.series.map((s, i) => serCategory(wb, s, i, catRef, catLabels, catLevels)).join("");
    const upDown = model.series.length >= 4 ? `<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>` : "";
    body = `<c:stockChart>${sers}<c:hiLowLines/>${upDown}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:stockChart>${catAx(AX1, AX2, "b", false, false, model.axes?.x)}${valAx(AX2, AX1, "l", false, model.axes?.y)}`;
  } else if (model.kind === "surface") {
    // No 2D surface; write a wireframe surfaceChart so it round-trips (rendered as a heatmap).
    const sers = model.series.map((s, i) => serCategory(wb, s, i, catRef, catLabels, catLevels)).join("");
    body = `<c:surfaceChart><c:wireframe val="0"/>${sers}<c:axId val="${AX1}"/><c:axId val="${AX2}"/><c:axId val="333333333"/></c:surfaceChart>${catAx(AX1, AX2, "b", false, false, model.axes?.x)}${valAx(AX2, AX1, "l", false, model.axes?.y)}<c:serAx><c:axId val="333333333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="${AX2}"/></c:serAx>`;
  } else {
    const sers = model.series.map((s, i) => serCategory(wb, s, i, catRef, catLabels, catLevels)).join("");
    const group = model.percent ? "percentStacked" : model.stacked ? "stacked" : model.kind === "line" || model.kind === "area" ? "standard" : "clustered";
    if (model.kind === "pie" && model.ofPie) {
      const o = model.ofPie;
      body = `<c:ofPieChart><c:ofPieType val="${o.type}"/><c:varyColors val="1"/>${sers}${dLbls}<c:gapWidth val="${o.gapWidth ?? 100}"/><c:splitType val="pos"/><c:splitPos val="${o.splitCount ?? 2}"/><c:secondPieSize val="${o.secondSize ?? 75}"/><c:serLines/></c:ofPieChart>`;
    } else if (model.kind === "pie" || model.kind === "doughnut") {
      const rot = model.rotation != null ? `<c:firstSliceAng val="${((model.rotation % 360) + 360) % 360}"/>` : "";
      body = `<c:${model.kind}Chart><c:varyColors val="1"/>${sers}${dLbls}${rot}${model.kind === "doughnut" ? `<c:holeSize val="${model.holeSize ?? 50}"/>` : ""}</c:${model.kind}Chart>`;
    } else if (model.kind === "radar") {
      body = `<c:radarChart><c:radarStyle val="marker"/>${sers}${dLbls}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:radarChart>${catAx(AX1, AX2, "b", false, false, model.axes?.x)}${valAx(AX2, AX1, "l", false, model.axes?.y)}`;
    } else if (model.kind === "line") {
      body = `<c:lineChart><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}${dLbls}<c:marker val="1"/><c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:lineChart>${catAx(AX1, AX2, "b", false, xDate, model.axes?.x)}${valAx(AX2, AX1, "l", false, model.axes?.y)}`;
    } else if (model.kind === "area") {
      body = `<c:areaChart><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}${dLbls}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:areaChart>${catAx(AX1, AX2, "b", false, xDate, model.axes?.x)}${valAx(AX2, AX1, "l", false, model.axes?.y)}`;
    } else {
      const dir = model.kind === "bar" ? "bar" : "col";
      const cAxPos = model.kind === "bar" ? "l" : "b";
      const vAxPos = model.kind === "bar" ? "b" : "l";
      const spacing = `${model.gapWidth != null ? `<c:gapWidth val="${model.gapWidth}"/>` : ""}${model.overlap != null ? `<c:overlap val="${model.overlap}"/>` : ""}`;
      body = `<c:barChart><c:barDir val="${dir}"/><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}${dLbls}${spacing}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:barChart>${catAx(AX1, AX2, cAxPos, false, model.kind !== "bar" && xDate, model.axes?.x)}${valAx(AX2, AX1, vAxPos, false, model.axes?.y)}`;
    }
  }
  const title = model.title ? `${richTitleXml(model.title, model.titleStyle)}<c:autoTitleDeleted val="0"/>` : `<c:autoTitleDeleted val="1"/>`;
  const legendEntries = (model.legend?.deleted ?? []).map((i) => `<c:legendEntry><c:idx val="${i}"/><c:delete val="1"/></c:legendEntry>`).join("");
  const legend = model.legend?.show ? `<c:legend><c:legendPos val="${(model.legend.pos ?? "b")[0]}"/>${legendEntries}<c:overlay val="${model.legend.overlay ? 1 : 0}"/>${txPrXml(model.legendStyle)}</c:legend>` : "";
  const blanks = model.blanksAs ? `<c:dispBlanksAs val="${model.blanksAs}"/>` : "";
  const view3D = d3 ? `<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:depthPercent val="100"/><c:rAngAx val="1"/></c:view3D>` : "";
  const plotFill = model.plotFill ? fillPr(model.plotFill) : "";
  const areaFill = model.areaFill ? fillPr(model.areaFill) : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart>${title}${view3D}<c:plotArea><c:layout/>${body}${plotFill}</c:plotArea>${legend}<c:plotVisOnly val="1"/>${blanks}</c:chart>${areaFill}</c:chartSpace>`;
}

function anchorXml(model: ChartModel, chartRid: string, frameId: number): string {
  const a = model.anchor;
  const pt = (tag: string, col: number, colOff: number, row: number, rowOff: number): string =>
    `<xdr:${tag}><xdr:col>${col - 1}</xdr:col><xdr:colOff>${pxToEmu(colOff)}</xdr:colOff><xdr:row>${row - 1}</xdr:row><xdr:rowOff>${pxToEmu(rowOff)}</xdr:rowOff></xdr:${tag}>`;
  return `<xdr:twoCellAnchor>${pt("from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff)}${pt("to", a.toCol, a.toColOff, a.toRow, a.toRowOff)}` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${frameId}" name="Chart ${frameId}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="${chartRid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

export function addContentType(wb: Workbook, partPath: string, ct: string): void {
  const doc = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (!doc || doc.documentElement.localName !== "Types") return;
  if (Array.from(doc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === `/${partPath}`)) return;
  // createElementNS in the content-types namespace so it inherits the root default (no xmlns="").
  const ov = doc.createElementNS(CT_NS, "Override");
  ov.setAttribute("PartName", `/${partPath}`);
  ov.setAttribute("ContentType", ct);
  doc.documentElement.appendChild(ov);
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

/** Add a relationship to a .rels part (creating it), returning the new id. */
export function addRel(wb: Workbook, relsPath: string, type: string, target: string): string {
  let doc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (!doc) doc = parseXmlOpt(new TextEncoder().encode(`<Relationships xmlns="${REL_NS}"></Relationships>`))!;
  const ids = new Set(Array.from(doc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let n = 1;
  while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const rel = doc.createElementNS(REL_NS, "Relationship");
  rel.setAttribute("Id", id);
  rel.setAttribute("Type", type);
  rel.setAttribute("Target", target);
  doc.documentElement.appendChild(rel);
  wb.files[relsPath] = serializeXml(doc);
  return id;
}

const uniquePath = (wb: Workbook, dir: string, base: string, ext: string): { path: string; n: number } => {
  let n = 1;
  while (wb.files[`${dir}/${base}${n}.${ext}`]) n++;
  return { path: `${dir}/${base}${n}.${ext}`, n };
};

/** The drawing part for a sheet (its rels reference it), creating and wiring one if absent. */
export function ensureSheetDrawing(wb: Workbook, sheet: Sheet): string {
  const sheetRels = sheet.path!.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const relsDoc = wb.files[sheetRels] ? parseXmlOpt(wb.files[sheetRels]) : undefined;
  const existing = relsDoc && Array.from(relsDoc.getElementsByTagName("Relationship")).find((r) => /drawing/i.test(r.getAttribute("Type") ?? "") && /drawings\//i.test(r.getAttribute("Target") ?? ""));
  if (existing) {
    const parts: string[] = [];
    for (const seg of `xl/worksheets/${existing.getAttribute("Target")}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    return parts.join("/");
  }
  const { path, n } = uniquePath(wb, "xl/drawings", "drawing", "xml");
  wb.files[path] = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}"></xdr:wsDr>`);
  addContentType(wb, path, CT_DRAWING);
  const rid = addRel(wb, sheetRels, `${R}/drawing`, `../drawings/drawing${n}.xml`);
  // Add <drawing r:id> to the worksheet XML. Ensure xmlns:r is declared on the root so the
  // attribute serializes as the conventional r:id (not a generated ns prefix).
  const wsDoc = parseXmlOpt(wb.files[sheet.path!]);
  const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  if (wsDoc && !Array.from(wsDoc.getElementsByTagName("*")).some((e) => e.localName === "drawing")) {
    if (!wsDoc.documentElement.getAttribute("xmlns:r")) wsDoc.documentElement.setAttribute("xmlns:r", R);
    const d = wsDoc.createElementNS(SS, "drawing");
    d.setAttribute("r:id", rid);
    wsDoc.documentElement.appendChild(d);
    wb.files[sheet.path!] = serializeXml(wsDoc);
    sheet.doc = wsDoc;
    sheet.sheetData = wsDoc.getElementsByTagName("sheetData")[0] ?? sheet.sheetData;
  }
  return path;
}

let frameSeq = 1000;

function createNew(wb: Workbook, sheet: Sheet, model: ChartModel): void {
  const { path: chartPath, n } = uniquePath(wb, "xl/charts", "chart", "xml");
  wb.files[chartPath] = new TextEncoder().encode(chartXml(model, wb));
  addContentType(wb, chartPath, CT_CHART);
  const drawingPath = ensureSheetDrawing(wb, sheet);
  const drawRels = drawingPath.replace(/drawings\/(drawing[^/]+\.xml)$/i, "drawings/_rels/$1.rels");
  const rid = addRel(wb, drawRels, `${R}/chart`, `../charts/chart${n}.xml`);
  // Append the anchor to the drawing XML (before </xdr:wsDr>).
  const xml = new TextDecoder().decode(wb.files[drawingPath]);
  const anchor = anchorXml(model, rid, ++frameSeq);
  wb.files[drawingPath] = new TextEncoder().encode(xml.replace(/<\/xdr:wsDr>\s*$/, `${anchor}</xdr:wsDr>`));
  model.original = { partPath: chartPath, drawingPath };
}

function rewriteExisting(wb: Workbook, model: ChartModel): void {
  const { partPath, drawingPath } = model.original!;
  if (partPath && wb.files[partPath]) wb.files[partPath] = new TextEncoder().encode(chartXml(model, wb));
  if (!drawingPath || !wb.files[drawingPath]) return;
  const doc = parseXmlOpt(wb.files[drawingPath]);
  if (!doc) return;
  const drawRels = drawingPath.replace(/drawings\/(drawing[^/]+\.xml)$/i, "drawings/_rels/$1.rels");
  const relsDoc = wb.files[drawRels] ? parseXmlOpt(wb.files[drawRels]) : undefined;
  const ridOfPart = new Map<string, string>(); // rId -> resolved part path
  if (relsDoc) for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
    const parts: string[] = [];
    for (const seg of `xl/drawings/${r.getAttribute("Target")}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    ridOfPart.set(r.getAttribute("Id") ?? "", parts.join("/"));
  }
  // Find the anchor whose graphicFrame chart r:id resolves to this chart's part; rewrite its from/to.
  const anchors = Array.from(doc.documentElement.children).filter((e) => /Anchor$/.test(e.localName));
  const a = model.anchor;
  for (const anchorEl of anchors) {
    const chartEl = Array.from(anchorEl.getElementsByTagName("*")).find((e) => e.localName === "chart");
    const rid = chartEl?.getAttributeNS(R, "id") ?? chartEl?.getAttribute("r:id");
    if (!rid || ridOfPart.get(rid) !== partPath) continue;
    const setPt = (tag: string, col: number, colOff: number, row: number, rowOff: number): void => {
      const p = Array.from(anchorEl.children).find((c) => c.localName === tag);
      if (!p) return;
      const set = (local: string, v: number): void => { const e = Array.from(p.children).find((x) => x.localName === local); if (e) e.textContent = String(v); };
      set("col", col - 1); set("colOff", pxToEmu(colOff)); set("row", row - 1); set("rowOff", pxToEmu(rowOff));
    };
    setPt("from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff);
    setPt("to", a.toCol, a.toColOff, a.toRow, a.toRowOff);
    break;
  }
  wb.files[drawingPath] = serializeXml(doc);
}

/** Persist all dirty charts to the workbook's DrawingML parts. */
export function writeXlsxCharts(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    if (!sheet.path) continue;
    for (const chart of sheet.charts ?? []) {
      if (!chart.dirty) continue;
      if (chart.original?.partPath) rewriteExisting(wb, chart);
      else createNew(wb, sheet, chart);
      chart.dirty = false;
    }
  }
}
