import { formatNumber, type Sheet, type Workbook } from "../model";
import { CHART_PALETTE, type ChartDataLabels, type ChartModel, type ChartTextStyle } from "../chart-model";
import { resolveNumbers, resolveLabels, seriesName } from "../chart-data";
import { backgroundPlugin, bar3DPlugin, errorBarsPlugin, line3DPlugin, multiLevelAxisPlugin, ofPiePlugin, pie3DPlugin, stockPlugin, surfacePlugin, trendlinePlugin } from "./chart-plugins";
import { registerDateAdapter } from "./date-adapter";

// DrawingML / ODF marker symbols -> Chart.js point styles.
const MARKER_STYLE: Record<string, string | false> = { circle: "circle", square: "rect", diamond: "rectRot", triangle: "triangle", star: "star", x: "crossRot", plus: "cross", dash: "line", dot: "circle", none: false };
// a:prstDash presets -> Chart.js borderDash arrays.
const DASH_MAP: Record<string, number[]> = { solid: [], dash: [6, 4], dot: [2, 3], dashDot: [6, 3, 2, 3], lgDash: [12, 4], lgDashDot: [12, 4, 2, 4], lgDashDotDot: [12, 4, 2, 4, 2, 4], sysDash: [4, 3], sysDot: [1, 3], sysDashDot: [4, 3, 1, 3], sysDashDotDot: [4, 3, 1, 3, 1, 3] };
const PT_PX = 4 / 3;

// The chart layer: floats a Chart.js canvas over the grid for every chart on the active sheet,
// anchored to cells and kept glued while scrolling. The layer sits over the data area (below the
// column header, right of the row numbers) and translates its content by the scroll offset, so
// charts scroll with the cells and clip under the frozen header. Chart.js is lazy-imported the
// first time a chart is drawn, so it never weighs on chart-free workbooks.

export interface ChartGeom { xOfCol: (c: number) => number; yOfRow: (r: number) => number; colAt: (px: number) => number; rowAt: (px: number) => number; rnW: number; headerH: number }

export interface ChartLayerDeps {
  wrap: HTMLElement;
  gridScroll: HTMLElement;
  getSheet: () => Sheet | undefined;
  getWorkbook: () => Workbook;
  geom: () => ChartGeom;
  onSelect?: (chart: ChartModel | null) => void;
  /** After a move/resize: the model's anchor was updated and dirty set; the host marks + persists. */
  onEdit?: (chart: ChartModel) => void;
}

export type ChartCtor = (new (ctx: CanvasRenderingContext2D, cfg: unknown) => { update(): void; destroy(): void; data: unknown; options: unknown }) & { register(p: unknown): void };
let ChartJs: ChartCtor | null = null;
let loading: Promise<void> | null = null;
/** Lazy-load Chart.js and the data-labels plugin; resolves to the Chart constructor. */
export async function loadChartJs(): Promise<ChartCtor> {
  if (ChartJs) return ChartJs;
  loading ??= Promise.all([import("chart.js/auto"), import("chartjs-plugin-datalabels"), import("chart.js")]).then(([m, dl, core]) => {
    ChartJs = (m.default ?? (m as { Chart: ChartCtor }).Chart) as ChartCtor;
    ChartJs.register((dl.default ?? dl) as unknown); // registered globally; per-chart display is opt-in
    ChartJs.register(trendlinePlugin as unknown); // no-ops unless a dataset carries a trendline
    ChartJs.register(errorBarsPlugin as unknown); // no-ops unless a dataset carries error bars
    ChartJs.register(stockPlugin as unknown); // no-ops unless datasets carry stock roles
    ChartJs.register(surfacePlugin as unknown); // no-ops unless datasets carry surface rows
    ChartJs.register(ofPiePlugin as unknown); // no-ops unless a dataset carries an ofPie config
    ChartJs.register(backgroundPlugin as unknown); // no-ops unless its plugin options carry a fill
    ChartJs.register(bar3DPlugin as unknown); // no-ops unless a dataset carries threeDBar
    ChartJs.register(pie3DPlugin as unknown); // no-ops unless a dataset carries threeDPie
    ChartJs.register(line3DPlugin as unknown); // no-ops unless a dataset carries threeDLine
    ChartJs.register(multiLevelAxisPlugin as unknown); // no-ops unless its plugin options carry levels
    // Chart.js reads the date adapter from the module-level _adapters singleton, not off the ctor.
    registerDateAdapter((core as { _adapters?: { _date?: { override(a: unknown): void } } })._adapters); // enables the time scale for date axes
  });
  await loading;
  return ChartJs!;
}
const ensureChartJs = async (): Promise<void> => { await loadChartJs(); };

const numbers = resolveNumbers;
const labels = resolveLabels;
const nameOf = seriesName;

/** Build a Chart.js config from the model + live data (shared by the overlay and the preview). */
export function chartConfig(model: ChartModel, wb: Workbook): unknown {
  return toConfig(model, wb);
}
const mapKind = (k: ChartModel["kind"]): string => (k === "column" || k === "bar" || k === "surface" ? "bar" : k === "area" || k === "stock" ? "line" : k);
function toConfig(model: ChartModel, wb: Workbook): unknown {
  const type = mapKind(model.kind);
  const cats = labels(wb, model.categories);
  // The chart's own palette (colors1.xml / theme accents) when present, else the built-in default.
  const PAL = model.palette?.length ? model.palette : CHART_PALETTE;
  const palette = (i: number, c?: string): string => c ?? PAL[i % PAL.length];
  const pieLike = model.kind === "pie" || model.kind === "doughnut";
  const hasSecondary = model.series.some((s) => s.secondaryAxis);
  // Date (category) axis for line/area: plot values against a timestamp on a linear scale with
  // date-formatted ticks, so points are spaced by their date. No date adapter needed.
  const xDate = model.axes?.x?.date === true && (model.kind === "line" || model.kind === "area");
  const excelToMs = (serial: number): number => Math.round((serial - 25569) * 86400000);
  const toTs = (c: unknown): number | null => {
    if (typeof c === "number") return excelToMs(c);
    if (typeof c === "string") {
      const t = c.trim();
      // A bare number is an Excel date serial; check this BEFORE Date.parse, which would read
      // "45292" as the year 45292 rather than a serial.
      if (/^-?\d+(?:\.\d+)?$/.test(t)) return excelToMs(Number(t));
      const ms = Date.parse(t); if (!isNaN(ms)) return ms;
    }
    return null;
  };
  const tss = xDate ? cats.map(toTs) : [];
  const blank = (v: number | null): number | null => (v == null ? (model.blanksAs === "gap" ? null : 0) : v); // "span" also keeps 0/null; spanGaps below connects
  // For a 100% stacked (percentStacked) chart, normalise each category to its total.
  // Category-like values (everything except scatter/bubble, which use xVal/yVal). Pie/doughnut ARE
  // category-like and need these to build their slice data.
  const rawBySeries = model.kind === "scatter" || model.kind === "bubble" ? [] : model.series.map((s) => numbers(wb, s.values));
  const totals: number[] = [];
  if (model.percent) for (let j = 0; j < Math.max(0, ...rawBySeries.map((a) => a.length)); j++) totals[j] = rawBySeries.reduce((t, a) => t + (a[j] ?? 0), 0);
  const perPoint = (s: typeof model.series[number], j: number): string => s.pointColors?.[j] ?? PAL[j % PAL.length];
  // Effective data-label spec for a series: its own, else the chart's, else the simple toggle.
  const labelSpecOf = (s: typeof model.series[number]): ChartDataLabels | undefined => s.labels ?? model.labels ?? (model.dataLabels ? { value: true } : undefined);
  const yFmtG = model.axes?.y?.numFmt;
  const fmtVal = (n: unknown): unknown => (yFmtG != null ? (formatNumber(yFmtG, String(n)) ?? n) : n);
  const posMap = (position: string | undefined): { anchor: string; align: string } => {
    switch (position) {
      case "ctr": case "bestFit": return { anchor: "center", align: "center" };
      case "inEnd": return { anchor: "end", align: "start" };
      case "inBase": return { anchor: "start", align: "end" };
      case "outEnd": return { anchor: "end", align: "end" };
      case "t": return { anchor: "end", align: "top" };
      case "b": return { anchor: "start", align: "bottom" };
      case "l": return { anchor: "start", align: "left" };
      case "r": return { anchor: "end", align: "right" };
      default: return pieLike ? { anchor: "center", align: "center" } : { anchor: "end", align: "top" };
    }
  };
  const yOf = (v: unknown): number => (typeof v === "object" && v ? ((v as { y?: number }).y ?? 0) : ((v as number) ?? 0));
  const makeFormatter = (spec: ChartDataLabels) => (val: unknown, ctx: { chart: { data: { labels?: unknown[] } }; dataset: { label?: string; data: unknown[] }; dataIndex: number }): string => {
    const y = yOf(val);
    const parts: string[] = [];
    if (spec.category) parts.push(String(ctx.chart.data.labels?.[ctx.dataIndex] ?? ""));
    if (spec.seriesName) parts.push(String(ctx.dataset.label ?? ""));
    if (spec.value) parts.push(String(fmtVal(y)));
    if (spec.percent) { const tot = ctx.dataset.data.reduce((t: number, x) => t + yOf(x), 0); parts.push(tot ? `${(y / tot * 100).toFixed(1)}%` : ""); }
    if (!parts.length) parts.push(String(fmtVal(y)));
    return parts.filter((p) => p !== "").join(spec.category || spec.seriesName ? "\n" : " ");
  };
  const datasets = model.series.map((s, i) => {
    const base: Record<string, unknown> = { label: nameOf(wb, s.name), borderColor: palette(i, s.color), backgroundColor: palette(i, s.color) };
    if (model.kind === "scatter" || model.kind === "bubble") {
      const xs = numbers(wb, s.xValues); const ys = numbers(wb, s.values); const rs = s.sizes ? numbers(wb, s.sizes) : [];
      base.data = ys.map((y, j) => ({ x: xs[j] ?? j, y: y ?? 0, r: model.kind === "bubble" ? (rs[j] ?? 5) : undefined }));
      if (s.smooth) { base.tension = 0.4; base.showLine = true; }
      if (s.marker) { const st = MARKER_STYLE[s.marker.symbol ?? "circle"] ?? "circle"; base.pointStyle = st; base.pointRadius = s.marker.size ? s.marker.size / 2 : 4; }
    } else if (xDate) {
      base.data = rawBySeries[i].map((v, j) => ({ x: tss[j] ?? j, y: blank(v) }));
      base.parsing = false; // data is already {x,y}
      if (s.type && mapKind(s.type) !== type) base.type = mapKind(s.type);
      if (model.kind === "area") base.fill = true;
      if (s.smooth) base.tension = 0.4;
      if (s.marker) { const st = MARKER_STYLE[s.marker.symbol ?? "circle"] ?? "circle"; base.pointStyle = st; base.pointRadius = st === false || s.marker.symbol === "none" ? 0 : (s.marker.size ? s.marker.size / 2 : 3); }
    } else {
      base.data = rawBySeries[i].map((v, j) => (model.percent ? (totals[j] ? (v ?? 0) / totals[j] * 100 : 0) : blank(v)));
      const effKind = s.type ?? model.kind;
      if (effKind === "area") base.fill = true;
      if (s.type && mapKind(s.type) !== type) base.type = mapKind(s.type); // combo: per-series type override
      if (s.secondaryAxis) base.yAxisID = "y1";
      if (s.smooth) base.tension = 0.4;
      if (s.marker) { const st = MARKER_STYLE[s.marker.symbol ?? "circle"] ?? "circle"; base.pointStyle = st; base.pointRadius = st === false || s.marker.symbol === "none" ? 0 : (s.marker.size ? s.marker.size / 2 : 3); }
      // Per-point colours: pie/doughnut slices always; a series with its own dPt colours otherwise.
      if (pieLike) { base.backgroundColor = (base.data as number[]).map((_, j) => perPoint(s, j)); base.borderColor = "#fff"; }
      else if (s.pointColors) base.backgroundColor = (base.data as number[]).map((_, j) => perPoint(s, j));
      // Bar spacing from the file's gapWidth (approximate mapping to Chart.js categoryPercentage).
      if ((model.kind === "column" || model.kind === "bar") && model.gapWidth != null) base.categoryPercentage = Math.max(0.1, Math.min(1, 1 / (1 + model.gapWidth / 100)));
      // Pie/doughnut slice explosion -> per-arc pixel offset.
      if (pieLike && s.explosion) base.offset = s.explosion.map((e) => (e ? Math.round(e * 0.6) : 0));
    }
    // Per-dataset data labels (content + position); the global plugin default is off.
    const spec = labelSpecOf(s);
    if (spec) { const pm = posMap(spec.position); base.datalabels = { display: true, anchor: pm.anchor, align: pm.align, color: pieLike ? "#fff" : "#444", font: { size: 10 }, formatter: makeFormatter(spec) }; }
    else base.datalabels = { display: false };
    // Regression trendline (drawn by the trendline plugin).
    if (s.trendline) base.trendline = { ...s.trendline, color: s.trendline.color ?? palette(i, s.color) };
    // Error bars (drawn by the error-bars plugin).
    if (s.errorBars) base.errorBars = s.errorBars;
    // Line width / dash (line + scatter series).
    if (s.lineWidth != null) base.borderWidth = Math.max(1, Math.round(s.lineWidth * PT_PX));
    if (s.dash) base.borderDash = DASH_MAP[s.dash] ?? [];
    return base;
  });
  // Stock: the O/H/L/C series become invisible scale-carriers; the stock plugin draws the candles.
  const isStock = model.kind === "stock";
  if (isStock) {
    const roles = model.series.length >= 4 ? ["open", "high", "low", "close"] : model.series.length === 3 ? ["high", "low", "close"] : ["high", "low"];
    datasets.forEach((d, i) => { d.type = "line"; d.borderColor = "rgba(0,0,0,0)"; d.backgroundColor = "rgba(0,0,0,0)"; d.pointRadius = 0; d.showLine = false; d.stockRole = roles[i] ?? "close"; d.datalabels = { display: false }; });
  }
  // Surface: no 2D equivalent, so render a heatmap (series = rows, categories = columns, value = colour).
  const isSurface = model.kind === "surface";
  if (isSurface) datasets.forEach((d, i) => { d.backgroundColor = "rgba(0,0,0,0)"; d.borderColor = "rgba(0,0,0,0)"; d.surfaceRow = i; d.surfaceName = nameOf(wb, model.series[i].name); d.datalabels = { display: false }; });
  // Pseudo-3D (isometric) rendering for 3-D column and pie/doughnut charts: the flat datasets go
  // transparent (scale/geometry carriers) and a plugin draws the extruded look.
  const threeDBar = model.threeD === true && (model.kind === "column" || model.kind === "bar");
  const threeDPie = model.threeD === true && pieLike;
  const threeDLine = model.threeD === true && (model.kind === "line" || model.kind === "area");
  // Tiered category axis: draw the outer level(s) below the innermost labels (a plugin does it).
  const multiLevel = !!model.categoryLevels && model.categoryLevels.length > 1 && !xDate && !pieLike && !isStock && !isSurface && model.kind !== "scatter" && model.kind !== "bubble" && model.kind !== "radar";
  if (threeDBar) datasets.forEach((d, i) => { d.threeDBar = palette(i, model.series[i]?.color); d.threeDHoriz = model.kind === "bar"; d.backgroundColor = "rgba(0,0,0,0)"; d.borderColor = "rgba(0,0,0,0)"; d.datalabels = { display: false }; });
  if (threeDPie && datasets[0]) { const arr = Array.isArray(datasets[0].backgroundColor) ? (datasets[0].backgroundColor as string[]) : []; datasets[0].threeDPie = arr; datasets[0].backgroundColor = arr.map(() => "rgba(0,0,0,0)"); datasets[0].borderColor = "rgba(0,0,0,0)"; datasets[0].datalabels = { display: false }; }
  if (threeDLine) datasets.forEach((d, i) => { const s = model.series[i]; d.threeDLine = { colour: palette(i, s?.color), area: model.kind === "area" || s?.type === "area" }; d.backgroundColor = "rgba(0,0,0,0)"; d.borderColor = "rgba(0,0,0,0)"; d.datalabels = { display: false }; });
  // Pie-of-pie / bar-of-pie: aggregate the last splitCount slices into "Other"; the plugin draws
  // the breakout as a small secondary pie or bar.
  const isOfPie = pieLike && !!model.ofPie && model.series.length >= 1;
  let ofPieLabels: string[] | undefined;
  if (isOfPie) {
    const vals = numbers(wb, model.series[0].values).map((v) => v ?? 0);
    const k = Math.min(Math.max(1, model.ofPie!.splitCount ?? 2), Math.max(1, vals.length - 1));
    const primN = vals.length - k;
    const other = vals.slice(primN).reduce((t, v) => t + v, 0);
    const secondary = vals.slice(primN).map((v, j) => ({ label: String(cats[primN + j] ?? ""), value: v, color: PAL[(primN + 1 + j) % PAL.length] }));
    const ds = datasets[0];
    ds.data = [...vals.slice(0, primN), other];
    ds.backgroundColor = [...Array.from({ length: primN }, (_, j) => PAL[j % PAL.length]), "#9c9c9c"];
    ds.borderColor = "#fff";
    ds.ofPie = { type: model.ofPie!.type, secondary };
    datasets.length = 1;
    ofPieLabels = [...cats.slice(0, primN).map(String), "Other"];
  }
  const stackOpt = model.stacked || model.percent ? { stacked: true } : {};
  const yFmt = model.axes?.y?.numFmt;
  // A ChartTextStyle -> Chart.js { font, color } fragment.
  const cjFont = (s?: ChartTextStyle): { font?: object; color?: string } => {
    if (!s) return {};
    const font: Record<string, unknown> = {};
    if (s.size != null) font.size = s.size;
    if (s.bold) font.weight = "bold";
    if (s.italic) font.style = "italic";
    if (s.font) font.family = s.font;
    return { ...(Object.keys(font).length ? { font } : {}), ...(s.color ? { color: s.color } : {}) };
  };
  const ticksOpt = (fmt?: string, style?: ChartTextStyle): object => {
    const ticks = { ...(fmt ? { callback: (v: unknown) => formatNumber(fmt, String(v)) ?? v } : {}), ...cjFont(style) };
    return Object.keys(ticks).length ? { ticks } : {};
  };
  const axisOpts = (title?: string, min?: number, max?: number, fmt?: string, labelStyle?: ChartTextStyle, titleStyle?: ChartTextStyle): Record<string, unknown> => ({ ...(title ? { title: { display: true, text: title, ...cjFont(titleStyle) } } : {}), ...(min != null ? { min } : {}), ...(max != null ? { max } : {}), ...ticksOpt(fmt, labelStyle) });
  const xa = xDate
    ? { type: "time", ticks: { ...cjFont(model.axes?.x?.labelStyle) }, ...(model.axes?.x?.title ? { title: { display: true, text: model.axes.x.title, ...cjFont(model.axes?.x?.titleStyle) } } : {}) }
    : axisOpts(model.axes?.x?.title, undefined, undefined, model.axes?.x?.numFmt, model.axes?.x?.labelStyle, model.axes?.x?.titleStyle);
  const ya = axisOpts(model.axes?.y?.title, model.axes?.y?.min, model.percent ? 100 : model.axes?.y?.max, yFmt, model.axes?.y?.labelStyle, model.axes?.y?.titleStyle);
  const catLinear = model.kind === "column" || model.kind === "line" || model.kind === "area" || model.kind === "stock" || model.kind === "surface";
  const scales: Record<string, unknown> | undefined = model.kind === "bar" ? { x: { ...stackOpt, beginAtZero: true, ...ya }, y: { ...stackOpt, ...xa } }
    : catLinear ? { x: { ...stackOpt, ...xa }, y: { ...stackOpt, beginAtZero: model.kind !== "line" && model.kind !== "stock", ...ya } }
    : undefined;
  if (scales && hasSecondary) scales.y1 = { position: "right", grid: { drawOnChartArea: false } };
  if (isSurface && scales) scales.y = { display: false }; // the heatmap draws its own rows
  return {
    type,
    data: { labels: isOfPie ? ofPieLabels : (cats.length && !xDate ? cats : undefined), datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      // Reserve space for the of-pie secondary plot (right) and the tiered-axis rows (bottom).
      ...((isOfPie || multiLevel) ? { layout: { padding: { ...(isOfPie ? { right: 130 } : {}), ...(multiLevel ? { bottom: (model.categoryLevels!.length - 1) * 20 } : {}) } } } : {}),
      indexAxis: model.kind === "bar" ? "y" : "x",
      spanGaps: model.blanksAs === "span",
      ...(model.kind === "doughnut" && model.holeSize != null ? { cutout: `${model.holeSize}%` } : {}),
      ...(pieLike && model.rotation != null ? { rotation: model.rotation } : {}),
      plugins: {
        legend: {
          display: isStock || isSurface ? false : (model.legend?.show ?? true),
          position: model.legend?.pos ?? "top",
          // Hide deleted legend entries (by slice index for pie-like, else by dataset index).
          labels: {
            ...(model.legend?.deleted?.length ? { filter: (item: { index: number; datasetIndex: number }) => !model.legend!.deleted!.includes(pieLike ? item.index : item.datasetIndex) } : {}),
            ...cjFont(model.legendStyle),
            // 3-D datasets are transparent; source the legend swatch colours from the 3D tags.
            ...(threeDBar ? { generateLabels: (ch: { data: { datasets: { label?: string; threeDBar?: string }[] } }) => ch.data.datasets.map((d, i) => ({ text: d.label ?? "", fillStyle: d.threeDBar, strokeStyle: d.threeDBar, lineWidth: 0, datasetIndex: i })) } : {}),
            ...(threeDLine ? { generateLabels: (ch: { data: { datasets: { label?: string; threeDLine?: { colour: string } }[] } }) => ch.data.datasets.map((d, i) => ({ text: d.label ?? "", fillStyle: d.threeDLine?.colour, strokeStyle: d.threeDLine?.colour, lineWidth: 0, datasetIndex: i })) } : {}),
            ...(threeDPie ? { generateLabels: (ch: { data: { labels?: unknown[]; datasets: { threeDPie?: string[] }[] } }) => (ch.data.labels ?? []).map((lab, i) => ({ text: String(lab), fillStyle: ch.data.datasets[0]?.threeDPie?.[i], strokeStyle: "#fff", lineWidth: 1, index: i })) } : {}),
          },
        },
        title: { display: !!model.title, text: model.title, ...cjFont(model.titleStyle) },
        // Registered globally but off by default; each dataset opts in via its own datalabels config.
        datalabels: { display: false },
        // Background fills (drawn by the background plugin); absent -> plugin no-ops.
        ...(model.plotFill || model.areaFill ? { sheeteditBg: { plot: model.plotFill, area: model.areaFill } } : {}),
        // Outer levels of a tiered category axis (drawn by the multi-level plugin).
        ...(multiLevel ? { sheeteditMultiLevel: model.categoryLevels } : {}),
      },
      ...(scales ? { scales } : {}),
    },
  };
}

export function setupChartLayer(deps: ChartLayerDeps): { refresh(): void; update(): void; select(id: string | null): void; boxRect(id: string): DOMRect | null; teardown(): void } {
  const { wrap, gridScroll } = deps;
  const layer = document.createElement("div");
  layer.className = "sheetedit-chartlayer";
  const inner = document.createElement("div");
  inner.className = "sheetedit-chartlayer-inner";
  layer.appendChild(inner);
  wrap.appendChild(layer);

  const instances = new Map<string, { box: HTMLElement; canvas: HTMLCanvasElement; chart: InstanceType<ChartCtor> | null; kind?: string }>();
  let selectedId: string | null = null;

  const setSelected = (id: string | null): void => {
    selectedId = id;
    for (const [k, inst] of instances) inst.box.classList.toggle("sel", k === id);
    deps.onSelect?.(id ? (deps.getSheet()?.charts ?? []).find((c) => c.id === id) ?? null : null);
  };

  // Commit a dragged/resized pixel rect (content coords) back to the model's cell anchor.
  const rectToAnchor = (model: ChartModel, x: number, y: number, w: number, h: number): void => {
    const g = deps.geom();
    const set = (px: number, at: (p: number) => number, of: (i: number) => number): [number, number] => { const i = Math.max(1, at(px)); return [i, Math.max(0, px - of(i))]; };
    const [fc, fco] = set(x, g.colAt, g.xOfCol);
    const [fr, fro] = set(y, g.rowAt, g.yOfRow);
    const [tc, tco] = set(x + w, g.colAt, g.xOfCol);
    const [tr, tro] = set(y + h, g.rowAt, g.yOfRow);
    model.anchor = { fromCol: fc, fromRow: fr, fromColOff: fco, fromRowOff: fro, toCol: tc, toRow: tr, toColOff: tco, toRowOff: tro };
    model.dirty = true;
  };

  const positionLayer = (): void => {
    const g = deps.geom();
    const gr = gridScroll.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    layer.style.left = `${gr.left - wr.left + g.rnW}px`;
    layer.style.top = `${gr.top - wr.top + g.headerH}px`;
    layer.style.width = `${Math.max(0, gr.width - g.rnW)}px`;
    layer.style.height = `${Math.max(0, gr.height - g.headerH)}px`;
  };
  const syncScroll = (): void => { inner.style.transform = `translate(${-gridScroll.scrollLeft}px, ${-gridScroll.scrollTop}px)`; };

  const positionBox = (box: HTMLElement, model: ChartModel): void => {
    const g = deps.geom();
    const a = model.anchor;
    const x = g.xOfCol(a.fromCol) + a.fromColOff;
    const y = g.yOfRow(a.fromRow) + a.fromRowOff;
    const x2 = g.xOfCol(a.toCol) + a.toColOff;
    const y2 = g.yOfRow(a.toRow) + a.toRowOff;
    box.style.left = `${x}px`;
    box.style.top = `${y}px`;
    box.style.width = `${Math.max(60, x2 - x)}px`;
    box.style.height = `${Math.max(40, y2 - y)}px`;
  };

  // Drag the box to move, or its corner handle to resize (mouse or touch); commit on release.
  function attachDrag(box: HTMLElement, handle: HTMLElement, model: ChartModel): void {
    const start = (e: PointerEvent, mode: "move" | "resize"): void => {
      e.preventDefault();
      e.stopPropagation();
      setSelected(model.id);
      const sx = e.clientX;
      const sy = e.clientY;
      const x0 = parseFloat(box.style.left) || 0;
      const y0 = parseFloat(box.style.top) || 0;
      const w0 = box.offsetWidth;
      const h0 = box.offsetHeight;
      const onMove = (ev: PointerEvent): void => {
        const dx = ev.clientX - sx;
        const dy = ev.clientY - sy;
        if (mode === "move") { box.style.left = `${Math.max(0, x0 + dx)}px`; box.style.top = `${Math.max(0, y0 + dy)}px`; }
        else { box.style.width = `${Math.max(80, w0 + dx)}px`; box.style.height = `${Math.max(60, h0 + dy)}px`; }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        rectToAnchor(model, parseFloat(box.style.left) || 0, parseFloat(box.style.top) || 0, box.offsetWidth, box.offsetHeight);
        deps.onEdit?.(model);
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    box.addEventListener("pointerdown", (e) => { if (e.target !== handle) start(e, "move"); });
    handle.addEventListener("pointerdown", (e) => start(e, "resize"));
  }

  let drawSeq = 0;
  const refresh = (): void => {
    const sheet = deps.getSheet();
    const charts = sheet?.charts ?? [];
    positionLayer();
    syncScroll();
    // Drop instances for charts no longer present.
    for (const [id, inst] of instances) if (!charts.some((c) => c.id === id)) { inst.chart?.destroy(); inst.box.remove(); instances.delete(id); }
    if (!charts.length) return;
    const wb = deps.getWorkbook();
    const seq = ++drawSeq;
    void ensureChartJs().then(() => {
      if (seq !== drawSeq || !ChartJs) return; // a newer refresh superseded this one
      for (const model of charts) {
        let inst = instances.get(model.id);
        if (!inst) {
          const box = document.createElement("div");
          box.className = "sheetedit-chartbox";
          const canvas = document.createElement("canvas");
          box.appendChild(canvas);
          const handle = document.createElement("div");
          handle.className = "sheetedit-chart-resize";
          box.appendChild(handle);
          attachDrag(box, handle, model);
          inner.appendChild(box);
          inst = { box, canvas, chart: null };
          instances.set(model.id, inst);
        }
        positionBox(inst.box, model);
        const cfg = toConfig(model, wb);
        // Chart.js can't change its type via update(); recreate when the kind changed.
        if (inst.chart && inst.kind !== model.kind) { inst.chart.destroy(); inst.chart = null; }
        if (inst.chart) { const c = inst.chart as { data: unknown; options: unknown; update(): void }; const nc = cfg as { data: unknown; options: unknown }; c.data = nc.data; c.options = nc.options; c.update(); }
        else inst.chart = new ChartJs(inst.canvas.getContext("2d")!, cfg);
        inst.kind = model.kind;
      }
    });
  };

  // Data-only refresh (a cell edit changed a value a chart reads): re-resolve and update, no reposition.
  const update = (): void => {
    const wb = deps.getWorkbook();
    for (const model of deps.getSheet()?.charts ?? []) {
      const inst = instances.get(model.id);
      if (!inst?.chart) continue;
      const cfg = toConfig(model, wb) as { data: unknown; options: unknown };
      const c = inst.chart as { data: unknown; options: unknown; update(): void };
      c.data = cfg.data; c.options = cfg.options; c.update();
    }
  };

  gridScroll.addEventListener("scroll", syncScroll, { passive: true });
  window.addEventListener("resize", refresh);
  // Tap/click on empty grid deselects.
  gridScroll.addEventListener("pointerdown", () => { if (selectedId) setSelected(null); });

  /** Screen rect of a chart's box (for anchoring an external edit toolbar). */
  const boxRect = (id: string): DOMRect | null => instances.get(id)?.box.getBoundingClientRect() ?? null;

  return {
    refresh,
    update,
    select: setSelected,
    boxRect,
    teardown(): void {
      gridScroll.removeEventListener("scroll", syncScroll);
      window.removeEventListener("resize", refresh);
      for (const inst of instances.values()) inst.chart?.destroy();
      layer.remove();
    },
  };
}
