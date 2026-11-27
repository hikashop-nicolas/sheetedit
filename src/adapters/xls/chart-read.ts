// BIFF8 chart substreams ([MS-XLS] 2.1.7.20 / 2.4 chart records) into sheetedit's chart model.
// The records nest through BEGIN / END; the series' data comes from their BRAI formulas, with the
// SERIESDATA cells after the chart as the cache.
import { type BiffRecord, Cursor, records } from "./biff";
import type { ChartAxis, ChartDataLabels, ChartKind, ChartModel, ChartRef, ChartSeries, ChartTextStyle } from "../../core/chart-model";
import type { XlsFont } from "./types";

export interface CRec { type: number; data: Uint8Array; conts: Uint8Array[]; kids?: CRec[] }

/** What the chart reader needs from the workbook: formulas, fonts, number formats, colours. */
export interface ChartEnv {
  formula: (rgce: Uint8Array) => string | undefined;
  font: (ifnt: number) => XlsFont | undefined;
  format: (ifmt: number) => string | undefined;
  color: (icv: number) => string | undefined;
}

/** A chart read from BIFF, before it is anchored: the model plus the chart's own size in points. */
export interface XlsChart { model: Omit<ChartModel, "id" | "anchor">; width: number; height: number; autoTitle?: string }

/** The records of the chart substream starting at `pos` (its BOF), up to its EOF. */
export function chartRecords(stream: Uint8Array, pos: number): BiffRecord[] {
  const out: BiffRecord[] = [];
  let first = true;
  for (const r of records(stream, pos)) {
    if (first) { first = false; continue; }
    if (r.type === 0x000a) break;
    if (r.type === 0x0809) break; // a damaged substream with no EOF
    out.push(r);
    if (out.length > 1_000_000) break;
  }
  return out;
}

/** Nest the records: a record followed by BEGIN owns everything up to the matching END. */
export function chartTree(recs: BiffRecord[]): CRec[] {
  const top: CRec[] = [];
  const stack: CRec[][] = [top];
  for (const r of recs) {
    const cur = stack[stack.length - 1]!;
    if (r.type === 0x1033) {
      const owner = cur[cur.length - 1];
      if (owner) { owner.kids = []; stack.push(owner.kids); } else stack.push([]);
    } else if (r.type === 0x1034) {
      if (stack.length > 1) stack.pop();
    } else if (r.type !== 0x0852 && r.type !== 0x0853 && r.type !== 0x0850) cur.push({ type: r.type, data: r.data, conts: r.conts });
  }
  return top;
}

const kidsOf = (r: CRec | undefined, type: number): CRec[] => (r?.kids ?? []).filter((k) => k.type === type);
const kid = (r: CRec | undefined, type: number): CRec | undefined => r?.kids?.find((k) => k.type === type);
const u16 = (d: Uint8Array, o: number): number => (d[o] ?? 0) | ((d[o + 1] ?? 0) << 8);
const i16 = (d: Uint8Array, o: number): number => (u16(d, o) << 16) >> 16;
const u32 = (d: Uint8Array, o: number): number => (u16(d, o) | (u16(d, o + 2) << 16)) >>> 0;
const f64 = (d: Uint8Array, o: number): number => (o + 8 <= d.length ? new DataView(d.buffer, d.byteOffset + o, 8).getFloat64(0, true) : NaN);
const fixed = (d: Uint8Array, o: number): number => i16(d, o + 2) + u16(d, o) / 65536;
const rgbCss = (d: Uint8Array, o: number): string => "#" + [d[o] ?? 0, d[o + 1] ?? 0, d[o + 2] ?? 0].map((v) => v.toString(16).padStart(2, "0")).join("");

/** SeriesText: the literal text of a title or series name. */
function seriesText(r: CRec | undefined): string | undefined {
  if (!r || r.data.length < 3) return undefined;
  const c = new Cursor({ data: r.data, conts: r.conts });
  c.u16();
  const cch = c.u8();
  const flags = c.u8();
  return c.chars8(Math.min(cch, c.left), (flags & 1) === 1);
}

/** A BRAI: its kind, and a formula (rt 2) or literal (rt 1). */
interface Brai { id: number; rt: number; formula?: string; ifmt: number; sourceLinked: boolean }
function brai(r: CRec, env: ChartEnv): Brai {
  const d = r.data;
  const id = d[0] ?? 0, rt = d[1] ?? 0, flags = u16(d, 2), ifmt = u16(d, 4), cce = u16(d, 6);
  const out: Brai = { id, rt, ifmt, sourceLinked: !(flags & 1) };
  if (rt === 2 && cce) out.formula = env.formula(d.subarray(8, 8 + cce))?.replace(/^\((.*)\)$/, "$1");
  return out;
}

function textStyle(r: CRec | undefined, env: ChartEnv, text?: CRec): ChartTextStyle | undefined {
  const fx = kid(r, 0x1026);
  const f = fx ? env.font(u16(fx.data, 0)) : undefined;
  const out: ChartTextStyle = {};
  if (f) {
    out.size = f.height / 20;
    if (f.weight >= 700) out.bold = true;
    if (f.italic) out.italic = true;
    if (f.name) out.font = f.name;
  }
  // The TEXT record's own colour, unless it is automatic.
  if (text && !(u16(text.data, 24) & 1)) out.color = env.color(u16(text.data, 26)) ?? rgbCss(text.data, 4);
  return Object.keys(out).length ? out : undefined;
}

/** The text a TEXT node shows (its SeriesText), if any. */
const labelText = (t: CRec): string | undefined => seriesText(kid(t, 0x100d));
const linkOf = (t: CRec): { obj: number; v1: number; v2: number } | undefined => {
  const l = kid(t, 0x1027);
  return l ? { obj: u16(l.data, 0), v1: u16(l.data, 2), v2: u16(l.data, 4) } : undefined;
};

/** The labels a TEXT / ATTACHEDLABEL turns on. */
function labelsOf(flags: number, attached: boolean): ChartDataLabels | undefined {
  const out: ChartDataLabels = {};
  if (attached) {
    if (flags & 0x01) out.value = true;
    if (flags & 0x02) out.percent = true;
    if (flags & 0x04) { out.percent = true; out.category = true; }
    if (flags & 0x10) out.category = true;
    if (flags & 0x40) out.seriesName = true;
  } else {
    if (flags & 0x04) out.value = true;
    if (flags & 0x1000) out.percent = true;
    if (flags & 0x0800) { out.percent = true; out.category = true; }
    if (flags & 0x4000) out.category = true;
    if (flags & 0x02) out.legendKey = true;
  }
  return Object.keys(out).length ? out : undefined;
}

const DLP = ["", "outEnd", "inEnd", "ctr", "inBase", "", "", "", "", "bestFit", "l", "r", "t", "b"];
const MARKERS = ["none", "square", "diamond", "triangle", "x", "star", "dash", "dot", "circle", "plus"];
const DASHES = [undefined, "dash", "sysDot", "dashDot", "lgDashDotDot"];

interface Group { kind: ChartKind; stacked?: boolean; percent?: boolean; secondary: boolean; icrt: number; rec: CRec; crt: CRec; threeD: boolean; hiLo: boolean }

function groupOf(crt: CRec, secondary: boolean, pos: number): Group | null {
  const icrt = crt.data.length >= 20 ? u16(crt.data, 18) : pos;
  const threeD = !!kid(crt, 0x103a);
  const hiLo = kidsOf(crt, 0x101c).some((l) => u16(l.data, 0) === 1);
  for (const r of crt.kids ?? []) {
    const d = r.data;
    const base = { secondary, icrt, rec: r, crt, threeD, hiLo };
    switch (r.type) {
      case 0x1017: { const f = u16(d, 4); return { ...base, kind: f & 1 ? "bar" : "column", stacked: !!(f & 2) || undefined, percent: !!(f & 4) || undefined }; }
      case 0x1018: { const f = u16(d, 0); return { ...base, kind: hiLo ? "stock" : "line", stacked: !!(f & 1) || undefined, percent: !!(f & 2) || undefined }; }
      case 0x101a: { const f = u16(d, 0); return { ...base, kind: "area", stacked: !!(f & 1) || undefined, percent: !!(f & 2) || undefined }; }
      case 0x1019: return { ...base, kind: u16(d, 2) ? "doughnut" : "pie" };
      case 0x1061: return { ...base, kind: "pie" };
      case 0x101b: return { ...base, kind: u16(d, 4) & 1 ? "bubble" : "scatter" };
      case 0x103e: case 0x1040: return { ...base, kind: "radar" };
      case 0x103f: return { ...base, kind: "surface" };
    }
  }
  return null;
}

/** The cached values of each series from SERIESDATA: by SIIndex kind, then series, then point. */
function seriesData(top: CRec[]): Map<number, (string | number | null)[][]> {
  const out = new Map<number, (string | number | null)[][]>();
  let cur: (string | number | null)[][] | null = null;
  for (const r of top) {
    const d = r.data;
    if (r.type === 0x1065) { const k = u16(d, 0); cur = out.get(k) ?? []; out.set(k, cur); continue; }
    if (!cur) continue;
    let v: string | number | null | undefined;
    if (r.type === 0x0203) v = f64(d, 6);
    else if (r.type === 0x0204) { const c = new Cursor({ data: d, conts: r.conts }); c.skip(6); const n = c.u16(); const fl = c.u8(); v = c.chars8(Math.min(n, c.left), (fl & 1) === 1); }
    else if (r.type === 0x0201) v = null;
    else if (r.type === 0x0205) v = d[7] ? null : d[6] ?? 0;
    else continue;
    const row = u16(d, 0), col = u16(d, 2);
    if (row > 32000 || col > 255) continue;
    const list = (cur[col] ??= []);
    while (list.length < row) list.push(null);
    list[row] = v;
  }
  return out;
}

/** Read one chart substream into the chart model. Returns null when it holds no plot. */
export function readChart(recs: BiffRecord[], env: ChartEnv): XlsChart | null {
  const top = chartTree(recs);
  const chart = top.find((r) => r.type === 0x1002);
  if (!chart?.kids) return null;
  const width = fixed(chart.data, 8), height = fixed(chart.data, 12);
  const cache = seriesData(top);
  // Chart groups (CRT) per axis group, primary first.
  const groups: Group[] = [];
  const axisParents = kidsOf(chart, 0x1041);
  const axes: { x?: ChartAxis; y?: ChartAxis } = {};
  let plotFill: string | undefined;
  let legend: ChartModel["legend"] = { show: false, pos: "right" };
  let legendStyle: ChartTextStyle | undefined;
  for (const ap of axisParents) {
    const secondary = u16(ap.data, 0) === 1;
    for (const crt of kidsOf(ap, 0x1014)) {
      const g = groupOf(crt, secondary, groups.length);
      if (g) groups.push(g);
      const lg = kid(crt, 0x1015);
      if (lg) {
        const w = lg.data[16] ?? 3;
        legend = { show: true, pos: w === 0 ? "bottom" : w === 2 ? "top" : w === 4 ? "left" : "right" };
        const t = kid(lg, 0x1025);
        legendStyle = textStyle(t, env, t);
      }
    }
    if (secondary) continue;
    // Axes: a category (or X value) axis, then a value axis.
    const axisNodes = kidsOf(ap, 0x101d);
    const isXY = groups.some((g) => g.kind === "scatter" || g.kind === "bubble");
    for (const ax of axisNodes) {
      const wType = u16(ax.data, 0);
      if (wType === 2) continue;
      const a: ChartAxis = {};
      const vr = kid(ax, 0x101f);
      if (vr) {
        const fl = u16(vr.data, 40);
        if (!(fl & 1)) a.min = f64(vr.data, 0);
        if (!(fl & 2)) a.max = f64(vr.data, 8);
      }
      const ifmt = kid(ax, 0x104e);
      if (ifmt) { const code = env.format(u16(ifmt.data, 0)); if (code && code !== "General") a.numFmt = code; }
      const ext = kid(ax, 0x1062);
      if (ext && (u16(ext.data, 16) & 0x10) && !(u16(ext.data, 16) & 0x80)) a.date = true;
      const ls = textStyle(ax, env);
      if (ls) a.labelStyle = ls;
      const which = wType === 0 || (isXY && !axes.x && wType === 1 && axisNodes.filter((n) => u16(n.data, 0) === 1).length > 1) ? "x" : "y";
      if (!axes[which]) axes[which] = a;
    }
    // Axis titles: TEXT nodes linked to the value (2) or category (3) axis.
    for (const t of kidsOf(ap, 0x1025)) {
      const l = linkOf(t);
      const txt = labelText(t);
      if (!l || !txt) continue;
      const which = l.obj === 3 ? "x" : l.obj === 2 ? "y" : undefined;
      if (!which) continue;
      const a = (axes[which] ??= {});
      a.title = txt;
      const st = textStyle(t, env, t);
      if (st) a.titleStyle = st;
    }
    const pa = (ap.kids ?? []).findIndex((k) => k.type === 0x1035);
    const frame = pa >= 0 ? ap.kids![pa + 1] : undefined;
    if (frame?.type === 0x1032) plotFill = areaColor(kid(frame, 0x100a));
  }
  if (!groups.length) return null;
  groups.sort((a, b) => Number(a.secondary) - Number(b.secondary) || a.icrt - b.icrt);
  const base = groups[0]!;
  // Series, each tied to its chart group by SERTOCRT.
  const series: ChartSeries[] = [];
  let categories: ChartRef | undefined;
  const allSeries = kidsOf(chart, 0x1003);
  const valCache = cache.get(1) ?? [], catCache = cache.get(2) ?? [], bubCache = cache.get(3) ?? [];
  allSeries.forEach((sr, si) => {
    if (kid(sr, 0x104a)) return; // a trendline / error bar series belongs to its parent
    const brais = kidsOf(sr, 0x1051).map((b) => brai(b, env));
    const at = (id: number): Brai | undefined => brais.find((b) => b.id === id);
    const stc = kid(sr, 0x1045);
    const g = (stc ? groups.find((x) => x.icrt === u16(stc.data, 0)) : undefined) ?? base;
    const nameB = at(0);
    const after = (sr.kids ?? [])[(sr.kids ?? []).findIndex((k) => k.type === 0x1051 && k.data[0] === 0) + 1];
    const text = after?.type === 0x100d ? seriesText(after) : undefined;
    const ref = (b: Brai | undefined, c: (string | number | null)[] | undefined): ChartRef | undefined => {
      if (b?.formula) return { ref: b.formula, cache: c };
      return c?.length ? { cache: c } : undefined;
    };
    const s: ChartSeries = {
      name: nameB?.formula ? { ref: nameB.formula } : text,
      values: ref(at(1), valCache[si]) ?? { cache: [] },
    };
    const cat = ref(at(2), catCache[si]);
    if (g.kind === "scatter" || g.kind === "bubble") s.xValues = cat;
    else if (!categories && cat) categories = cat;
    if (g.kind === "bubble") s.sizes = ref(at(3), bubCache[si]);
    if (g !== base && g.kind !== base.kind) s.type = g.kind;
    if (g.secondary) s.secondaryAxis = true;
    applyFormats(s, sr, g.kind, env);
    // Trendlines and error bars are series of their own pointing back here.
    for (const aux of allSeries) {
      const sp = kid(aux, 0x104a);
      if (!sp || u16(sp.data, 0) !== si + 1) continue;
      const tr = kid(aux, 0x104b);
      if (tr) {
        const reg = tr.data[0] ?? 0, ord = tr.data[1] ?? 1;
        const type = (["poly", "exp", "log", "power", "movingAvg"] as const)[reg] ?? "linear";
        const icpt = f64(tr.data, 2);
        s.trendline = { type: reg === 0 && ord <= 1 ? "linear" : type };
        if (reg === 0 && ord > 1) s.trendline.order = ord;
        if (reg === 4) s.trendline.order = ord;
        if (Number.isFinite(icpt)) s.trendline.intercept = icpt;
        if (tr.data[10]) s.trendline.dispEq = true;
        if (tr.data[11]) s.trendline.dispRSqr = true;
        const fw = f64(tr.data, 12), bw = f64(tr.data, 20);
        if (fw > 0) s.trendline.forward = fw;
        if (bw > 0) s.trendline.backward = bw;
        const lf = kid(kidsOf(aux, 0x1006)[0], 0x1007);
        if (lf && !(u16(lf.data, 8) & 1)) s.trendline.color = env.color(u16(lf.data, 10)) ?? rgbCss(lf.data, 0);
      }
      const eb = kid(aux, 0x105b);
      if (eb) {
        const sertm = eb.data[0] ?? 3, ebsrc = eb.data[1] ?? 1;
        if (sertm !== 3 && sertm !== 4) continue; // X error bars are not modelled
        const dir = sertm === 3 ? "plus" : "minus";
        const prev = s.errorBars;
        const valueType = (["", "percentage", "fixedVal", "stdDev", "cust", "stdErr"] as const)[ebsrc] || "fixedVal";
        if (prev && prev.valueType === valueType) prev.direction = "both";
        else s.errorBars = { direction: dir, valueType, value: f64(eb.data, 4), noEndCap: !eb.data[2] || undefined };
      }
    }
    series.push(s);
  });
  // Data labels from TEXT nodes linked to a series or point.
  for (const t of kidsOf(chart, 0x1025)) {
    const l = linkOf(t);
    if (!l || l.obj !== 4) continue;
    const s = series[l.v1];
    const lb = u16(t.data, 24) & 0x40 ? undefined : labelsOf(u16(t.data, 24), false);
    if (!s || !lb || l.v2 !== 0xffff) continue;
    const pos = DLP[u16(t.data, 28) & 0xf];
    s.labels = { ...s.labels, ...lb, ...(pos ? { position: pos } : {}) };
  }
  const titleNode = kidsOf(chart, 0x1025).find((t) => linkOf(t)?.obj === 1);
  const model: Omit<ChartModel, "id" | "anchor"> = {
    kind: base.kind,
    stacked: base.stacked,
    percent: base.percent,
    series,
    categories,
    legend,
  };
  const title = titleNode ? labelText(titleNode) : undefined;
  if (title) model.title = title;
  // An automatic title shows the one series' name.
  let autoTitle: string | undefined;
  if (titleNode && !title && series.length === 1 && !(u16(titleNode.data, 24) & 0x40)) {
    const nm = series[0]!.name;
    if (typeof nm === "string" && nm) model.title = nm;
    else if (typeof nm === "object" && nm.ref) autoTitle = nm.ref;
  }
  const ts = titleNode ? textStyle(titleNode, env, titleNode) : undefined;
  if (ts) model.titleStyle = ts;
  if (legendStyle) model.legendStyle = legendStyle;
  const some = (a?: ChartAxis): ChartAxis | undefined => (a && Object.keys(a).length ? a : undefined);
  if (some(axes.x) || some(axes.y)) model.axes = { x: some(axes.x), y: some(axes.y) };
  if (plotFill) model.plotFill = plotFill;
  const chartFrame = kid(chart, 0x1032);
  const areaFill = chartFrame ? areaColor(kid(chartFrame, 0x100a)) : undefined;
  if (areaFill) model.areaFill = areaFill;
  if (base.threeD && base.kind !== "surface") model.threeD = true;
  const sp = kid(chart, 0x1044);
  if (sp) { const md = sp.data[2] ?? 0; model.blanksAs = md === 1 ? "zero" : md === 2 ? "span" : "gap"; }
  const r = base.rec, d = r.data;
  if (r.type === 0x1017) { model.overlap = i16(d, 0); model.gapWidth = u16(d, 2); }
  if (r.type === 0x1019) { model.rotation = u16(d, 0); if (u16(d, 2)) model.holeSize = u16(d, 2); }
  if (r.type === 0x1061) {
    const split = u16(d, 2), pos = u16(d, 4);
    model.ofPie = { type: d[0] === 2 ? "bar" : "pie", splitCount: split === 0 ? pos : undefined, secondSize: u16(d, 8), gapWidth: u16(d, 10) };
  }
  const anyLabels = series.some((s) => s.labels?.value);
  if (anyLabels && series.every((s) => s.labels?.value)) model.dataLabels = true;
  return { model, width, height, autoTitle };
}

/** The solid fill of an AREAFORMAT, unless it is automatic or empty. */
function areaColor(af: CRec | undefined): string | undefined {
  if (!af) return undefined;
  const d = af.data;
  if (u16(d, 10) & 1) return undefined;
  if (u16(d, 8) === 0) return undefined;
  return rgbCss(d, 0);
}

/** Colours, markers, smoothing, explosion and labels from the series' DATAFORMAT nodes. */
function applyFormats(s: ChartSeries, sr: CRec, kind: ChartKind, env: ChartEnv): void {
  const lineLike = kind === "line" || kind === "scatter" || kind === "radar" || kind === "stock";
  for (const df of kidsOf(sr, 0x1006)) {
    const xi = u16(df.data, 0);
    const lf = kid(df, 0x1007), af = kid(df, 0x100a), mf = kid(df, 0x1009), pf = kid(df, 0x100b), sf = kid(df, 0x105d), al = kid(df, 0x100c);
    const lineColor = lf && !(u16(lf.data, 8) & 1) && u16(lf.data, 4) !== 5 ? (env.color(u16(lf.data, 10)) ?? rgbCss(lf.data, 0)) : undefined;
    const fill = af && !(u16(af.data, 10) & 1) && u16(af.data, 8) ? (env.color(u16(af.data, 12)) ?? rgbCss(af.data, 0)) : undefined;
    const marker = mf && !(u16(mf.data, 10) & 1) ? (env.color(u16(mf.data, 12)) ?? rgbCss(mf.data, 0)) : undefined;
    if (xi === 0xffff) {
      const c = lineLike ? lineColor ?? marker ?? fill : fill ?? lineColor;
      if (c) s.color = c;
      if (lf && lineLike) {
        const we = i16(lf.data, 6);
        if (!(u16(lf.data, 8) & 1) && we > 0) s.lineWidth = we === 1 ? 2 : 3;
        const dash = DASHES[u16(lf.data, 4)];
        if (dash) s.dash = dash;
      }
      if (mf && lineLike) {
        const imk = u16(mf.data, 8), auto = u16(mf.data, 10) & 1;
        if (!auto) s.marker = { symbol: MARKERS[imk] ?? "none", size: Math.round(u32(mf.data, 16) / 20) || undefined };
      }
      if (sf && (u16(sf.data, 0) & 1)) s.smooth = true;
      if (pf && u16(pf.data, 0)) s.explosion = [u16(pf.data, 0)];
      if (al) { const lb = labelsOf(u16(al.data, 0), true); if (lb) s.labels = { ...s.labels, ...lb }; }
    } else if (xi < 4000) {
      const c = lineLike ? lineColor ?? fill : fill ?? lineColor;
      if (c) (s.pointColors ??= [])[xi] = c;
      if (pf && u16(pf.data, 0)) (s.explosion ??= [])[xi] = u16(pf.data, 0);
    }
  }
  // A whole-series explosion applies to every slice.
  if (s.explosion && s.explosion.length === 1 && s.explosion[0] != null) {
    const n = s.values.cache?.length ?? 0;
    if (n > 1) s.explosion = new Array(n).fill(s.explosion[0]);
  }
}
