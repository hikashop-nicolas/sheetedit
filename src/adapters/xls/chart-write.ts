// The chart model as a BIFF8 chart substream (BOF type chart ... EOF), laid out the way Excel 97-2003
// writes one: CHART, its series with their BRAI formulas, the axis groups with their chart groups,
// then the titles. Embedded charts follow their OBJ in the sheet; chart sheets are their own substream.
import { CHART_PALETTE, type ChartAxis, type ChartKind, type ChartModel, type ChartRef, type ChartSeries, type ChartTextStyle } from "../../core/chart-model";
import { body, type Bytes, RecStream } from "./bytes";

export interface ChartWriteEnv {
  /** The tokens of a reference ("Sheet1!$B$2:$B$9"), or null when it cannot be written. */
  ref: (text: string) => Uint8Array | null;
  /** A FONT index for a text style (undefined = the chart's default text). */
  font: (style: ChartTextStyle | undefined, bold?: boolean, size?: number) => number;
  /** A palette index for a colour. */
  icv: (css: string) => number;
  /** A FORMAT index for a number format code. */
  fmt: (code: string) => number;
  /** Something the model has that the chart could not keep (an i18n key). */
  warn: (key: string) => void;
}

const BAR_LIKE = new Set<ChartKind>(["column", "bar"]);
const LINE_LIKE = new Set<ChartKind>(["line", "scatter", "radar", "stock"]);
const NO_AXES = new Set<ChartKind>(["pie", "doughnut"]);
const DLP: Record<string, number> = { outEnd: 1, inEnd: 2, ctr: 3, inBase: 4, bestFit: 9, l: 10, r: 11, t: 12, b: 13 };
const MARKERS: Record<string, number> = { none: 0, square: 1, diamond: 2, triangle: 3, x: 4, star: 5, dash: 6, dot: 7, circle: 8, plus: 9, auto: 1 };
const DASHES: Record<string, number> = { dash: 1, sysDash: 1, lgDash: 1, dashDot: 3, sysDashDot: 3, lgDashDot: 3, dot: 2, sysDot: 2, lgDashDotDot: 4, sysDashDotDot: 4 };

const rgb = (b: Bytes, css: string): Bytes => b.u8(parseInt(css.slice(1, 3), 16)).u8(parseInt(css.slice(3, 5), 16)).u8(parseInt(css.slice(5, 7), 16)).u8(0);
const fixed = (b: Bytes, v: number): Bytes => { const i = Math.floor(v); return b.u16(Math.round((v - i) * 65535)).u16(i); };

/** How many points a reference covers ("Sheet1!$B$2:$B$9" -> 8). */
export function refCount(ref: string | undefined): number {
  const m = /\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?\s*$/.exec(ref ?? "");
  if (!m) return 0;
  const col = (l: string): number => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const r1 = Number(m[2]), r2 = m[4] ? Number(m[4]) : r1, c1 = col(m[1]!), c2 = m[3] ? col(m[3]) : c1;
  return (Math.abs(r2 - r1) + 1) * (Math.abs(c2 - c1) + 1);
}

/** The kinds a BIFF8 chart group can hold; the rest are written as the closest one. */
const kindOf = (k: ChartKind | undefined, base: ChartKind): ChartKind => k ?? base;

interface Grp { kind: ChartKind; secondary: boolean; icrt: number; series: number[] }

/** The chart substream for `m`. `sheet` makes it a chart sheet's (with its own WINDOW2). */
export function writeChartStream(m: ChartModel, env: ChartWriteEnv, size: { w: number; h: number }, opts: { sheet?: boolean; codeName?: string } = {}): Uint8Array {
  const rs = new RecStream();
  const rec = (type: number, b?: Bytes): void => { rs.rec(type, b); };
  const begin = (): void => rec(0x1033);
  const end = (): void => rec(0x1034);
  const base = m.kind;
  const series = m.series.filter((s) => s.values.ref || s.values.cache?.length);
  const isXY = (k: ChartKind): boolean => k === "scatter" || k === "bubble";
  const xy = isXY(base);
  // Chart groups: one per kind and axis, the base one first.
  const groups: Grp[] = [];
  const groupFor = (s: ChartSeries, i: number): Grp => {
    let kind = kindOf(s.type, base);
    if (NO_AXES.has(base)) kind = base;
    const secondary = !!s.secondaryAxis && !NO_AXES.has(base);
    // An XY group and a category group can share a chart only on different axes.
    const other = groups.find((x) => x.secondary === secondary);
    const clash = isXY(kind) !== (other ? isXY(other.kind) : secondary ? isXY(kind) : xy);
    if (NO_AXES.has(kind) || clash) { env.warn("xlsChartComboKind"); kind = other?.kind ?? base; }
    let g = groups.find((x) => x.kind === kind && x.secondary === secondary);
    if (!g) { g = { kind, secondary, icrt: groups.length, series: [] }; groups.push(g); }
    g.series.push(i);
    return g;
  };
  if (!series.length) groups.push({ kind: base, secondary: false, icrt: 0, series: [] });
  const seriesGroup = series.map((s, i) => groupFor(s, i));
  groups.sort((a, b) => Number(a.secondary) - Number(b.secondary));
  groups.forEach((g, i) => (g.icrt = i));

  const defaultFont = env.font(undefined);
  const titleFont = env.font(m.titleStyle, m.titleStyle?.bold ?? true, m.titleStyle?.size ?? 12);
  const fonts = new Set<number>([defaultFont, titleFont]);
  const fontx = (ifnt: number): void => { fonts.add(ifnt); rec(0x1026, body().u16(ifnt)); };

  // ---- the substream head ----
  rec(0x0809, body().u16(0x0600).u16(0x0020).u16(0x0dbb).u16(0x07cc).u32(0x000100c1).u32(0x00000006));
  rec(0x0014); rec(0x0015);
  rec(0x0083, body().u16(0)); rec(0x0084, body().u16(0));
  rec(0x00a1, body().u16(9).u16(100).u16(1).u16(1).u16(1).u16(0x0044).u16(600).u16(600).f64(0.5).f64(0.5).u16(1));
  rec(0x0033, body().u16(3));
  const fbiAt = rs.len;
  rec(0x0012, body().u16(0));
  rec(0x1001, body().u16(0));
  const chartBody = body().u32(0).u32(0);
  fixed(chartBody, size.w); fixed(chartBody, size.h);
  rec(0x1002, chartBody);
  begin();
  rec(0x00a0, body().u16(1).u16(1));
  rec(0x1064, body().u32(0x10000).u32(0x10000));
  frame(m.areaFill, true, 2);

  // ---- series ----
  const auxSeries: { parent: number; kind: "trend" | "err"; s: ChartSeries; dir?: "plus" | "minus" }[] = [];
  series.forEach((s, i) => {
    const g = seriesGroup[i]!;
    const vals = s.values, cats = isXY(g.kind) ? s.xValues : m.categories, sizes = s.sizes;
    const nVals = refCount(vals.ref) || vals.cache?.length || 0;
    const nCats = refCount(cats?.ref) || cats?.cache?.length || 0;
    const textCats = cats?.cache?.some((v) => typeof v === "string") ?? false;
    const nSizes = g.kind === "bubble" ? refCount(sizes?.ref) || sizes?.cache?.length || 0 : 0;
    rec(0x1003, body().u16(textCats ? 3 : 1).u16(1).u16(nCats).u16(nVals).u16(1).u16(nSizes));
    begin();
    const nameRef = typeof s.name === "object" ? s.name.ref : undefined;
    const nameText = typeof s.name === "string" ? s.name : typeof s.name === "object" && !s.name.ref ? String(s.name.cache?.[0] ?? "") : undefined;
    brai(0, nameRef);
    if (!nameRef && nameText) rec(0x100d, body().u16(0).u8(Math.min(255, nameText.length)).chars(nameText.slice(0, 255)));
    brai(1, vals.ref);
    brai(2, cats?.ref);
    brai(3, g.kind === "bubble" ? sizes?.ref : undefined);
    seriesFormat(s, i, g.kind);
    rec(0x1045, body().u16(g.icrt));
    end();
    if (s.trendline) auxSeries.push({ parent: i, kind: "trend", s });
    if (s.errorBars) {
      if (s.errorBars.valueType === "cust") env.warn("xlsChartCustomErrorBars");
      else {
        const dir = s.errorBars.direction ?? "both";
        if (dir !== "minus") auxSeries.push({ parent: i, kind: "err", s, dir: "plus" });
        if (dir !== "plus") auxSeries.push({ parent: i, kind: "err", s, dir: "minus" });
      }
    }
  });
  // Trendlines and error bars: series of their own, tied to their parent.
  auxSeries.forEach((a, k) => {
    const idx = series.length + k;
    rec(0x1003, body().u16(1).u16(1).u16(0).u16(0).u16(1).u16(0));
    begin();
    for (let id = 0; id < 4; id++) rec(0x1051, body().u8(id).u8(1).u16(0).u16(0).u16(0));
    rec(0x1006, body().u16(0xffff).u16(idx).u16(idx).u16(0));
    begin();
    const color = a.kind === "trend" ? a.s.trendline!.color : undefined;
    lineFormat(color, 0);
    rec(0x100a, body().u32(0xffffff).u32(0).u16(0).u16(0).u16(0x4e).u16(0x4d));
    rec(0x100b, body().u16(0));
    rec(0x1009, body().u32(0).u32(0).u16(0).u16(0).u16(0x4d).u16(0x4d).u32(100));
    end();
    rec(0x104a, body().u16(a.parent + 1));
    if (a.kind === "trend") {
      const t = a.s.trendline!;
      const reg = t.type === "linear" || t.type === "poly" ? 0 : t.type === "exp" ? 1 : t.type === "log" ? 2 : t.type === "power" ? 3 : 4;
      const ord = t.type === "poly" ? Math.max(2, Math.min(6, t.order ?? 2)) : t.type === "movingAvg" ? Math.max(2, t.order ?? 2) : 1;
      const b = body().u8(reg).u8(ord);
      if (t.intercept != null) b.f64(t.intercept); else b.u32(0xffffffff).u32(0xffffffff);
      b.u8(t.dispEq ? 1 : 0).u8(t.dispRSqr ? 1 : 0).f64(t.forward ?? 0).f64(t.backward ?? 0);
      rec(0x104b, b);
    } else {
      const e = a.s.errorBars!;
      const src = e.valueType === "percentage" ? 1 : e.valueType === "stdDev" ? 3 : e.valueType === "stdErr" ? 5 : 2;
      rec(0x105b, body().u8(a.dir === "plus" ? 3 : 4).u8(src).u8(e.noEndCap ? 0 : 1).u8(1).f64(e.value ?? (src === 1 ? 5 : 1)).u16(0));
    }
    end();
  });
  // Sheet properties: plot visible cells only, how blanks plot.
  rec(0x1044, body().u16(0x000a).u8(m.blanksAs === "zero" ? 1 : m.blanksAs === "span" ? 2 : 0).u8(0));
  for (const id of [2, 3]) {
    rec(0x1024, body().u16(id));
    text(0x00b1, 0x2b50);
    begin();
    pos(2, 2);
    fontx(defaultFont);
    rec(0x1051, body().u8(0).u8(1).u16(0).u16(0).u16(0));
    end();
  }
  const anySecondary = groups.some((g) => g.secondary);
  rec(0x1046, body().u16(anySecondary ? 2 : 1));
  axisParent(false);
  if (anySecondary) axisParent(true);
  // Title, then the series' data label positions.
  const auto = !!m.title && m.title === m.autoTitle && series.length === 1;
  if (m.title) {
    text(auto ? 0x00b1 : 0x0081, 0x2b20, true);
    begin();
    pos(2, 2);
    fontx(titleFont);
    rec(0x1051, body().u8(0).u8(1).u16(0).u16(0).u16(0));
    if (!auto) seriesText(m.title);
    rec(0x1027, body().u16(1).u16(0).u16(0));
    end();
  }
  series.forEach((s, i) => {
    const lb = s.labels ?? m.labels;
    const position = lb?.position && DLP[lb.position];
    if (!position || !(lb.value || lb.category || lb.percent || lb.seriesName)) return;
    const grbit = 0x0081 | (lb.value ? 0x04 : 0) | (lb.percent ? 0x1000 : 0) | (lb.category ? 0x4000 : 0) | (lb.legendKey ? 0x02 : 0);
    text(grbit, position, false, 0);
    begin();
    pos(2, 2);
    rec(0x1051, body().u8(0).u8(1).u16(0).u16(0).u16(0));
    rec(0x1027, body().u16(4).u16(i).u16(0xffff));
    end();
  });
  end();
  // The fonts the chart uses, before UNITS.
  const fbi = new RecStream();
  for (const f of fonts) fbi.rec(0x1060, body().u16(0x1a28).u16(0x0f9c).u16(0x00c8).u16(0).u16(f));
  rec(0x0200, body().u32(0).u32(Math.max(0, ...series.map((s) => s.values.cache?.length ?? 0))).u16(0).u16(series.length).u16(0));
  seriesData();
  if (opts.sheet) {
    rec(0x023e, body().u16(0x0000).u16(0).u16(0).u16(0).u16(0));
    if (opts.codeName) rec(0x01ba, body().str16(opts.codeName));
  }
  rec(0x000a);
  const all = rs.out();
  const out = new Uint8Array(all.length + fbi.len);
  out.set(all.subarray(0, fbiAt));
  out.set(fbi.out(), fbiAt);
  out.set(all.subarray(fbiAt), fbiAt + fbi.len);
  return out;

  function brai(id: number, ref: string | undefined): void {
    const rgce = ref ? env.ref(ref) : null;
    if (ref && !rgce) env.warn("xlsChartRefDropped");
    if (rgce) rec(0x1051, body().u8(id).u8(2).u16(0).u16(0).u16(rgce.length).bytes(rgce));
    else rec(0x1051, body().u8(id).u8(1).u16(0).u16(0).u16(0));
  }

  function seriesText(t: string): void {
    rec(0x100d, body().u16(0).u8(Math.min(255, t.length)).chars(t.slice(0, 255)));
  }

  function lineFormat(color: string | undefined, we: number, none = false, flags?: number): void {
    const b = body();
    if (color) rgb(b, color); else b.u32(0);
    b.u16(none ? 5 : 0).u16(we).u16(flags ?? (color || none ? 0 : 0x0009)).u16(color ? env.icv(color) : 0x4d);
    rec(0x1007, b);
  }

  function areaFormat(color: string | undefined, none = false): void {
    const b = body();
    if (color) rgb(b, color); else b.u32(0xffffff);
    b.u32(0).u16(none ? 0 : 1).u16(color || none ? 0 : 1).u16(color ? env.icv(color) : 0x4e).u16(0x4d);
    rec(0x100a, b);
  }

  function frame(fill: string | undefined, chartArea: boolean, flags: number): void {
    rec(0x1032, body().u16(0).u16(flags));
    begin();
    if (chartArea) { lineFormat(undefined, 0xffff); areaFormat(fill); }
    else { lineFormat(undefined, 0, true); areaFormat(fill, !fill); }
    end();
  }

  function pos(tl: number, br: number, rect: [number, number, number, number] = [0, 0, 0, 0]): void {
    rec(0x104f, body().u16(tl).u16(br).u32(rect[0]).u32(rect[1]).u32(rect[2]).u32(rect[3]));
  }

  function text(grbit: number, grbit2: number, title = false, bkg = 1): void {
    const b = body().u8(2).u8(2).u16(bkg).u32(0);
    if (title) b.u32(0x05ea).u32(0x4e).u32(0x03c5).u32(0x01c3);
    else b.u32(0xffffffce).u32(0xffffffac).u32(0).u32(0);
    rec(0x1025, b.u16(grbit).u16(0x4d).u16(grbit2).u16(0));
  }

  function seriesFormat(s: ChartSeries, i: number, kind: ChartKind): void {
    const lineLike = LINE_LIKE.has(kind);
    rec(0x1006, body().u16(0xffff).u16(i).u16(i).u16(0));
    begin();
    if (BAR_LIKE.has(kind)) rec(0x105f, body().u8(0).u8(0));
    const color = s.color;
    const stock = kind === "stock";
    if (color || stock || s.lineWidth || s.dash) {
      const we = s.lineWidth ? (s.lineWidth >= 2.5 ? 2 : s.lineWidth >= 1.5 ? 1 : 0) : 0;
      const b = body();
      if (color) rgb(b, color); else b.u32(0);
      b.u16(stock ? 5 : (s.dash && DASHES[s.dash]) || 0).u16(we).u16(color || stock ? 0 : 0x0009).u16(color ? env.icv(color) : 0x4d);
      rec(0x1007, b);
      if (lineLike) areaFormat(undefined); else areaFormat(color);
      rec(0x100b, body().u16(s.explosion?.find((x) => x != null) ?? 0));
    } else if (s.explosion?.some((x) => x != null)) {
      lineFormat(undefined, 0);
      areaFormat(undefined);
      rec(0x100b, body().u16(s.explosion.find((x) => x != null) ?? 0));
    }
    if (s.smooth && lineLike) rec(0x105d, body().u16(1));
    if (lineLike && (s.marker || stock)) {
      const imk = stock ? 0 : MARKERS[s.marker?.symbol ?? "auto"] ?? 1;
      // A stated marker is not automatic, so it needs colours: the series' own or the palette's.
      const mc = color ?? m.palette?.[i % m.palette.length] ?? CHART_PALETTE[i % CHART_PALETTE.length]!;
      const b = body();
      rgb(b, mc); rgb(b, mc);
      b.u16(imk).u16(s.marker || stock ? 0 : 1).u16(env.icv(mc)).u16(env.icv(mc)).u32(Math.round((s.marker?.size ?? 5) * 20));
      rec(0x1009, b);
    }
    const lb = s.labels ?? m.labels ?? (m.dataLabels ? { value: true } : undefined);
    if (lb && (lb.value || lb.percent || lb.category || lb.seriesName)) {
      const f = (lb.value ? 1 : 0) | (lb.percent && lb.category ? 4 : lb.percent ? 2 : 0) | (lb.category && !lb.percent ? 0x10 : 0) | (lb.seriesName ? 0x40 : 0);
      rec(0x100c, body().u16(f));
    }
    end();
    // Point colours and exploded slices.
    const n = Math.max(s.pointColors?.length ?? 0, s.explosion?.length ?? 0);
    for (let p = 0; p < n && p < 4000; p++) {
      const pc = s.pointColors?.[p], ex = s.explosion?.[p];
      if (!pc && ex == null) continue;
      rec(0x1006, body().u16(p).u16(i).u16(i).u16(0));
      begin();
      if (BAR_LIKE.has(kind)) rec(0x105f, body().u8(0).u8(0));
      lineFormat(lineLike ? pc : undefined, 0);
      areaFormat(lineLike ? undefined : pc);
      rec(0x100b, body().u16(ex ?? 0));
      end();
    }
  }

  function axis(type: number, a: ChartAxis | undefined, value: boolean, hidden: boolean): void {
    rec(0x101d, body().u16(type).zeros(16));
    begin();
    if (value) {
      const fl = 0x011f & ~(a?.min != null ? 1 : 0) & ~(a?.max != null ? 2 : 0);
      rec(0x101f, body().f64(a?.min ?? 0).f64(a?.max ?? 0).f64(0).f64(0).f64(0).u16(fl));
      if (a?.numFmt) rec(0x104e, body().u16(env.fmt(a.numFmt)));
    } else {
      rec(0x1020, body().u16(1).u16(1).u16(1).u16(1));
      rec(0x1062, body().u16(0).u16(0).u16(1).u16(0).u16(1).u16(0).u16(0).u16(0).u16(a?.date ? 0x6f : 0xef));
    }
    const tick = body().u8(hidden ? 0 : 2).u8(0).u8(hidden ? 0 : 3).u8(1).u32(0).zeros(16).u16(0x0023).u16(0x4d).u16(0);
    rec(0x101e, tick);
    if (a?.labelStyle) fontx(env.font(a.labelStyle));
    if (value && type === 1 && !hidden) {
      rec(0x1021, body().u16(1));
      lineFormat(undefined, 0xffff);
    }
    end();
  }

  function axisTitle(t: string, style: ChartTextStyle | undefined, link: number): void {
    text(0x0081, 0x2b20, false, 1);
    begin();
    pos(2, 2);
    fontx(env.font(style, style?.bold ?? true, style?.size));
    rec(0x1051, body().u8(0).u8(1).u16(0).u16(0).u16(0));
    seriesText(t);
    rec(0x1027, body().u16(link).u16(0).u16(0));
    end();
  }

  function axisParent(secondary: boolean): void {
    rec(0x1041, body().u16(secondary ? 1 : 0).u32(0x028c).u32(0x0138).u32(0x0a47).u32(0x0992));
    begin();
    pos(2, 2, [0x4f, 0x84, 0x0c84, 0x0e98]);
    const mine = groups.filter((g) => g.secondary === secondary);
    if (!NO_AXES.has(base)) {
      const ax = m.axes;
      if (mine.some((g) => isXY(g.kind))) axis(0, ax?.x, true, secondary);
      else axis(0, ax?.x, false, secondary);
      axis(1, ax?.y, true, false);
      if (!secondary) {
        if (ax?.x?.title) axisTitle(ax.x.title, ax.x.titleStyle, 3);
        if (ax?.y?.title) axisTitle(ax.y.title, ax.y.titleStyle, 2);
      }
      rec(0x1035);
      frame(m.plotFill, false, 3);
    }
    for (const g of mine) chartGroup(g, !secondary && g === mine[0]);
    end();
  }

  function chartGroup(g: Grp, withLegend: boolean): void {
    const pie = g.kind === "pie" || g.kind === "doughnut";
    rec(0x1014, body().zeros(16).u16(pie || (series.length === 1 && BAR_LIKE.has(g.kind) && !!series[0]!.pointColors) ? 1 : 0).u16(g.icrt));
    begin();
    const stacked = m.stacked || m.percent ? 1 : 0, pct = m.percent ? 1 : 0;
    switch (g.kind) {
      case "column": case "bar":
        rec(0x1017, body().u16((m.overlap ?? 0) & 0xffff).u16(m.gapWidth ?? 150).u16((g.kind === "bar" ? 1 : 0) | (stacked ? 2 : 0) | (pct ? 4 : 0)));
        break;
      case "line": case "stock": rec(0x1018, body().u16(g.kind === "stock" ? 0 : stacked | (pct << 1))); break;
      case "area": rec(0x101a, body().u16(stacked | (pct << 1))); break;
      case "pie": case "doughnut":
        if (m.ofPie && g.kind === "pie") {
          const op = m.ofPie;
          rec(0x1061, body().u8(op.type === "bar" ? 2 : 1).u8(op.splitCount == null ? 1 : 0).u16(0).u16(op.splitCount ?? 2).u16(0).u16(op.secondSize ?? 75).u16(op.gapWidth ?? 100).f64(0).u16(1));
        } else rec(0x1019, body().u16(m.rotation ?? 0).u16(g.kind === "doughnut" ? m.holeSize ?? 50 : 0).u16(2));
        break;
      case "scatter": case "bubble": rec(0x101b, body().u16(100).u16(1).u16(g.kind === "bubble" ? 1 : 0)); break;
      case "radar": rec(0x103e, body().u16(1).u16(0)); break;
      case "surface": rec(0x103f, body().u16(0)); break;
    }
    rec(0x1022, body().zeros(8).u16(0x000f));
    if (m.threeD || g.kind === "surface") {
      const pieLike = pie;
      rec(0x103a, body().u16(pieLike ? 0 : 20).u16(pieLike ? 30 : 15).u16(30).u16(100).u16(100).u16(150).u16(pieLike ? 0 : 0x14));
    }
    if (withLegend && m.legend?.show !== false && m.legend) {
      const p = m.legend.pos;
      const w = p === "bottom" ? 0 : p === "top" ? 2 : p === "left" ? 4 : 3;
      rec(0x1015, body().u32(0x0d30).u32(0x0565).u32(0x0253).u32(0x012c).u8(w).u8(1).u16(p === "left" || p === "right" ? 0x1f : 0x0f));
      begin();
      pos(5, 2);
      text(0x00b1, 0x3d90);
      begin();
      pos(2, 2);
      if (m.legendStyle) fontx(env.font(m.legendStyle));
      rec(0x1051, body().u8(0).u8(1).u16(0).u16(0).u16(0));
      end();
      end();
    }
    if (g.kind === "stock") {
      rec(0x101c, body().u16(1));
      lineFormat("#000000", 0);
      if (g.series.length >= 4) {
        for (const c of ["#ffffff", "#000000"]) {
          rec(0x103d, body().u16(150));
          begin();
          lineFormat("#000000", 0);
          areaFormat(c);
          end();
        }
      }
    }
    end();
  }

  function seriesData(): void {
    // SIINDEX 2 (categories), 1 (values), 3 (bubble sizes), each with the cached points of
    // series that have no reference to read them from.
    const cell = (row: number, col: number, v: string | number | null | undefined): void => {
      if (typeof v === "number" && Number.isFinite(v)) rec(0x0203, body().u16(row).u16(col).u16(0).f64(v));
      else if (typeof v === "string") rec(0x0204, body().u16(row).u16(col).u16(0).str16(v.slice(0, 255)));
    };
    const literal = (r: ChartRef | undefined): (string | number | null)[] | undefined => (r && !r.ref ? r.cache : undefined);
    for (const k of [2, 1, 3]) {
      rec(0x1065, body().u16(k));
      series.forEach((s, i) => {
        const src = k === 1 ? literal(s.values) : k === 2 ? literal(isXY(seriesGroup[i]!.kind) ? s.xValues : m.categories) : literal(s.sizes);
        src?.forEach((v, p) => cell(p, i, v));
      });
    }
  }
}
