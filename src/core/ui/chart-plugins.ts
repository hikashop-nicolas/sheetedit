// Custom Chart.js plugins for chart features Chart.js has no built-in for: regression trendlines
// and error bars. Each is a global plugin that no-ops unless a dataset carries its config
// (dataset.trendline / dataset.errorBars), so it never affects charts that do not use it. Kept
// dependency-free: the regressions and the drawing are computed here.

import type { ChartErrorBars, ChartTrendline } from "../chart-model";

interface Pt { x: number; y: number }
type FitFn = (x: number) => number;

const mean = (a: number[]): number => (a.length ? a.reduce((t, v) => t + v, 0) / a.length : 0);

/** R-squared of a fit against the observed points. */
function rSquared(pts: Pt[], f: FitFn): number {
  const my = mean(pts.map((p) => p.y));
  let ssr = 0, sst = 0;
  for (const p of pts) { const e = p.y - f(p.x); ssr += e * e; const d = p.y - my; sst += d * d; }
  return sst ? 1 - ssr / sst : 0;
}

/** Least-squares line y = m x + b (b forced when intercept is given). */
function linear(pts: Pt[], intercept?: number): { m: number; b: number } {
  if (intercept != null) {
    let sxx = 0, sxy = 0;
    for (const p of pts) { sxx += p.x * p.x; sxy += (p.y - intercept) * p.x; }
    return { m: sxx ? sxy / sxx : 0, b: intercept };
  }
  const n = pts.length;
  let sx = 0, sy = 0, sxx = 0, sxy = 0;
  for (const p of pts) { sx += p.x; sy += p.y; sxx += p.x * p.x; sxy += p.x * p.y; }
  const d = n * sxx - sx * sx;
  const m = d ? (n * sxy - sx * sy) / d : 0;
  return { m, b: (sy - m * sx) / n };
}

/** Solve a linear system A x = b by Gaussian elimination with partial pivoting. */
function gauss(A: number[][], b: number[]): number[] {
  const n = b.length;
  const M = A.map((row, i) => [...row, b[i]]);
  for (let c = 0; c < n; c++) {
    let piv = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[piv][c])) piv = r;
    [M[c], M[piv]] = [M[piv], M[c]];
    if (Math.abs(M[c][c]) < 1e-12) continue;
    for (let r = 0; r < n; r++) if (r !== c) { const f = M[r][c] / M[c][c]; for (let k = c; k <= n; k++) M[r][k] -= f * M[c][k]; }
  }
  return M.map((row, i) => (Math.abs(M[i][i]) < 1e-12 ? 0 : row[n] / M[i][i]));
}

/** Polynomial coefficients [c0..ck] for y = sum ci x^i (least squares, order k). */
function poly(pts: Pt[], order: number): number[] {
  const n = order + 1;
  const X = Array.from({ length: n }, () => new Array(n).fill(0));
  const Y = new Array(n).fill(0);
  for (const p of pts) for (let i = 0; i < n; i++) { for (let j = 0; j < n; j++) X[i][j] += p.x ** (i + j); Y[i] += p.x ** i * p.y; }
  return gauss(X, Y);
}

const fmt = (v: number): string => {
  if (!isFinite(v)) return "0";
  const a = Math.abs(v);
  return a !== 0 && (a < 1e-3 || a >= 1e5) ? v.toExponential(2) : Number(v.toPrecision(4)).toString();
};

/** Fit a trendline: returns the fitted function, an equation string, and R-squared.
    Moving average returns discrete points instead of a continuous function. */
export function fitTrendline(t: ChartTrendline, pts: Pt[]): { f?: FitFn; points?: Pt[]; equation: string; r2: number } {
  const sorted = [...pts].sort((a, b) => a.x - b.x);
  switch (t.type) {
    case "linear": {
      const { m, b } = linear(sorted, t.intercept);
      const f: FitFn = (x) => m * x + b;
      return { f, equation: `y = ${fmt(m)}x ${b < 0 ? "-" : "+"} ${fmt(Math.abs(b))}`, r2: rSquared(sorted, f) };
    }
    case "poly": {
      const order = Math.max(2, Math.min(6, t.order ?? 2));
      const c = poly(sorted, order);
      const f: FitFn = (x) => c.reduce((s, ci, i) => s + ci * x ** i, 0);
      const eq = "y = " + c.map((ci, i) => (i === 0 ? fmt(ci) : `${fmt(ci)}x${i > 1 ? "^" + i : ""}`)).reverse().join(" + ");
      return { f, equation: eq, r2: rSquared(sorted, f) };
    }
    case "exp": {
      const pos = sorted.filter((p) => p.y > 0);
      const { m, b } = linear(pos.map((p) => ({ x: p.x, y: Math.log(p.y) })));
      const a = Math.exp(b);
      const f: FitFn = (x) => a * Math.exp(m * x);
      return { f, equation: `y = ${fmt(a)}e^(${fmt(m)}x)`, r2: rSquared(pos, f) };
    }
    case "log": {
      const pos = sorted.filter((p) => p.x > 0);
      const { m, b } = linear(pos.map((p) => ({ x: Math.log(p.x), y: p.y })));
      const f: FitFn = (x) => (x > 0 ? m * Math.log(x) + b : NaN);
      return { f, equation: `y = ${fmt(m)}ln(x) ${b < 0 ? "-" : "+"} ${fmt(Math.abs(b))}`, r2: rSquared(pos, f) };
    }
    case "power": {
      const pos = sorted.filter((p) => p.x > 0 && p.y > 0);
      const { m, b } = linear(pos.map((p) => ({ x: Math.log(p.x), y: Math.log(p.y) })));
      const a = Math.exp(b);
      const f: FitFn = (x) => (x > 0 ? a * x ** m : NaN);
      return { f, equation: `y = ${fmt(a)}x^${fmt(m)}`, r2: rSquared(pos, f) };
    }
    case "movingAvg": {
      const period = Math.max(2, t.order ?? 2);
      const points: Pt[] = [];
      for (let i = period - 1; i < sorted.length; i++) {
        let s = 0; for (let j = i - period + 1; j <= i; j++) s += sorted[j].y;
        points.push({ x: sorted[i].x, y: s / period });
      }
      return { points, equation: `${period}-pt moving avg`, r2: 0 };
    }
  }
}

type Scale = { getPixelForValue: (v: number) => number };
interface Elem { x: number; y: number }
type StockRole = "open" | "high" | "low" | "close";
interface DrawChart { ctx: CanvasRenderingContext2D; chartArea: { left: number; right: number; top: number; bottom: number }; data: { datasets: { data: unknown[]; trendline?: ChartTrendline; errorBars?: ChartErrorBars; borderColor?: string; stockRole?: StockRole; surfaceRow?: number; surfaceName?: string }[] }; getDatasetMeta: (i: number) => { xScale?: Scale; yScale?: Scale; hidden?: boolean; data: Elem[] } }

/** Extract (x,y) points from a dataset's raw data (category index or {x,y}). */
function ptsOfDataset(data: unknown[]): Pt[] {
  const out: Pt[] = [];
  for (let j = 0; j < data.length; j++) {
    const v = data[j];
    const y = v && typeof v === "object" ? (v as { y?: number }).y : (v as number);
    if (y == null || isNaN(y)) continue;
    const x = v && typeof v === "object" && "x" in (v as object) ? (v as { x: number }).x : j;
    out.push({ x, y });
  }
  return out;
}

/** Draws regression trendlines for any dataset carrying a `trendline` config. */
export const trendlinePlugin = {
  id: "sheeteditTrendline",
  afterDatasetsDraw(chart: DrawChart): void {
    const { ctx, chartArea } = chart;
    chart.data.datasets.forEach((ds, i) => {
      const t = ds.trendline;
      if (!t) return;
      const meta = chart.getDatasetMeta(i);
      if (meta.hidden || !meta.xScale || !meta.yScale) return;
      const pts = ptsOfDataset(ds.data);
      if (pts.length < 2) return;
      const fit = fitTrendline(t, pts);
      if (!fit) return;
      const xs = pts.map((p) => p.x);
      const xmin = Math.min(...xs) - (t.backward ?? 0);
      const xmax = Math.max(...xs) + (t.forward ?? 0);
      const color = t.color ?? "#666";
      ctx.save();
      ctx.beginPath();
      ctx.rect(chartArea.left, chartArea.top, chartArea.right - chartArea.left, chartArea.bottom - chartArea.top);
      ctx.clip();
      ctx.strokeStyle = color;
      ctx.lineWidth = 1.5;
      ctx.setLineDash([6, 4]);
      ctx.beginPath();
      const plot = (px: number, py: number, first: boolean): void => { if (first) ctx.moveTo(px, py); else ctx.lineTo(px, py); };
      if (fit.points) {
        fit.points.forEach((p, k) => plot(meta.xScale!.getPixelForValue(p.x), meta.yScale!.getPixelForValue(p.y), k === 0));
      } else if (fit.f) {
        const N = 64;
        for (let k = 0; k <= N; k++) { const x = xmin + (xmax - xmin) * (k / N); const y = fit.f(x); if (isNaN(y)) continue; plot(meta.xScale!.getPixelForValue(x), meta.yScale!.getPixelForValue(y), k === 0); }
      }
      ctx.stroke();
      ctx.setLineDash([]);
      // Equation / R-squared label near the top-right of the plot.
      const lines: string[] = [];
      if (t.dispEq) lines.push(fit.equation);
      if (t.dispRSqr) lines.push(`R² = ${fmt(fit.r2)}`);
      if (lines.length) {
        ctx.fillStyle = color;
        ctx.font = "11px system-ui, sans-serif";
        ctx.textAlign = "right";
        ctx.textBaseline = "top";
        lines.forEach((ln, k) => ctx.fillText(ln, chartArea.right - 6, chartArea.top + 4 + k * 14));
      }
      ctx.restore();
    });
  },
};

const yOfRaw = (v: unknown): number | null => {
  const y = v && typeof v === "object" ? (v as { y?: number }).y : (v as number);
  return y == null || isNaN(y) ? null : y;
};
function stddev(vals: number[]): number {
  if (vals.length < 2) return 0;
  const m = mean(vals);
  return Math.sqrt(vals.reduce((t, v) => t + (v - m) ** 2, 0) / (vals.length - 1));
}

/** Draws error-bar whiskers for any dataset carrying an `errorBars` config. */
export const errorBarsPlugin = {
  id: "sheeteditErrorBars",
  afterDatasetsDraw(chart: DrawChart): void {
    const { ctx } = chart;
    chart.data.datasets.forEach((ds, i) => {
      const eb = ds.errorBars;
      if (!eb) return;
      const meta = chart.getDatasetMeta(i);
      if (meta.hidden || !meta.yScale) return;
      const ys = ds.data.map(yOfRaw).filter((v): v is number => v != null);
      const sd = stddev(ys);
      const magOf = (y: number, idx: number): { plus: number; minus: number } => {
        switch (eb.valueType) {
          case "fixedVal": return { plus: eb.value ?? 0, minus: eb.value ?? 0 };
          case "percentage": return { plus: Math.abs(y) * (eb.value ?? 0) / 100, minus: Math.abs(y) * (eb.value ?? 0) / 100 };
          case "stdDev": return { plus: sd * (eb.value ?? 1), minus: sd * (eb.value ?? 1) };
          case "stdErr": { const e = sd / Math.sqrt(Math.max(1, ys.length)); return { plus: e, minus: e }; }
          case "cust": return { plus: eb.plus?.[idx] ?? 0, minus: eb.minus?.[idx] ?? 0 };
        }
      };
      const dir = eb.direction ?? "both";
      const cap = eb.noEndCap ? 0 : 4;
      ctx.save();
      ctx.strokeStyle = ds.borderColor ?? "#555";
      ctx.lineWidth = 1;
      ds.data.forEach((v, idx) => {
        const y = yOfRaw(v);
        const el = meta.data[idx];
        if (y == null || !el) return;
        const px = el.x;
        const { plus, minus } = magOf(y, idx);
        const drawTo = (val: number): void => {
          const py = meta.yScale!.getPixelForValue(val);
          ctx.beginPath(); ctx.moveTo(px, el.y); ctx.lineTo(px, py); ctx.stroke();
          if (cap) { ctx.beginPath(); ctx.moveTo(px - cap, py); ctx.lineTo(px + cap, py); ctx.stroke(); }
        };
        if (dir !== "minus" && plus) drawTo(y + plus);
        if (dir !== "plus" && minus) drawTo(y - minus);
      });
      ctx.restore();
    });
  },
};

/** Draws candlesticks (open-high-low-close) or high-low-close bars for stock charts. The O/H/L/C
    series are invisible line datasets tagged with stockRole; this reads them and draws. */
export const stockPlugin = {
  id: "sheeteditStock",
  afterDatasetsDraw(chart: DrawChart): void {
    const dss = chart.data.datasets;
    if (!dss.some((d) => d.stockRole)) return;
    const roleMeta: Partial<Record<StockRole, { data: unknown[]; meta: { xScale?: Scale; yScale?: Scale } }>> = {};
    dss.forEach((d, i) => { if (d.stockRole) roleMeta[d.stockRole] = { data: d.data, meta: chart.getDatasetMeta(i) }; });
    const close = roleMeta.close ?? roleMeta.low;
    if (!close?.meta.xScale || !close.meta.yScale) return;
    const xs = close.meta.xScale, ys = close.meta.yScale;
    const n = close.data.length;
    const step = n > 1 ? Math.abs(xs.getPixelForValue(1) - xs.getPixelForValue(0)) : 40;
    const bw = Math.max(3, Math.min(24, step * 0.5));
    const ctx = chart.ctx;
    const val = (role: StockRole, j: number): number | null => { const r = roleMeta[role]; if (!r) return null; return yOfRaw(r.data[j]); };
    ctx.save();
    ctx.lineWidth = 1.2;
    for (let j = 0; j < n; j++) {
      const hi = val("high", j), lo = val("low", j), op = val("open", j), cl = val("close", j);
      if (hi == null || lo == null) continue;
      const px = xs.getPixelForValue(j);
      const up = op != null && cl != null ? cl >= op : true;
      const color = up ? "#2e9e5b" : "#d1493f";
      ctx.strokeStyle = color; ctx.fillStyle = color;
      ctx.beginPath(); ctx.moveTo(px, ys.getPixelForValue(hi)); ctx.lineTo(px, ys.getPixelForValue(lo)); ctx.stroke();
      if (op != null && cl != null) {
        const y1 = ys.getPixelForValue(op), y2 = ys.getPixelForValue(cl);
        const top = Math.min(y1, y2), h = Math.max(1, Math.abs(y2 - y1));
        ctx.fillRect(px - bw / 2, top, bw, h);
      } else if (cl != null) {
        const yc = ys.getPixelForValue(cl); ctx.beginPath(); ctx.moveTo(px, yc); ctx.lineTo(px + bw / 2, yc); ctx.stroke();
      }
    }
    ctx.restore();
  },
};

// A blue -> cyan -> yellow -> red colour ramp for the surface heatmap.
const HEAT_STOPS: [number, [number, number, number]][] = [
  [0, [49, 54, 149]], [0.25, [69, 117, 180]], [0.5, [255, 255, 191]], [0.75, [253, 174, 97]], [1, [165, 0, 38]],
];
function heat(t: number): string {
  t = Math.max(0, Math.min(1, t));
  for (let i = 1; i < HEAT_STOPS.length; i++) {
    if (t <= HEAT_STOPS[i][0]) {
      const [t0, c0] = HEAT_STOPS[i - 1], [t1, c1] = HEAT_STOPS[i];
      const f = (t - t0) / (t1 - t0 || 1);
      const c = c0.map((a, k) => Math.round(a + (c1[k] - a) * f));
      return `rgb(${c[0]},${c[1]},${c[2]})`;
    }
  }
  return "rgb(165,0,38)";
}

/** Renders a surface chart as a heatmap: one row per series, one column per category, cell colour
    mapped from the value. Chart.js has no 3D surface; this is the standard 2D flattening. */
export const surfacePlugin = {
  id: "sheeteditSurface",
  afterDatasetsDraw(chart: DrawChart): void {
    const rows = chart.data.datasets.map((d, i) => ({ d, i })).filter((r) => r.d.surfaceRow != null);
    if (!rows.length) return;
    const first = chart.getDatasetMeta(rows[0].i);
    const xs = first.xScale;
    if (!xs) return;
    const { ctx, chartArea } = chart;
    const nCols = Math.max(...rows.map((r) => r.d.data.length));
    const allVals = rows.flatMap((r) => r.d.data.map(yOfRaw).filter((v): v is number => v != null));
    if (!allVals.length) return;
    const lo = Math.min(...allVals), hi = Math.max(...allVals);
    const step = nCols > 1 ? Math.abs(xs.getPixelForValue(1) - xs.getPixelForValue(0)) : (chartArea.right - chartArea.left) / Math.max(1, nCols);
    const rowH = (chartArea.bottom - chartArea.top) / rows.length;
    ctx.save();
    ctx.font = "10px system-ui, sans-serif";
    ctx.textBaseline = "middle";
    rows.forEach((r, ri) => {
      const top = chartArea.top + ri * rowH;
      for (let j = 0; j < r.d.data.length; j++) {
        const v = yOfRaw(r.d.data[j]);
        if (v == null) continue;
        const cx = xs.getPixelForValue(j);
        ctx.fillStyle = heat(hi > lo ? (v - lo) / (hi - lo) : 0.5);
        ctx.fillRect(cx - step / 2, top, step, rowH);
      }
      ctx.strokeStyle = "rgba(255,255,255,0.5)";
      ctx.strokeRect(chartArea.left, top, chartArea.right - chartArea.left, rowH);
      if (r.d.surfaceName) { ctx.fillStyle = "#111"; ctx.textAlign = "left"; ctx.fillText(r.d.surfaceName, chartArea.left + 3, top + rowH / 2); }
    });
    ctx.restore();
  },
};

interface BgChart { ctx: CanvasRenderingContext2D; width: number; height: number; chartArea: { left: number; right: number; top: number; bottom: number } }
/** Fills the chart-area (whole canvas) and/or plot-area background from spPr fills. Reads its
    colours from the plugin options (options.plugins.sheeteditBg). */
export const backgroundPlugin = {
  id: "sheeteditBg",
  beforeDraw(chart: BgChart, _args: unknown, opts: { area?: string; plot?: string } | undefined): void {
    if (!opts) return;
    const { ctx, chartArea } = chart;
    if (opts.area) { ctx.save(); ctx.fillStyle = opts.area; ctx.fillRect(0, 0, chart.width, chart.height); ctx.restore(); }
    if (opts.plot && chartArea) { ctx.save(); ctx.fillStyle = opts.plot; ctx.fillRect(chartArea.left, chartArea.top, chartArea.right - chartArea.left, chartArea.bottom - chartArea.top); ctx.restore(); }
  },
};

interface OfPieSec { label: string; value: number; color: string }
interface Arc { x: number; y: number; startAngle: number; endAngle: number; outerRadius: number }
interface OfPieChart {
  ctx: CanvasRenderingContext2D; width: number;
  chartArea: { left: number; right: number; top: number; bottom: number };
  data: { datasets: { ofPie?: { type: "pie" | "bar"; secondary: OfPieSec[] } }[] };
  getDatasetMeta: (i: number) => { data: Arc[] };
}

/** Draws the secondary plot of a pie-of-pie / bar-of-pie: the "Other" slice's breakdown as a small
    pie or stacked bar in the right strip (reserved via a static layout padding on the config), with
    a connector to the Other slice. */
export const ofPiePlugin = {
  id: "sheeteditOfPie",
  afterDatasetsDraw(chart: OfPieChart): void {
    const ds = chart.data.datasets[0];
    if (!ds?.ofPie) return;
    const arcs = chart.getDatasetMeta(0).data;
    if (!arcs.length) return;
    const other = arcs[arcs.length - 1];
    const { ctx, chartArea } = chart;
    const stripL = chartArea.right + 10, stripR = chart.width - 10;
    const cx = (stripL + stripR) / 2, cy = (chartArea.top + chartArea.bottom) / 2;
    const R = Math.max(20, Math.min((stripR - stripL) / 2, (chartArea.bottom - chartArea.top) / 2) * 0.8);
    const sec = ds.ofPie.secondary;
    const total = sec.reduce((t, s) => t + s.value, 0) || 1;
    ctx.save();
    // connector from the Other slice's outer edge to the secondary plot
    const oa = (other.startAngle + other.endAngle) / 2;
    ctx.strokeStyle = "#bbb"; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(other.x + Math.cos(oa) * other.outerRadius, other.y + Math.sin(oa) * other.outerRadius); ctx.lineTo(cx - R, cy); ctx.stroke();
    ctx.font = "10px system-ui, sans-serif"; ctx.textAlign = "center"; ctx.textBaseline = "middle";
    if (ds.ofPie.type === "bar") {
      const bw = Math.min(44, (stripR - stripL) * 0.55), bh = R * 2;
      let yy = cy - bh / 2;
      for (const s of sec) { const h = bh * (s.value / total); ctx.fillStyle = s.color; ctx.fillRect(cx - bw / 2, yy, bw, h); if (h > 12) { ctx.fillStyle = "#fff"; ctx.fillText(s.label, cx, yy + h / 2); } yy += h; }
    } else {
      let a0 = -Math.PI / 2;
      for (const s of sec) {
        const a1 = a0 + Math.PI * 2 * (s.value / total);
        ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R, a0, a1); ctx.closePath();
        ctx.fillStyle = s.color; ctx.fill(); ctx.strokeStyle = "#fff"; ctx.lineWidth = 1; ctx.stroke();
        if (a1 - a0 > 0.35) { ctx.fillStyle = "#fff"; ctx.fillText(s.label, cx + Math.cos((a0 + a1) / 2) * R * 0.6, cy + Math.sin((a0 + a1) / 2) * R * 0.6); }
        a0 = a1;
      }
    }
    ctx.restore();
  },
};

interface MultiLevelChart { ctx: CanvasRenderingContext2D; scales?: { x?: { bottom: number; getPixelForValue: (i: number) => number } } }
/** Draws the outer level(s) of a multi-level (tiered) category axis below the innermost labels
    (which Chart.js already draws): each outer level is a row of grouped, centred labels boxed to
    span the categories it covers. The levels come from the plugin options (innermost first). */
export const multiLevelAxisPlugin = {
  id: "sheeteditMultiLevel",
  afterDraw(chart: MultiLevelChart, _args: unknown, opts: (string | number | null)[][] | undefined): void {
    const levels = opts;
    if (!levels || levels.length < 2) return;
    const xs = chart.scales?.x;
    if (!xs) return;
    const { ctx } = chart;
    const n = Math.max(...levels.map((l) => l.length));
    const half = n > 1 ? Math.abs(xs.getPixelForValue(1) - xs.getPixelForValue(0)) / 2 : 20;
    const rowH = 18;
    let y = xs.bottom + 1;
    ctx.save();
    ctx.font = "11px system-ui, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    for (let L = 1; L < levels.length; L++) {
      const level = levels[L]!;
      const defs: number[] = [];
      for (let i = 0; i < level.length; i++) if (level[i] != null && level[i] !== "") defs.push(i);
      if (!defs.length) { y += rowH; continue; }
      for (let g = 0; g < defs.length; g++) {
        const a = defs[g]!, b = g + 1 < defs.length ? defs[g + 1]! - 1 : n - 1;
        const left = xs.getPixelForValue(a) - half, right = xs.getPixelForValue(b) + half;
        ctx.strokeStyle = "rgba(0,0,0,0.15)";
        ctx.lineWidth = 1;
        ctx.strokeRect(left, y, right - left, rowH);
        ctx.fillStyle = "#555";
        ctx.fillText(String(level[a]), (left + right) / 2, y + rowH / 2);
      }
      y += rowH;
    }
    ctx.restore();
  },
};

// ---- Pseudo-3D (isometric extrusion), matching how Excel draws its "3-D" charts ----

/** Multiply an #rrggbb colour by a factor (>1 lightens, <1 darkens); passes non-hex through. */
function shade(colour: string, f: number): string {
  const h = colour.replace("#", "");
  if (h.length < 6 || /[^0-9a-fA-F]/.test(h.slice(0, 6))) return colour;
  const n = parseInt(h.slice(0, 6), 16);
  const ch = (v: number): number => Math.max(0, Math.min(255, Math.round(v * f)));
  return `rgb(${ch((n >> 16) & 255)},${ch((n >> 8) & 255)},${ch(n & 255)})`;
}

interface BarEl { x: number; y: number; base: number; width: number; height: number; getProps?: (k: string[], f?: boolean) => { x: number; y: number; base: number; width: number; height: number } }
interface Bar3DChart { ctx: CanvasRenderingContext2D; chartArea: { left: number; right: number; top: number; bottom: number }; data: { datasets: { threeDBar?: string; threeDHoriz?: boolean }[] }; getDatasetMeta: (i: number) => { data: BarEl[] } }
/** Draws each bar/column as an extruded (isometric) 3-D block. The flat bars are made transparent
    by the caller and only serve as scale carriers; this plugin draws the front, top and right faces
    (works for vertical columns and horizontal bars). */
export const bar3DPlugin = {
  id: "sheeteditBar3D",
  afterDatasetsDraw(chart: Bar3DChart): void {
    const dss = chart.data.datasets;
    if (!dss.some((d) => d.threeDBar)) return;
    const { ctx, chartArea } = chart;
    const depth = Math.max(8, Math.min(26, (chartArea.right - chartArea.left) * 0.035));
    const dx = depth, dy = -depth * 0.6;
    const bars: { x0: number; x1: number; top: number; bot: number; c: string }[] = [];
    dss.forEach((ds, i) => {
      if (!ds.threeDBar) return;
      for (const el of chart.getDatasetMeta(i).data) {
        const p = el.getProps ? el.getProps(["x", "y", "base", "width", "height"], true) : el;
        if (ds.threeDHoriz) bars.push({ x0: Math.min(p.base, p.x), x1: Math.max(p.base, p.x), top: p.y - p.height / 2, bot: p.y + p.height / 2, c: ds.threeDBar });
        else bars.push({ x0: p.x - p.width / 2, x1: p.x + p.width / 2, top: Math.min(p.y, p.base), bot: Math.max(p.y, p.base), c: ds.threeDBar });
      }
    });
    ctx.save();
    // Sides + tops first (they recede), then the front faces on top (nearest plane).
    for (const b of bars) {
      ctx.fillStyle = shade(b.c, 1.22);
      ctx.beginPath(); ctx.moveTo(b.x0, b.top); ctx.lineTo(b.x1, b.top); ctx.lineTo(b.x1 + dx, b.top + dy); ctx.lineTo(b.x0 + dx, b.top + dy); ctx.closePath(); ctx.fill();
      ctx.fillStyle = shade(b.c, 0.7);
      ctx.beginPath(); ctx.moveTo(b.x1, b.top); ctx.lineTo(b.x1, b.bot); ctx.lineTo(b.x1 + dx, b.bot + dy); ctx.lineTo(b.x1 + dx, b.top + dy); ctx.closePath(); ctx.fill();
    }
    for (const b of bars) { ctx.fillStyle = b.c; ctx.fillRect(b.x0, b.top, b.x1 - b.x0, b.bot - b.top); }
    ctx.restore();
  },
};

interface PieArc { x: number; y: number; startAngle: number; endAngle: number; outerRadius: number; innerRadius: number }
interface Pie3DChart { ctx: CanvasRenderingContext2D; data: { datasets: { threeDPie?: string[] }[] }; getDatasetMeta: (i: number) => { data: PieArc[] } }
/** Draws a pie / doughnut as a flattened (elliptical) disc with a shaded side wall on its front
    rim, for the pseudo-3D look. The flat arcs are transparent and only supply the geometry. */
export const pie3DPlugin = {
  id: "sheeteditPie3D",
  afterDatasetsDraw(chart: Pie3DChart): void {
    const ds = chart.data.datasets[0];
    if (!ds?.threeDPie) return;
    const arcs = chart.getDatasetMeta(0).data;
    if (!arcs.length) return;
    const { ctx } = chart;
    const cx = arcs[0].x, r = arcs[0].outerRadius, inner = arcs[0].innerRadius || 0;
    const flat = 0.55;
    const depth = Math.max(10, r * 0.22);
    const cy = arcs[0].y - depth / 2;
    const cols = ds.threeDPie;
    const pt = (a: number, rr: number): [number, number] => [cx + rr * Math.cos(a), cy + rr * flat * Math.sin(a)];
    ctx.save();
    // Front rim wall: the portion of each slice's outer arc that faces the viewer (sin > 0 = bottom).
    arcs.forEach((a, j) => {
      const s = a.startAngle, e = a.endAngle;
      const N = Math.max(2, Math.ceil((e - s) / 0.12));
      ctx.fillStyle = shade(cols[j] || "#888", 0.6);
      for (let k = 0; k < N; k++) {
        const a0 = s + (e - s) * k / N, a1 = s + (e - s) * (k + 1) / N;
        if (Math.sin((a0 + a1) / 2) <= 0) continue;
        const [x0, y0] = pt(a0, r), [x1, y1] = pt(a1, r);
        ctx.beginPath(); ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); ctx.lineTo(x1, y1 + depth); ctx.lineTo(x0, y0 + depth); ctx.closePath(); ctx.fill();
      }
    });
    // Top faces (flattened slices).
    arcs.forEach((a, j) => {
      ctx.beginPath();
      if (inner > 0) {
        ctx.ellipse(cx, cy, r, r * flat, 0, a.startAngle, a.endAngle);
        ctx.ellipse(cx, cy, inner, inner * flat, 0, a.endAngle, a.startAngle, true);
      } else {
        ctx.moveTo(cx, cy);
        ctx.ellipse(cx, cy, r, r * flat, 0, a.startAngle, a.endAngle);
      }
      ctx.closePath();
      ctx.fillStyle = cols[j] || "#888";
      ctx.fill();
      ctx.strokeStyle = "rgba(255,255,255,0.7)"; ctx.lineWidth = 1; ctx.stroke();
    });
    ctx.restore();
  },
};

interface LinePt { x: number; y: number }
interface Line3DChart { ctx: CanvasRenderingContext2D; chartArea: { left: number; right: number; top: number; bottom: number }; data: { datasets: { threeDLine?: { colour: string; area: boolean } }[] }; getDatasetMeta: (i: number) => { data: LinePt[]; yScale?: { getPixelForValue: (v: number) => number } } }
/** Draws a 3-D line (an extruded ribbon) or 3-D area (a filled front face with a receding top lip
    and a side wall) for the pseudo-3D look. The flat line/area is transparent (geometry carrier). */
export const line3DPlugin = {
  id: "sheeteditLine3D",
  afterDatasetsDraw(chart: Line3DChart): void {
    const dss = chart.data.datasets;
    if (!dss.some((d) => d.threeDLine)) return;
    const { ctx, chartArea } = chart;
    const depth = Math.max(8, Math.min(24, (chartArea.right - chartArea.left) * 0.03));
    const dx = depth, dy = -depth * 0.6;
    dss.forEach((ds, i) => {
      const cfg = ds.threeDLine;
      if (!cfg) return;
      const meta = chart.getDatasetMeta(i);
      const pts = meta.data.map((el) => ({ x: el.x, y: el.y })).filter((p) => isFinite(p.x) && isFinite(p.y));
      if (pts.length < 2) return;
      const base = Math.min(chartArea.bottom, meta.yScale ? meta.yScale.getPixelForValue(0) : chartArea.bottom);
      const c = cfg.colour;
      ctx.save();
      if (cfg.area) {
        // Front face (area to baseline).
        ctx.fillStyle = c;
        ctx.beginPath(); ctx.moveTo(pts[0]!.x, base); for (const p of pts) ctx.lineTo(p.x, p.y); ctx.lineTo(pts[pts.length - 1]!.x, base); ctx.closePath(); ctx.fill();
        // Right side wall.
        const last = pts[pts.length - 1]!;
        ctx.fillStyle = shade(c, 0.7);
        ctx.beginPath(); ctx.moveTo(last.x, last.y); ctx.lineTo(last.x, base); ctx.lineTo(last.x + dx, base + dy); ctx.lineTo(last.x + dx, last.y + dy); ctx.closePath(); ctx.fill();
      }
      // Receding top ribbon/lip.
      ctx.fillStyle = cfg.area ? shade(c, 1.15) : shade(c, 0.82);
      for (let k = 0; k < pts.length - 1; k++) {
        const p0 = pts[k]!, p1 = pts[k + 1]!;
        ctx.beginPath(); ctx.moveTo(p0.x, p0.y); ctx.lineTo(p1.x, p1.y); ctx.lineTo(p1.x + dx, p1.y + dy); ctx.lineTo(p0.x + dx, p0.y + dy); ctx.closePath(); ctx.fill();
      }
      // The front edge line.
      ctx.strokeStyle = shade(c, 0.7); ctx.lineWidth = 2; ctx.lineJoin = "round";
      ctx.beginPath(); pts.forEach((p, k) => (k ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y))); ctx.stroke();
      ctx.restore();
    });
  },
};
