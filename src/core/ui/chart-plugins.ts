// Custom Chart.js plugins for chart features Chart.js has no built-in for: regression trendlines
// and error bars. Each is a global plugin that no-ops unless a dataset carries its config
// (dataset.trendline / dataset.errorBars), so it never affects charts that do not use it. Kept
// dependency-free: the regressions and the drawing are computed here.

import type { ChartTrendline } from "../chart-model";

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
interface DrawChart { ctx: CanvasRenderingContext2D; chartArea: { left: number; right: number; top: number; bottom: number }; data: { datasets: { data: unknown[]; trendline?: ChartTrendline; errorBars?: unknown }[] }; getDatasetMeta: (i: number) => { xScale?: Scale; yScale?: Scale; hidden?: boolean } }

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
