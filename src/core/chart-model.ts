// The normalized chart model: the single representation the xlsx and ods readers both produce,
// the two writers both consume, the Chart.js renderer maps from, and the create UI builds.
// Deliberately format-agnostic. Anchor offsets are in pixels (a neutral unit); the readers
// convert from EMU (xlsx) / cm (ods) and the writers convert back.

/** A data reference: a live sheet range ("Sheet1!$B$2:$B$10") with cached values as a fallback. */
export interface ChartRef {
  ref?: string;
  cache?: (string | number | null)[];
}

export interface ChartAxis {
  title?: string;
  min?: number;
  max?: number;
}

export interface ChartSeries {
  name?: ChartRef | string;
  values: ChartRef;
  /** scatter / bubble only. */
  xValues?: ChartRef;
  sizes?: ChartRef;
  color?: string;
  /** Combo charts: this series renders as a different kind from the chart's base kind. */
  type?: ChartKind;
  /** Plot this series against a secondary (right-hand) value axis. */
  secondaryAxis?: boolean;
}

/** A two-cell anchor in 1-based grid coordinates; offsets are pixels within the from/to cell. */
export interface ChartAnchor {
  fromCol: number;
  fromRow: number;
  fromColOff: number;
  fromRowOff: number;
  toCol: number;
  toRow: number;
  toColOff: number;
  toRowOff: number;
}

export type ChartKind = "column" | "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "bubble" | "radar";

export interface ChartModel {
  id: string;
  kind: ChartKind;
  stacked?: boolean;
  title?: string;
  legend?: { show: boolean; pos: "top" | "bottom" | "left" | "right" };
  /** Show the value on each data point. */
  dataLabels?: boolean;
  /** Shared category (x) labels; not used for scatter/bubble. */
  categories?: ChartRef;
  series: ChartSeries[];
  axes?: { x?: ChartAxis; y?: ChartAxis };
  anchor: ChartAnchor;
  /** For a chart read from the file and not yet edited: the part paths to preserve verbatim. */
  original?: { partPath?: string; drawingPath?: string; objectDir?: string };
  /** Created or edited in the UI -> written from this model; otherwise preserved untouched. */
  dirty?: boolean;
}

/** EMU (English Metric Units) per pixel at 96 dpi: 914400 / 96. */
export const EMU_PER_PX = 9525;
export const emuToPx = (emu: number): number => Math.round(emu / EMU_PER_PX);
export const pxToEmu = (px: number): number => Math.round(px * EMU_PER_PX);

/** A short, stable-ish id for a chart (no Date/random available here; caller may pass a seed). */
export const chartId = (seed: string | number): string => `chart-${seed}`;

/** The default Chart.js-facing colour palette (used when a series has no explicit colour). */
export const CHART_PALETTE = [
  "#4e79a7", "#f28e2b", "#59a14f", "#e15759", "#76b7b2",
  "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#bab0ac",
];
