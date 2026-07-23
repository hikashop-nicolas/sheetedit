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
  /** Number format code for the axis tick labels (e.g. "0.0%", "#,##0"). */
  numFmt?: string;
  /** The category axis is a date axis (c:dateAx). Rendered as a category axis; preserved on write. */
  date?: boolean;
}

/** Data-label content + placement (c:dLbls: showVal/showCatName/showSerName/showPercent + dLblPos). */
export interface ChartDataLabels {
  /** Show the numeric value (showVal). */
  value?: boolean;
  /** Show the category name (showCatName). */
  category?: boolean;
  /** Show the series name (showSerName). */
  seriesName?: boolean;
  /** Show the value as a percentage of its category/total (showPercent). */
  percent?: boolean;
  /** Show the legend key swatch (showLegendKey). */
  legendKey?: boolean;
  /** Label position (c:dLblPos): ctr, inEnd, inBase, outEnd, bestFit, l, r, t, b. */
  position?: string;
}

export interface ChartSeries {
  name?: ChartRef | string;
  values: ChartRef;
  /** scatter / bubble only. */
  xValues?: ChartRef;
  sizes?: ChartRef;
  color?: string;
  /** Per-point colours (pie/doughnut slices, bar bars) from c:dPt; undefined = use the palette. */
  pointColors?: (string | undefined)[];
  /** Smoothed line (c:smooth) for line/scatter series. */
  smooth?: boolean;
  /** Marker (c:marker) on a line/scatter series: symbol name + size (px). */
  marker?: { symbol?: string; size?: number };
  /** Combo charts: this series renders as a different kind from the chart's base kind. */
  type?: ChartKind;
  /** Plot this series against a secondary (right-hand) value axis. */
  secondaryAxis?: boolean;
  /** Per-series data labels (c:dLbls on the series); overrides the chart-level labels. */
  labels?: ChartDataLabels;
  /** Per-point pie/doughnut slice explosion (c:dPt c:explosion), as a % of the radius. */
  explosion?: (number | undefined)[];
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
  /** 100% stacked (percentStacked): each category normalised to 100%. Implies stacked. */
  percent?: boolean;
  /** How empty cells are plotted: "gap" (skip), "zero", or "span" (connect across). */
  blanksAs?: "gap" | "zero" | "span";
  /** Doughnut hole size (% of radius), and bar spacing (gap width %, series overlap %). */
  holeSize?: number;
  gapWidth?: number;
  overlap?: number;
  /** Pie/doughnut first-slice angle (degrees clockwise from top), from c:firstSliceAng. */
  rotation?: number;
  title?: string;
  legend?: { show: boolean; pos: "top" | "bottom" | "left" | "right"; deleted?: number[]; overlay?: boolean };
  /** Show the value on each data point (the simple UI toggle). */
  dataLabels?: boolean;
  /** Chart-level data labels (content + position); richer than the dataLabels toggle. */
  labels?: ChartDataLabels;
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
