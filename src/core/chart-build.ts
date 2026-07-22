import { colToLetters } from "./model";
import type { ChartAnchor, ChartKind, ChartModel } from "./chart-model";

// Build a ChartModel from a selected data range and options (the create UI). Produces live refs
// into the sheet; the first column is the categories / x values and each remaining column a
// series, with the first row optionally supplying series names.

export interface Rect { r1: number; c1: number; r2: number; c2: number }
export interface BuildOpts { firstRowHeader: boolean; firstColLabels: boolean }

const abs = (sheet: string, col: number, r1: number, r2: number): string =>
  `${sheet}!$${colToLetters(col)}$${r1}:$${colToLetters(col)}$${r2}`;
const cell = (sheet: string, col: number, row: number): string => `${sheet}!$${colToLetters(col)}$${row}`;

/** A default anchor placed to the right of the selection, ~8 columns by ~15 rows. */
export function defaultAnchor(rect: Rect): ChartAnchor {
  const fromCol = rect.c2 + 2;
  const fromRow = rect.r1;
  return { fromCol, fromRow, fromColOff: 0, fromRowOff: 0, toCol: fromCol + 8, toRow: fromRow + 15, toColOff: 0, toRowOff: 0 };
}

export function buildChart(sheetName: string, kind: ChartKind, rect: Rect, opts: BuildOpts, id: string, anchor: ChartAnchor): ChartModel {
  const dataR1 = opts.firstRowHeader ? rect.r1 + 1 : rect.r1;
  const firstIsLabels = opts.firstColLabels && rect.c2 > rect.c1;
  const labelCol = firstIsLabels ? rect.c1 : null;
  const valueCols: number[] = [];
  for (let c = labelCol != null ? rect.c1 + 1 : rect.c1; c <= rect.c2; c++) valueCols.push(c);

  const nameOf = (col: number) => (opts.firstRowHeader ? { ref: cell(sheetName, col, rect.r1) } : undefined);

  if (kind === "scatter" || kind === "bubble") {
    // First value column is x, the next is y (bubble: a third is the size); one series.
    const x = valueCols[0];
    const y = valueCols[1] ?? valueCols[0];
    const size = valueCols[2];
    return {
      id, kind, dirty: true,
      legend: { show: false, pos: "bottom" },
      series: [{ name: nameOf(y), values: { ref: abs(sheetName, y, dataR1, rect.r2) }, xValues: x != null ? { ref: abs(sheetName, x, dataR1, rect.r2) } : undefined, sizes: kind === "bubble" && size != null ? { ref: abs(sheetName, size, dataR1, rect.r2) } : undefined }],
      anchor,
    };
  }

  const categories = labelCol != null ? { ref: abs(sheetName, labelCol, dataR1, rect.r2) } : undefined;
  const series = valueCols.map((c) => ({ name: nameOf(c), values: { ref: abs(sheetName, c, dataR1, rect.r2) } }));
  return {
    id, kind, dirty: true,
    legend: { show: series.length > 1 || kind === "pie" || kind === "doughnut", pos: "bottom" },
    categories,
    series,
    anchor,
  };
}
