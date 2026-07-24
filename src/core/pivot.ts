import type { Sheet } from "./model";
import { getCell, numToStr } from "./model";

// Pivot compute engine (pure, format-agnostic). Given a source range and a spec (which source
// columns are row / column / value fields, and each value's aggregation), it produces the distinct
// sorted items per field, the occurring row-key / column-key tuples, the aggregated values with
// grand totals, the raw records, and a materialised output matrix. Both the xlsx and ODS writers
// consume this single result: the writers differ only in the definition XML they emit around it.
//
// v1 layout scope: any number of nested row fields, at most one column field, one or more value
// fields; no intermediate subtotals (only leaf combinations + grand totals). This mirrors the
// outline/no-subtotal form LibreOffice emits, so the generated files round-trip cleanly.

export type PivotFunc = "sum" | "count" | "countNums" | "average" | "min" | "max";

export interface PivotSpec {
  /** 1-based inclusive source range; the header row is r1. */
  source: { r1: number; c1: number; r2: number; c2: number };
  rows: number[]; // field indices, 0-based within the source columns
  cols: number[]; // at most one in v1
  values: { field: number; func: PivotFunc }[];
}

interface PivotItem { label: string; value: string | number; num: boolean; }
interface PivotFieldInfo { index: number; name: string; items: PivotItem[]; indexOf: Map<string, number>; }
interface PivotRecord { cells: { value: string | number | null; num: boolean }[]; }

export interface PivotOutCell { value: string | number; kind: "s" | "n"; bold?: boolean; }

export interface PivotComputed {
  spec: PivotSpec;
  fields: PivotFieldInfo[]; // one per source column, index-aligned to 0..width-1
  records: PivotRecord[];
  rowKeys: number[][]; // sorted item-index tuples over spec.rows
  colKeys: number[][]; // sorted item-index tuples over spec.cols
  valueLabels: string[];
  /** Aggregated value; a null key means the grand total over that axis. */
  agg(rowKey: number[] | null, colKey: number[] | null, vi: number): number | null;
  matrix: PivotOutCell[][];
  width: number;
  height: number;
  /** Header rows above the body (for the writers' location/firstDataRow). */
  headerRows: number;
  /** Row-header columns to the left of the body (for firstDataCol). */
  headerCols: number;
}

const FUNC_LABEL: Record<PivotFunc, string> = { sum: "Sum", count: "Count", countNums: "Count", average: "Average", min: "Min", max: "Max" };

export function pivotValueLabel(func: PivotFunc, name: string): string {
  return `${FUNC_LABEL[func]} - ${name}`;
}

function cellVal(sheet: Sheet, r: number, c: number): { value: string | number | null; num: boolean } {
  const cell = getCell(sheet, r, c);
  if (!cell || cell.value === "" || cell.kind === "blank") return { value: null, num: false };
  if (cell.kind === "n") { const n = Number(cell.value); return Number.isFinite(n) ? { value: n, num: true } : { value: cell.value, num: false }; }
  return { value: cell.value, num: false };
}

function itemLabel(v: string | number | null): string {
  if (v === null || v === "") return "(empty)";
  return typeof v === "number" ? numToStr(v) : v;
}

// Distinct items of a field, ascending: numbers numerically, strings by locale, numbers before strings.
function collectItems(records: PivotRecord[], field: number): PivotItem[] {
  const seen = new Map<string, PivotItem>();
  for (const rec of records) {
    const cv = rec.cells[field]!;
    const v = cv.value === null ? "(empty)" : cv.value;
    const key = (typeof v === "number" ? "n:" : "s:") + v;
    if (!seen.has(key)) seen.set(key, { label: itemLabel(cv.value), value: v, num: cv.num });
  }
  return Array.from(seen.values()).sort((a, b) => {
    if (a.num && b.num) return (a.value as number) - (b.value as number);
    if (a.num !== b.num) return a.num ? -1 : 1;
    return String(a.value).localeCompare(String(b.value));
  });
}

// Occurring key tuples over the given fields, in sorted item-index order.
function collectKeys(records: PivotRecord[], fieldObjs: PivotFieldInfo[]): number[][] {
  if (!fieldObjs.length) return [[]];
  const seen = new Set<string>();
  const keys: number[][] = [];
  for (const rec of records) {
    const key = fieldObjs.map((f) => {
      const cv = rec.cells[f.index]!;
      const v = cv.value === null ? "(empty)" : cv.value;
      return f.indexOf.get((typeof v === "number" ? "n:" : "s:") + v) ?? 0;
    });
    const s = key.join(",");
    if (!seen.has(s)) { seen.add(s); keys.push(key); }
  }
  keys.sort((a, b) => { for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return a[i]! - b[i]!; return 0; });
  return keys;
}

interface Acc { sum: number; count: number; countNums: number; min: number; max: number; }
const newAcc = (): Acc => ({ sum: 0, count: 0, countNums: 0, min: Infinity, max: -Infinity });
function accumulate(acc: Acc, cv: { value: string | number | null; num: boolean }): void {
  if (cv.value === null) return;
  acc.count++;
  if (cv.num) { const n = cv.value as number; acc.countNums++; acc.sum += n; if (n < acc.min) acc.min = n; if (n > acc.max) acc.max = n; }
}
function finalize(acc: Acc, func: PivotFunc): number | null {
  switch (func) {
    case "count": return acc.count || null;
    case "countNums": return acc.countNums || null;
    case "sum": return acc.countNums ? acc.sum : null;
    case "average": return acc.countNums ? acc.sum / acc.countNums : null;
    case "min": return acc.countNums ? acc.min : null;
    case "max": return acc.countNums ? acc.max : null;
  }
}

export function computePivot(sheet: Sheet, spec: PivotSpec): PivotComputed {
  const { source } = spec;
  const width = source.c2 - source.c1 + 1;
  // Field headers + records.
  const headers: string[] = [];
  for (let c = 0; c < width; c++) { const cv = cellVal(sheet, source.r1, source.c1 + c); headers.push(cv.value === null ? "" : String(cv.value)); }
  const records: PivotRecord[] = [];
  for (let r = source.r1 + 1; r <= source.r2; r++) {
    const cells: PivotRecord["cells"] = [];
    for (let c = 0; c < width; c++) cells.push(cellVal(sheet, r, source.c1 + c));
    records.push({ cells });
  }
  const fields: PivotFieldInfo[] = [];
  for (let c = 0; c < width; c++) {
    const items = collectItems(records, c);
    const indexOf = new Map<string, number>();
    items.forEach((it, i) => indexOf.set((it.num ? "n:" : "s:") + it.value, i));
    fields.push({ index: c, name: headers[c] || `Column${c + 1}`, items, indexOf });
  }
  const rowFieldObjs = spec.rows.map((i) => fields[i]!);
  const colFieldObjs = spec.cols.map((i) => fields[i]!);
  const rowKeys = collectKeys(records, rowFieldObjs);
  const colKeys = collectKeys(records, colFieldObjs);
  const valueLabels = spec.values.map((v) => pivotValueLabel(v.func, fields[v.field]!.name));

  // Aggregate in one pass into buckets keyed by (rowKeyStr, colKeyStr), per value field, plus the
  // per-axis grand accumulators. A record's key is derived the same way collectKeys did.
  const keyOf = (rec: PivotRecord, objs: PivotFieldInfo[]): string =>
    objs.map((f) => { const cv = rec.cells[f.index]!; const v = cv.value === null ? "(empty)" : cv.value; return f.indexOf.get((typeof v === "number" ? "n:" : "s:") + v) ?? 0; }).join(",");
  const cellAcc = new Map<string, Acc[]>();
  const rowGrand = new Map<string, Acc[]>(); // by rowKeyStr, over all cols
  const colGrand = new Map<string, Acc[]>(); // by colKeyStr, over all rows
  const grand: Acc[] = spec.values.map(newAcc);
  const bucket = (m: Map<string, Acc[]>, k: string): Acc[] => { let a = m.get(k); if (!a) { a = spec.values.map(newAcc); m.set(k, a); } return a; };
  for (const rec of records) {
    const rk = keyOf(rec, rowFieldObjs);
    const ck = keyOf(rec, colFieldObjs);
    const cellB = bucket(cellAcc, rk + "|" + ck);
    const rgB = bucket(rowGrand, rk);
    const cgB = bucket(colGrand, ck);
    spec.values.forEach((v, vi) => { const cv = rec.cells[v.field]!; accumulate(cellB[vi]!, cv); accumulate(rgB[vi]!, cv); accumulate(cgB[vi]!, cv); accumulate(grand[vi]!, cv); });
  }
  const agg = (rowKey: number[] | null, colKey: number[] | null, vi: number): number | null => {
    const v = spec.values[vi]!;
    if (rowKey === null && colKey === null) return finalize(grand[vi]!, v.func);
    if (rowKey === null) return finalize(colGrand.get(colKey!.join(","))?.[vi] ?? newAcc(), v.func);
    if (colKey === null) return finalize(rowGrand.get(rowKey.join(","))?.[vi] ?? newAcc(), v.func);
    return finalize(cellAcc.get(rowKey.join(",") + "|" + colKey.join(","))?.[vi] ?? newAcc(), v.func);
  };

  const built = materialize(spec, rowFieldObjs, colFieldObjs, rowKeys, colKeys, valueLabels, agg);
  return { spec, fields, records, rowKeys, colKeys, valueLabels, agg, ...built };
}

// Flat outline layout (row fields nested across the left columns, an optional single column field,
// one or more value fields as the data columns), with a grand-total row and column.
function materialize(
  spec: PivotSpec,
  rowFieldObjs: PivotFieldInfo[],
  colFieldObjs: PivotFieldInfo[],
  rowKeys: number[][],
  colKeys: number[][],
  valueLabels: string[],
  agg: (r: number[] | null, c: number[] | null, vi: number) => number | null,
): { matrix: PivotOutCell[][]; width: number; height: number; headerRows: number; headerCols: number } {
  const R = rowFieldObjs.length;
  const C = colFieldObjs.length; // 0 or 1
  const V = spec.values.length;
  const headerCols = Math.max(1, R);
  const colField = colFieldObjs[0];

  // Data columns: (colKey x value), then a grand-total group (only when there is a column field).
  interface DCol { colKey: number[] | null; vi: number; }
  const dcols: DCol[] = [];
  if (C === 1) {
    for (const ck of colKeys) for (let vi = 0; vi < V; vi++) dcols.push({ colKey: ck, vi });
    for (let vi = 0; vi < V; vi++) dcols.push({ colKey: null, vi }); // Total Result group
  } else {
    for (let vi = 0; vi < V; vi++) dcols.push({ colKey: [], vi });
  }

  // Header band: a value-caption row when there is more than one value; a column-item row when
  // there is a column field. Plus the row that names the row fields.
  const showValueRow = V > 1;
  const showColRow = C === 1;
  const headerRows = 1 + (showColRow ? 1 : 0) + (showValueRow ? 1 : 0);
  const width = headerCols + dcols.length;
  const matrix: PivotOutCell[][] = Array.from({ length: headerRows + rowKeys.length + (R > 0 ? 1 : 0) }, () => Array.from({ length: width }, () => ({ value: "", kind: "s" as const })));

  const S = (r: number, c: number, value: string | number, bold = false, kind: "s" | "n" = typeof value === "number" ? "n" : "s") => { matrix[r]![c] = { value, kind, bold }; };

  // Top caption row: value caption (single value) or the column-field name.
  if (C === 1) S(0, headerCols, colField!.name, true);
  else if (!showValueRow) S(0, headerCols, valueLabels[0]!, true);
  // Column-item header row.
  const colItemRow = 1;
  if (showColRow) {
    dcols.forEach((d, j) => {
      const label = d.colKey === null ? "Total Result" : colField!.items[d.colKey[0]!]!.label;
      // Only label once per colKey group (first value column of the group).
      if (V === 1 || d.vi === 0) S(colItemRow, headerCols + j, label, true);
    });
  }
  // Value-caption row (multiple values): the value label under each column group.
  if (showValueRow) {
    const vrow = headerRows - 1;
    dcols.forEach((d, j) => S(vrow, headerCols + j, valueLabels[d.vi]!, true));
  }
  // Row-field-name row (last header row): the row field names on the left.
  const nameRow = headerRows - 1;
  for (let k = 0; k < R; k++) S(nameRow, k, rowFieldObjs[k]!.name, true);
  if (R === 0) S(nameRow, 0, "Total", true);
  // When there is no column field and a single value, the single data column is captioned already.

  // Body rows.
  let rr = headerRows;
  for (const rk of rowKeys) {
    for (let k = 0; k < R; k++) S(rr, k, rowFieldObjs[k]!.items[rk[k]!]!.label);
    dcols.forEach((d, j) => { const v = agg(rk, d.colKey, d.vi); if (v !== null) S(rr, headerCols + j, v); });
    rr++;
  }
  // Grand total row.
  if (R > 0) {
    S(rr, 0, "Total Result", true);
    dcols.forEach((d, j) => { const v = agg(null, d.colKey, d.vi); if (v !== null) S(rr, headerCols + j, v, true); });
  }
  return { matrix, width, height: matrix.length, headerRows, headerCols };
}
