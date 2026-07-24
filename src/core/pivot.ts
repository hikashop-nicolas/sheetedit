import type { Sheet } from "./model";
import { getCell, numToStr } from "./model";

// Pivot compute engine (pure, format-agnostic). Given a source range and a spec (which source
// columns are row / column / value / page fields, each value's aggregation, and whether subtotals
// are shown), it produces the distinct sorted items per field, the row/column axes (leaf +
// subtotal + grand lines), prefix-aware aggregation, and a materialised output matrix. Both the
// xlsx and ODS writers consume this one result; they differ only in the definition XML they emit.
//
// Layout scope: any number of nested row and column fields, one or more value fields, optional
// report/page filters, optional subtotals. Grand totals always. Values are laid out as adjacent
// data columns (both LibreOffice and Excel re-flow the body from the source on open anyway).

export type PivotFunc = "sum" | "count" | "countNums" | "average" | "min" | "max";

export interface PivotSpec {
  /** 1-based inclusive source range; the header row is r1. */
  source: { r1: number; c1: number; r2: number; c2: number };
  rows: number[]; // field indices, 0-based within the source columns
  cols: number[];
  values: { field: number; func: PivotFunc }[];
  /** Report/page filters: a field restricted to one item (null = All). */
  pages?: { field: number; item: number | null }[];
  /** Show per-group subtotals for the outer nested fields (Excel default when nested). */
  subtotals?: boolean;
}

export interface PivotItem { label: string; value: string | number; num: boolean; }
interface PivotFieldInfo { index: number; name: string; items: PivotItem[]; indexOf: Map<string, number>; }
interface PivotRecord { cells: { value: string | number | null; num: boolean }[]; }

export interface PivotOutCell { value: string | number; kind: "s" | "n"; bold?: boolean; }

/** One line of a row/column axis: a leaf data line, a per-group subtotal, or the grand total.
    `key` holds the item indices down to this line's depth (empty for grand). */
export interface AxisNode { key: number[]; kind: "leaf" | "subtotal" | "grand"; }

export interface PivotComputed {
  spec: PivotSpec;
  fields: PivotFieldInfo[]; // one per source column, index-aligned to 0..width-1
  records: PivotRecord[]; // ALL source rows (the cache mirrors the whole source, unfiltered)
  rowKeys: number[][]; // sorted leaf item-index tuples over spec.rows (after page filtering)
  colKeys: number[][];
  rowAxis: AxisNode[];
  colAxis: AxisNode[];
  valueLabels: string[];
  /** Items available for each page field (from all data, so the picker lists every value). */
  pageItems: { field: number; items: PivotItem[] }[];
  /** Aggregated value; keys are prefixes (partial = a subtotal), null = grand over that axis. */
  agg(rowKey: number[] | null, colKey: number[] | null, vi: number): number | null;
  matrix: PivotOutCell[][];
  width: number;
  height: number;
  headerRows: number; // header rows above the body (drives firstDataRow)
  headerCols: number; // row-header columns to the left (drives firstDataCol)
}

const FUNC_LABEL: Record<PivotFunc, string> = { sum: "Sum", count: "Count", countNums: "Count", average: "Average", min: "Min", max: "Max" };

export function pivotValueLabel(func: PivotFunc, name: string): string {
  return `${FUNC_LABEL[func]} - ${name}`;
}

/** The distinct sorted items of one source column (same order the engine indexes them), so a page
    filter picker can list every value and map a selection straight to an item index. */
export function pivotColumnItems(sheet: Sheet, source: PivotSpec["source"], col: number): PivotItem[] {
  const records: PivotRecord[] = [];
  for (let r = source.r1 + 1; r <= source.r2; r++) records.push({ cells: [cellVal(sheet, r, source.c1 + col)] });
  return collectItems(records, 0);
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

const itemKey = (v: string | number): string => (typeof v === "number" ? "n:" : "s:") + v;
const recItem = (cv: { value: string | number | null }): string | number => (cv.value === null ? "(empty)" : cv.value);

function collectItems(records: PivotRecord[], field: number): PivotItem[] {
  const seen = new Map<string, PivotItem>();
  for (const rec of records) {
    const cv = rec.cells[field]!;
    const v = recItem(cv);
    if (!seen.has(itemKey(v))) seen.set(itemKey(v), { label: itemLabel(cv.value), value: v, num: cv.num });
  }
  return Array.from(seen.values()).sort((a, b) => {
    if (a.num && b.num) return (a.value as number) - (b.value as number);
    if (a.num !== b.num) return a.num ? -1 : 1;
    return String(a.value).localeCompare(String(b.value));
  });
}

function collectKeys(records: PivotRecord[], fieldObjs: PivotFieldInfo[]): number[][] {
  if (!fieldObjs.length) return [[]];
  const seen = new Set<string>();
  const keys: number[][] = [];
  for (const rec of records) {
    const key = fieldObjs.map((f) => f.indexOf.get(itemKey(recItem(rec.cells[f.index]!))) ?? 0);
    const s = key.join(",");
    if (!seen.has(s)) { seen.add(s); keys.push(key); }
  }
  keys.sort((a, b) => { for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return a[i]! - b[i]!; return 0; });
  return keys;
}

// Build an axis (ordered lines) from the leaf key tuples: each leaf, per-group subtotals for the
// outer levels when enabled, and a trailing grand line when the axis has any fields.
function buildAxis(keys: number[][], depth: number, subtotals: boolean): AxisNode[] {
  const out: AxisNode[] = [];
  for (let i = 0; i < keys.length; i++) {
    out.push({ key: keys[i]!, kind: "leaf" });
    if (subtotals && depth >= 2) {
      const next = keys[i + 1];
      for (let L = depth - 2; L >= 0; L--) {
        const ends = !next || keys[i]!.slice(0, L + 1).join(",") !== next.slice(0, L + 1).join(",");
        if (ends) out.push({ key: keys[i]!.slice(0, L + 1), kind: "subtotal" });
      }
    }
  }
  if (depth >= 1) out.push({ key: [], kind: "grand" });
  return out;
}

interface Acc { sum: number; count: number; countNums: number; min: number; max: number; }
const newAcc = (): Acc => ({ sum: 0, count: 0, countNums: 0, min: Infinity, max: -Infinity });
function accumulate(acc: Acc, cv: { value: string | number | null; num: boolean }): void {
  if (cv.value === null) return;
  acc.count++;
  if (cv.num) { const n = cv.value as number; acc.countNums++; acc.sum += n; if (n < acc.min) acc.min = n; if (n > acc.max) acc.max = n; }
}
function merge(into: Acc, from: Acc): void {
  into.sum += from.sum; into.count += from.count; into.countNums += from.countNums;
  if (from.min < into.min) into.min = from.min; if (from.max > into.max) into.max = from.max;
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

const prefixOf = (key: number[], prefix: number[]): boolean => { for (let i = 0; i < prefix.length; i++) if (key[i] !== prefix[i]) return false; return true; };

export function computePivot(sheet: Sheet, spec: PivotSpec): PivotComputed {
  const { source } = spec;
  const width = source.c2 - source.c1 + 1;
  const headers: string[] = [];
  for (let c = 0; c < width; c++) { const cv = cellVal(sheet, source.r1, source.c1 + c); headers.push(cv.value === null ? "" : String(cv.value)); }
  const allRecords: PivotRecord[] = [];
  for (let r = source.r1 + 1; r <= source.r2; r++) {
    const cells: PivotRecord["cells"] = [];
    for (let c = 0; c < width; c++) cells.push(cellVal(sheet, r, source.c1 + c));
    allRecords.push({ cells });
  }
  // Fields (items from ALL rows so page pickers list every value); indexOf keyed by item.
  const fields: PivotFieldInfo[] = [];
  for (let c = 0; c < width; c++) {
    const items = collectItems(allRecords, c);
    const indexOf = new Map<string, number>();
    items.forEach((it, i) => indexOf.set(itemKey(it.value), i));
    fields.push({ index: c, name: headers[c] || `Column${c + 1}`, items, indexOf });
  }
  // Apply page/report filters (aggregation only; the cache still mirrors all source rows).
  const pages = spec.pages ?? [];
  const filtered = allRecords.filter((rec) => pages.every((p) => p.item == null || fields[p.field]!.indexOf.get(itemKey(recItem(rec.cells[p.field]!))) === p.item));

  const rowFieldObjs = spec.rows.map((i) => fields[i]!);
  const colFieldObjs = spec.cols.map((i) => fields[i]!);
  const rowKeys = collectKeys(filtered, rowFieldObjs);
  const colKeys = collectKeys(filtered, colFieldObjs);
  const subtotals = !!spec.subtotals;
  const rowAxis = buildAxis(rowKeys, spec.rows.length, subtotals);
  const colAxis = buildAxis(colKeys, spec.cols.length, subtotals);
  const valueLabels = spec.values.map((v) => pivotValueLabel(v.func, fields[v.field]!.name));
  const pageItems = pages.map((p) => ({ field: p.field, items: fields[p.field]!.items }));

  // Per full (rowKey,colKey) accumulators; agg merges the subset a prefix / grand covers.
  const keyOf = (rec: PivotRecord, objs: PivotFieldInfo[]): string => objs.map((f) => fields[f.index]!.indexOf.get(itemKey(recItem(rec.cells[f.index]!))) ?? 0).join(",");
  const cellAcc = new Map<string, Acc[]>();
  const bucket = (k: string): Acc[] => { let a = cellAcc.get(k); if (!a) { a = spec.values.map(newAcc); cellAcc.set(k, a); } return a; };
  for (const rec of filtered) {
    const accs = bucket(keyOf(rec, rowFieldObjs) + "|" + keyOf(rec, colFieldObjs));
    spec.values.forEach((v, vi) => accumulate(accs[vi]!, rec.cells[v.field]!));
  }
  const agg = (rowKey: number[] | null, colKey: number[] | null, vi: number): number | null => {
    const acc = newAcc();
    const rks = rowKey === null ? rowKeys : rowKeys.filter((k) => prefixOf(k, rowKey));
    const cks = colKey === null ? colKeys : colKeys.filter((k) => prefixOf(k, colKey));
    for (const rk of rks) for (const ck of cks) { const a = cellAcc.get(rk.join(",") + "|" + ck.join(",")); if (a) merge(acc, a[vi]!); }
    return finalize(acc, spec.values[vi]!.func);
  };

  const built = materialize(spec, rowFieldObjs, colFieldObjs, rowAxis, colAxis, valueLabels, agg);
  return { spec, fields, records: allRecords, rowKeys, colKeys, rowAxis, colAxis, valueLabels, pageItems, agg, ...built };
}

// Flat layout: row fields nested across the left columns, column fields nested across the top header
// rows, one or more value fields as adjacent data columns, with subtotal and grand lines from the
// axes. `agg` handles prefix (subtotal) and null (grand) keys.
function materialize(
  spec: PivotSpec,
  rowFieldObjs: PivotFieldInfo[],
  colFieldObjs: PivotFieldInfo[],
  rowAxis: AxisNode[],
  colAxis: AxisNode[],
  valueLabels: string[],
  agg: (r: number[] | null, c: number[] | null, vi: number) => number | null,
): { matrix: PivotOutCell[][]; width: number; height: number; headerRows: number; headerCols: number } {
  const R = rowFieldObjs.length;
  const C = colFieldObjs.length;
  const V = spec.values.length;
  const headerCols = Math.max(1, R);

  // Data columns: each column-axis line (leaf/subtotal/grand) times each value field.
  interface DCol { node: AxisNode; vi: number; }
  const dcols: DCol[] = [];
  for (const node of colAxis) for (let vi = 0; vi < V; vi++) dcols.push({ node, vi });

  const headerRows = V > 1 ? C + 1 : Math.max(1, C);
  const nameRow = headerRows - 1;
  const width = headerCols + dcols.length;
  const matrix: PivotOutCell[][] = Array.from({ length: headerRows + rowAxis.length }, () => Array.from({ length: width }, () => ({ value: "", kind: "s" as const })));
  const S = (r: number, c: number, value: string | number, bold = false, kind: "s" | "n" = typeof value === "number" ? "n" : "s") => { matrix[r]![c] = { value, kind, bold }; };

  const colKeyOf = (n: AxisNode): number[] | null => (n.kind === "grand" ? null : n.key);
  const colItemLabel = (n: AxisNode, level: number): string => {
    if (n.kind === "grand") return level === 0 ? "Grand Total" : "";
    if (level >= n.key.length) return "";
    const base = colFieldObjs[level]!.items[n.key[level]!]!.label;
    return n.kind === "subtotal" && level === n.key.length - 1 ? `${base} Total` : base;
  };
  // Column-item header rows (one per column field level); a label is written once per group.
  for (let L = 0; L < C; L++) dcols.forEach((d, j) => { const lbl = colItemLabel(d.node, L); if (lbl && (V === 1 || d.vi === 0)) S(L, headerCols + j, lbl, true); });
  // Value captions: their own cells when there are several values, else the single value caption
  // sits over the lone data column when there is no column field.
  if (V > 1) dcols.forEach((d, j) => S(nameRow, headerCols + j, valueLabels[d.vi]!, true));
  else if (C === 0) S(nameRow, headerCols, valueLabels[0]!, true);
  // Row-field names in the left band of the last header row.
  for (let k = 0; k < R; k++) S(nameRow, k, rowFieldObjs[k]!.name, true);
  if (R === 0) S(nameRow, 0, "Total", true);

  // Body rows from the row axis (leaf / subtotal / grand).
  let rr = headerRows;
  for (const rn of rowAxis) {
    const rowKey = rn.kind === "grand" ? null : rn.key;
    const bold = rn.kind !== "leaf";
    if (rn.kind === "grand") S(rr, 0, "Grand Total", true);
    else if (rn.kind === "subtotal") S(rr, rn.key.length - 1, `${rowFieldObjs[rn.key.length - 1]!.items[rn.key[rn.key.length - 1]!]!.label} Total`, true);
    else for (let k = 0; k < R; k++) S(rr, k, rowFieldObjs[k]!.items[rn.key[k]!]!.label);
    dcols.forEach((d, j) => { const v = agg(rowKey, colKeyOf(d.node), d.vi); if (v !== null) S(rr, headerCols + j, v, bold); });
    rr++;
  }
  return { matrix, width, height: matrix.length, headerRows, headerCols };
}
