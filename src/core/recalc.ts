import FormulaParser from "fast-formula-parser";
import type { Cell, CellKind, Sheet, Workbook } from "./model";
import { ensureCell, formatNumber, getCell, numToStr, parseA1Ref, shiftFormula, typedValue } from "./model";
import { isDateFmt } from "./dates";
import { dynamicArrayFunctions } from "./dynamic-arrays";
import { extraFunctions } from "./functions";
import { financialFunctions } from "./financial";
import { allowRawRefs, fixIndexSheet, referenceFunctions } from "./reference-fns";
import { expandLet, hasLet } from "./let-expand";
import { expandTableRefs, hasTableRef, tableBodyRef } from "./table-refs";

type CellRef = { sheet: string; row: number; col: number };
type RangeRef = { sheet: string; from: { row: number; col: number }; to: { row: number; col: number } };

// Parse a defined-name target ("Sheet1!$A$1:$B$2", "A1") into a cell or range reference for
// fast-formula-parser. The $ absolute markers and an optional (quoted) sheet prefix are handled.
function parseNameRef(ref: string, defaultSheet: string): CellRef | RangeRef | null {
  let sheet = defaultSheet;
  let body = ref.trim();
  const bang = body.lastIndexOf("!");
  if (bang >= 0) {
    sheet = body.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'");
    body = body.slice(bang + 1);
  }
  const parts = body.split(":");
  const a = parseA1Ref(parts[0]!.replace(/\$/g, ""));
  if (!a) return null;
  if (parts.length === 1) return { sheet, row: a.row, col: a.col };
  const b = parseA1Ref(parts[1]!.replace(/\$/g, ""));
  if (!b) return null;
  return {
    sheet,
    from: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    to: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}
// ---------------------------------------------------------------------------
// Recalc engine (shared)
// ---------------------------------------------------------------------------

export interface FNode {
  sheet: Sheet;
  cell: Cell;
  id: string;
  deps: Set<string>;
}

// Formulas whose result may be a 2-D array (so it must be requested whole and spilled).
const SPILL_FN_RE = /\b(?:UNIQUE|SORT|SORTBY|FILTER|SEQUENCE|TRANSPOSE|RANDARRAY|MMULT|MUNIT|TAKE|DROP|CHOOSEROWS|CHOOSECOLS|EXPAND|HSTACK|VSTACK|TOROW|TOCOL|WRAPROWS|WRAPCOLS|TEXTSPLIT)\s*\(/i;
// A bare range reference (optionally sheet-qualified), e.g. A1:C3 or Sheet2!$A$1:$B$9.
const BARE_RANGE_RE = /^(?:'[^']+'|[^!]+)?!?\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+$/;

export function applyResult(cell: Cell, res: unknown): void {
  let value: string;
  let kind: CellKind;
  if (res == null) {
    value = "";
    kind = "blank";
  } else if (typeof res === "number") {
    value = Number.isFinite(res) ? numToStr(res) : "#NUM!";
    kind = Number.isFinite(res) ? "n" : "e";
  } else if (typeof res === "boolean") {
    value = res ? "TRUE" : "FALSE";
    kind = "b";
  } else if (Array.isArray(res)) {
    applyResult(cell, (res[0] && (res[0] as unknown[])[0]) ?? "");
    return;
  } else if (typeof res === "object") {
    value = String(res); // FormulaError -> "#DIV/0!" etc.
    kind = "e";
  } else {
    value = String(res);
    kind = "s";
  }
  if (value !== cell.value || kind !== cell.kind) {
    cell.value = value;
    cell.kind = kind;
    cell.recomputed = true;
  }
  // Refresh the formatted display from the (possibly new) value.
  cell.display = kind === "n" && cell.numFmt != null ? formatNumber(cell.numFmt, value) ?? undefined : undefined;
}

// ---------------------------------------------------------------------------
// Linear aggregate overrides. fast-formula-parser's built-in aggregates copy
// arrays per element while coercing arguments, which is quadratic on a range:
// one SUM over a 100k-row column took ~30s and froze the tab. These overrides
// receive the raw range arrays and run in a single pass. Semantics follow
// Excel: inside ranges only numbers count; direct literal booleans and
// numeric strings coerce; errors propagate.
// ---------------------------------------------------------------------------

interface AggState {
  sum: number;
  count: number;
  countA: number;
  min: number;
  max: number;
  err: unknown;
}

function fastAggregates(): Record<string, (...args: unknown[]) => unknown> {
  const FE = (FormulaParser as unknown as { FormulaError: { new (n: string): object; DIV0: unknown } }).FormulaError;
  const isErr = (v: unknown): boolean => v instanceof (FE as unknown as new (...a: never[]) => object);
  const feed = (agg: AggState, n: number) => {
    agg.sum += n;
    agg.count++;
    agg.countA++;
    if (n < agg.min) agg.min = n;
    if (n > agg.max) agg.max = n;
  };
  const walk = (v: unknown, inRange: boolean, agg: AggState): void => {
    if (agg.err || v == null) return;
    if (Array.isArray(v)) {
      for (const x of v) walk(x, true, agg);
      return;
    }
    if (typeof v === "object") {
      if (isErr(v)) {
        agg.err = v;
        return;
      }
      // Range/union wrappers carry their cells in .data; single results in .value.
      const o = v as { data?: unknown; value?: unknown };
      if (o.data !== undefined) walk(o.data, true, agg);
      else if (o.value !== undefined) walk(o.value, inRange, agg);
      return;
    }
    if (typeof v === "number") {
      feed(agg, v);
      return;
    }
    if (!inRange && typeof v === "boolean") {
      feed(agg, v ? 1 : 0);
      return;
    }
    if (!inRange && typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v))) {
      feed(agg, Number(v));
      return;
    }
    if (v !== "") agg.countA++; // non-numeric, non-empty: text or an in-range boolean
  };
  const run = (args: unknown[]): AggState => {
    const agg: AggState = { sum: 0, count: 0, countA: 0, min: Infinity, max: -Infinity, err: null };
    for (const a of args) walk(a, false, agg);
    return agg;
  };
  return {
    SUM: (...args) => {
      const a = run(args);
      return a.err ?? a.sum;
    },
    AVERAGE: (...args) => {
      const a = run(args);
      if (a.err) return a.err;
      return a.count ? a.sum / a.count : FE.DIV0;
    },
    MIN: (...args) => {
      const a = run(args);
      return a.err ?? (a.count ? a.min : 0);
    },
    MAX: (...args) => {
      const a = run(args);
      return a.err ?? (a.count ? a.max : 0);
    },
    COUNT: (...args) => {
      const a = run(args);
      return a.err ?? a.count;
    },
    COUNTA: (...args) => {
      const a = run(args);
      return a.err ?? a.countA;
    },
  };
}

/**
 * Whether opening this workbook needs a calculation pass before anything is shown.
 *
 * True when a formula cell has no cached result: openpyxl and other libraries write the
 * formula and leave the value out, and a file is allowed to ship that way. Without this the
 * grid renders those cells empty, because rendering reads the cached value and only an edit
 * ever triggers the engine.
 *
 * False for the common case, where the producer stored every result, so a large workbook is
 * not recomputed for nothing on the way in.
 */
export function needsCalcOnLoad(wb: Workbook): boolean {
  for (const sheet of wb.sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.formula !== undefined && cell.value === "") return true;
    }
  }
  return false;
}

export interface RecalcOptions {
  /** Leave every formula cell that already carries a result alone, computing only the blanks.
      The load pass uses this: this engine is not Excel, and a disagreement on a formula it
      renders differently would silently replace the file's own numbers with worse ones. */
  keepCached?: boolean;
}

/** Recompute every formula cell's cached value, in dependency order. */
export function recalc(wb: Workbook, opts: RecalcOptions = {}): void {
  const FP = FormulaParser as unknown as {
    new (config: unknown): { parse(f: string, pos: unknown, allowReturnArray?: boolean): unknown };
    DepParser: new (config: unknown) => { parse(f: string, pos: unknown): Array<Record<string, unknown>> };
  };
  const FE = (FormulaParser as unknown as { FormulaError: { REF: unknown } }).FormulaError;
  const byName = new Map<string, Sheet>();
  for (const s of wb.sheets) byName.set(s.name, s);
  const defaultSheet = wb.sheets[0]?.name;
  /** Unknown sheet names are a reference error, not a silent read of sheet 1. */
  const sheetOf = (sheetName: string | undefined): Sheet | null | undefined => {
    if (sheetName == null) return defaultSheet ? byName.get(defaultSheet) : undefined;
    return byName.get(sheetName) ?? null; // null = named but missing
  };
  const lookup = (sheetName: string | undefined, r: number, c: number): Cell | undefined => {
    const sheet = sheetOf(sheetName);
    return sheet ? getCell(sheet, r, c) : undefined;
  };

  const nodes: FNode[] = [];
  const index = new Map<string, FNode>();
  const idOf = (sheetName: string, r: number, c: number) => `${sheetName} ${r}:${c}`;
  for (const sheet of wb.sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.formula == null) continue;
      cell.calcFailed = undefined; // re-diagnosed on every pass
      const node: FNode = { sheet, cell, id: idOf(sheet.name, cell.row, cell.col), deps: new Set() };
      nodes.push(node);
      index.set(node.id, node);
    }
  }
  if (!nodes.length) return;

  // Formula nodes per sheet, sorted by row, so a range dependency is found by binary-searching
  // the row band instead of scanning every formula cell (was O(refs x formulaCells)).
  const bySheet = new Map<string, FNode[]>();
  for (const node of nodes) {
    const arr = bySheet.get(node.sheet.name);
    if (arr) arr.push(node);
    else bySheet.set(node.sheet.name, [node]);
  }
  for (const arr of bySheet.values()) arr.sort((a, b) => a.cell.row - b.cell.row);

  // Resolve a defined name to a cell/range reference (null = unknown name -> #NAME?). A bare table
  // name is the other thing that looks like one: it means that table's data body.
  const nameToRef = (name: string): CellRef | RangeRef | null => {
    const ref = wb.definedNames?.get(name);
    if (ref) return parseNameRef(ref, defaultSheet ?? "");
    const body = tableBodyRef(wb, name);
    return body ? parseNameRef(body, defaultSheet ?? "") : null;
  };

  /** What the parser actually sees: LET expanded, and structured table references turned into the
      plain ranges they name. Both are textual rewrites the engine has no notion of. */
  const prepare = (formula: string, row: number): string => {
    let f = formula;
    if (hasTableRef(f)) f = expandTableRefs(f, wb, row);
    if (hasLet(f)) f = expandLet(f);
    return f;
  };

  const depParser = new FP.DepParser({ onVariable: (name: string) => nameToRef(name) });
  for (const node of nodes) {
    let refs: Array<Record<string, unknown>> = [];
    try {
      const depSrc = prepare(node.cell.formula!, node.cell.row);
      refs = depParser.parse(depSrc, { row: node.cell.row, col: node.cell.col, sheet: node.sheet.name });
    } catch {
      refs = [];
    }
    for (const ref of refs) {
      const sName = (ref.sheet as string) ?? node.sheet.name;
      if (ref.from) {
        const from = ref.from as { row: number; col: number };
        const to = ref.to as { row: number; col: number };
        const arr = bySheet.get(sName);
        if (arr) {
          let lo = 0;
          let hi = arr.length;
          while (lo < hi) {
            const mid = (lo + hi) >> 1;
            if (arr[mid]!.cell.row < from.row) lo = mid + 1;
            else hi = mid;
          }
          for (let i = lo; i < arr.length && arr[i]!.cell.row <= to.row; i++) {
            const col = arr[i]!.cell.col;
            if (col >= from.col && col <= to.col) node.deps.add(arr[i]!.id);
          }
        }
      } else {
        const depId = idOf(sName, ref.row as number, ref.col as number);
        if (index.has(depId)) node.deps.add(depId);
      }
    }
  }

  // Kahn topological sort: dependencies evaluated before dependents.
  const indeg = new Map<string, number>();
  const dependents = new Map<string, string[]>();
  for (const node of nodes) indeg.set(node.id, node.deps.size);
  for (const node of nodes)
    for (const d of node.deps) {
      if (!dependents.has(d)) dependents.set(d, []);
      dependents.get(d)!.push(node.id);
    }
  const queue: string[] = [];
  for (const node of nodes) if ((indeg.get(node.id) ?? 0) === 0) queue.push(node.id);
  const order: FNode[] = [];
  while (queue.length) {
    const id = queue.shift()!;
    order.push(index.get(id)!);
    for (const dep of dependents.get(id) ?? []) {
      const d = (indeg.get(dep) ?? 1) - 1;
      indeg.set(dep, d);
      if (d === 0) queue.push(dep);
    }
  }
  if (order.length < nodes.length) {
    const seen = new Set(order.map((n) => n.id)); // cycles: best-effort single pass
    const leftovers = nodes.filter((n) => !seen.has(n.id));
    // A leftover merely DOWNSTREAM of a cycle computes (from stale inputs) and is
    // not itself circular; the true cycle members also survive a reverse peel
    // (repeatedly removing leftovers that nothing left depends on).
    const remaining = new Set(leftovers.map((n) => n.id));
    let changed = true;
    while (changed) {
      changed = false;
      for (const n of leftovers) {
        if (!remaining.has(n.id)) continue;
        const hasDependent = leftovers.some((m) => remaining.has(m.id) && m.deps.has(n.id));
        if (!hasDependent) {
          remaining.delete(n.id);
          changed = true;
        }
      }
    }
    for (const node of leftovers) {
      if (remaining.has(node.id)) node.cell.calcFailed = "circular"; // surfaced as a badge
      order.push(node);
    }
  }

  // Excel propagates a date format to a formula that reads a date (=A1+1 shows a date, not a
  // serial). Note the first date-formatted cell each formula reads, then inherit it for display.
  let sawDateFmt: string | number | undefined;
  const noteFmt = (c: Cell | undefined) => {
    if (c?.numFmt != null && isDateFmt(c.numFmt)) sawDateFmt ??= c.numFmt;
  };
  const parser = new FP({
    onVariable: (name: string) => nameToRef(name),
    onCell: (ref: { sheet?: string; row: number; col: number }) => {
      if (sheetOf(ref.sheet) === null) return FE.REF;
      const c = lookup(ref.sheet, ref.row, ref.col);
      noteFmt(c);
      return typedValue(c);
    },
    onRange: (ref: { sheet?: string; from: { row: number; col: number }; to: { row: number; col: number } }) => {
      if (sheetOf(ref.sheet) === null) return [[FE.REF]];
      const out: unknown[][] = [];
      for (let r = ref.from.row; r <= ref.to.row; r++) {
        const rowArr: unknown[] = [];
        for (let c = ref.from.col; c <= ref.to.col; c++) {
          const cl = lookup(ref.sheet, r, c);
          noteFmt(cl);
          rowArr.push(typedValue(cl));
        }
        out.push(rowArr);
      }
      return out;
    },
    functions: { ...fastAggregates(), ...dynamicArrayFunctions(), ...extraFunctions(), ...financialFunctions(), ...referenceFunctions() },
  });
  allowRawRefs(parser);
  fixIndexSheet(parser);

  // Dynamic-array spill: a plain formula (no legacy arrayRef) whose result is a 2-D array with
  // more than one cell spills into the anchor + the range below/right of it. A non-empty obstacle
  // in that range yields #SPILL!. dynSpill on the anchor records the last spilled rectangle so a
  // later recalc can clear a stale/shrunk spill before writing the new one.
  const clearPrevSpill = (anchor: Cell, sheet: Sheet): void => {
    const prev = anchor.dynSpill;
    if (!prev) return;
    for (let r = anchor.row; r <= prev.r; r++)
      for (let c = anchor.col; c <= prev.c; c++) {
        if (r === anchor.row && c === anchor.col) continue;
        const t = getCell(sheet, r, c);
        if (t?.spill) {
          t.value = "";
          t.kind = "blank";
          t.display = undefined;
          t.spill = false;
          t.recomputed = true;
        }
      }
    anchor.dynSpill = undefined;
  };
  const spillDynamic = (node: FNode, grid: unknown[][], rows: number, cols: number): void => {
    const sheet = node.sheet;
    const r0 = node.cell.row;
    const c0 = node.cell.col;
    clearPrevSpill(node.cell, sheet);
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        if (i === 0 && j === 0) continue;
        const t = getCell(sheet, r0 + i, c0 + j);
        if (t && !t.spill && (t.value !== "" || t.formula != null)) {
          node.cell.value = "#SPILL!";
          node.cell.kind = "e";
          node.cell.display = undefined;
          node.cell.recomputed = true;
          return;
        }
      }
    applyResult(node.cell, grid[0]?.[0] ?? null);
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        if (i === 0 && j === 0) continue;
        const target = ensureCell(sheet, r0 + i, c0 + j);
        applyResult(target, grid[i]?.[j] ?? null);
        target.spill = true;
      }
    node.cell.dynSpill = { r: r0 + rows - 1, c: c0 + cols - 1 };
  };
  for (const node of order) {
    let res: unknown;
    sawDateFmt = undefined; // reset per formula; onCell/onRange set it while evaluating
    try {
      // Request the whole 2D result only for genuine spill candidates: a legacy array formula, a
      // dynamic-array producer, or a bare range reference. Requesting it for every formula would
      // change fast-formula-parser's error propagation on lone cell refs (a #REF! becomes #VALUE!).
      // LET has no runtime scope in the parser, so it is expanded away textually before parsing.
      const f = prepare(node.cell.formula!, node.cell.row);
      const wantArray = node.cell.arrayRef != null || SPILL_FN_RE.test(f) || BARE_RANGE_RE.test(f);
      res = parser.parse(f, { row: node.cell.row, col: node.cell.col, sheet: node.sheet.name }, wantArray);
    } catch (err) {
      // Unsupported function / parse error: keep the file's cached value, but
      // say so in the grid instead of silently showing a stale number. On the load pass a cell
      // that came with a result is not stale and nobody asked for it to be computed, so it is
      // not flagged: a blank template sheet whose dates are empty would open covered in badges.
      if (opts.keepCached && node.cell.value !== "") continue;
      if (node.cell.calcFailed !== "circular") {
        // The library throws "#ERROR! Function X is not implemented." for unknown names.
        const msg = `${String(err)} ${(err as { message?: string })?.message ?? ""}`;
        node.cell.calcFailed = /#NAME|not implemented/i.test(msg) ? "name" : "eval";
      }
      continue;
    }
    // The load pass fills the blanks and nothing else: a cell that arrived with a result keeps it.
    if (opts.keepCached && node.cell.value !== "") continue;
    // A fresh recompute can error on blank inputs (e.g. DATEDIF on an empty date) even
    // though the file holds a valid cached result; keep that result rather than show an
    // error. A #NAME? is different: the function does not exist, so the stale value can
    // never refresh; badge it.
    const isErr = res != null && typeof res === "object" && !Array.isArray(res);
    if (isErr && node.cell.value !== "" && node.cell.kind !== "e") {
      if (String(res) === "#NAME?" && node.cell.calcFailed !== "circular") node.cell.calcFailed = "name";
      continue;
    }
    // Dynamic-array spill: a plain formula returning a multi-cell 2D array (UNIQUE/SORT/FILTER/
    // SEQUENCE/TRANSPOSE, a range ref, etc.) spills instead of collapsing to its top-left cell.
    if (!node.cell.arrayRef && Array.isArray(res)) {
      const grid = res as unknown[][];
      const rows = grid.length;
      const cols = grid.reduce((m, r) => Math.max(m, Array.isArray(r) ? r.length : 1), 0);
      if (rows * cols > 1) {
        spillDynamic(node, grid, rows, cols);
        continue;
      }
    }
    // A formula that previously spilled now returns a single value: clear its stale spill.
    if (!node.cell.arrayRef && node.cell.dynSpill) clearPrevSpill(node.cell, node.sheet);
    applyResult(node.cell, res);
    // Legacy array formula: spill the 2D result across its ref range so the other cells stay
    // fresh (they are cached literals in the file and would otherwise go stale on recompute).
    if (node.cell.arrayRef && Array.isArray(res)) {
      const rng = parseNameRef(node.cell.arrayRef, node.sheet.name);
      if (rng && "from" in rng) {
        const grid = res as unknown[][];
        for (let i = 0; i <= rng.to.row - rng.from.row; i++)
          for (let j = 0; j <= rng.to.col - rng.from.col; j++) {
            if (i === 0 && j === 0) continue; // top-left already applied
            const target = ensureCell(node.sheet, rng.from.row + i, rng.from.col + j);
            applyResult(target, grid[i]?.[j] ?? null);
            target.spill = true;
          }
      }
    }
    // Inherit a date format for display when the cell has none of its own and read a date
    // (Excel behaviour); display-only, so the file's format is not changed on save.
    if (node.cell.numFmt == null && node.cell.kind === "n" && sawDateFmt) {
      node.cell.display = formatNumber(sawDateFmt, node.cell.value) ?? undefined;
    }
  }
}

// ---------------------------------------------------------------------------
// Reusable one-off evaluator: parse a single formula against the workbook's
// CURRENT cell values (no dependency ordering; CF reads results recalc already
// produced). Used to evaluate conditional-format expression rules and formula
// operands. `at` is the origin the formula was authored for; `shiftFormula`
// re-anchors relative refs per target cell before evaluation.
// ---------------------------------------------------------------------------

export interface FormulaEvaluator {
  /** Evaluate `formula` as if authored at (r0,c0) but applied at (r,c); returns the scalar result. */
  at(formula: string, r0: number, c0: number, r: number, c: number, sheetName: string): unknown;
}

export function makeFormulaEvaluator(wb: Workbook): FormulaEvaluator {
  const FP = FormulaParser as unknown as {
    new (config: unknown): { parse(f: string, pos: unknown, allowReturnArray?: boolean): unknown };
  };
  const FE = (FormulaParser as unknown as { FormulaError: { REF: unknown } }).FormulaError;
  const byName = new Map<string, Sheet>();
  for (const s of wb.sheets) byName.set(s.name, s);
  const defaultSheet = wb.sheets[0]?.name;
  const sheetOf = (name: string | undefined): Sheet | null | undefined =>
    name == null ? (defaultSheet ? byName.get(defaultSheet) : undefined) : byName.get(name) ?? null;
  const lookup = (name: string | undefined, r: number, c: number): Cell | undefined => {
    const sheet = sheetOf(name);
    return sheet ? getCell(sheet, r, c) : undefined;
  };
  const nameToRef = (name: string): CellRef | RangeRef | null => {
    const ref = wb.definedNames?.get(name) ?? tableBodyRef(wb, name);
    return ref ? parseNameRef(ref, defaultSheet ?? "") : null;
  };
  const parser = new FP({
    onVariable: (name: string) => nameToRef(name),
    onCell: (ref: { sheet?: string; row: number; col: number }) => {
      if (sheetOf(ref.sheet) === null) return FE.REF;
      return typedValue(lookup(ref.sheet, ref.row, ref.col));
    },
    onRange: (ref: { sheet?: string; from: { row: number; col: number }; to: { row: number; col: number } }) => {
      if (sheetOf(ref.sheet) === null) return [[FE.REF]];
      const out: unknown[][] = [];
      for (let r = ref.from.row; r <= ref.to.row; r++) {
        const rowArr: unknown[] = [];
        for (let c = ref.from.col; c <= ref.to.col; c++) rowArr.push(typedValue(lookup(ref.sheet, r, c)));
        out.push(rowArr);
      }
      return out;
    },
    functions: { ...fastAggregates(), ...dynamicArrayFunctions(), ...extraFunctions(), ...financialFunctions(), ...referenceFunctions() },
  });
  allowRawRefs(parser);
  fixIndexSheet(parser);
  return {
    at(formula, r0, c0, r, c, sheetName) {
      let f = shiftFormula(formula.replace(/^=/, ""), r - r0, c - c0);
      if (hasTableRef(f)) f = expandTableRefs(f, wb, r);
      if (hasLet(f)) f = expandLet(f);
      try {
        return parser.parse(f, { row: r, col: c, sheet: sheetName }, false);
      } catch {
        return FE.REF;
      }
    },
  };
}

