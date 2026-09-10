import FormulaParser from "fast-formula-parser";
import { parseA1Ref } from "./model";

// OFFSET and INDIRECT. Both are *reference-returning* functions: they must hand the parser a
// reference ({ref:{row,col}} or {ref:{from,to}}) rather than a value, so the engine can then read
// the cells behind it (and so a range result can spill or feed an aggregate like SUM).
//
// The parser passes its own context as a leading argument for the functions on its needs-context
// list, and it strips `.ref` off arguments unless the function is also on its no-data-retrieve
// list. OFFSET needs its first argument's reference intact, so registerReferenceFns() adds it to
// that list on the parser instance (see recalc.ts).

const REF = (FormulaParser as unknown as { FormulaError: { REF: unknown } }).FormulaError.REF;

interface Ctx { utils: { extractRefValue(arg: unknown): { val: unknown; isArray: boolean } } }
type CellRef = { row: number; col: number; sheet?: string };
type RangeRef = { from: { row: number; col: number }; to: { row: number; col: number }; sheet?: string };
type RefArg = { ref?: CellRef | RangeRef };

const isRange = (r: CellRef | RangeRef): r is RangeRef => (r as RangeRef).from !== undefined;

/** A raw (non-retrieved) argument's scalar value, via the parser's own extractor. */
function val(ctx: Ctx, arg: unknown): unknown {
  if (arg == null) return undefined;
  const o = arg as { value?: unknown; ref?: unknown };
  if (o.ref !== undefined && ctx?.utils) return ctx.utils.extractRefValue(arg).val;
  if (o.value !== undefined) return o.value;
  return arg;
}
function numOf(ctx: Ctx, arg: unknown, dflt: number): number {
  const v = val(ctx, arg);
  if (v == null || v === "") return dflt;
  const n = Number(Array.isArray(v) ? (v as unknown[][])[0]?.[0] ?? NaN : v);
  return Number.isFinite(n) ? n : dflt;
}

/** OFFSET(reference, rows, cols, [height], [width]) -> a shifted / resized reference. */
function OFFSET(ctx: Ctx, refArg: unknown, rowsA: unknown, colsA: unknown, heightA: unknown, widthA: unknown): unknown {
  const base = (refArg as RefArg)?.ref;
  if (!base) return REF; // an array constant or plain value has no reference to offset
  const from = isRange(base) ? base.from : { row: base.row, col: base.col };
  const to = isRange(base) ? base.to : { row: base.row, col: base.col };
  const baseH = to.row - from.row + 1, baseW = to.col - from.col + 1;
  const dr = Math.trunc(numOf(ctx, rowsA, 0)), dc = Math.trunc(numOf(ctx, colsA, 0));
  const h = heightA == null ? baseH : Math.trunc(numOf(ctx, heightA, baseH));
  const w = widthA == null ? baseW : Math.trunc(numOf(ctx, widthA, baseW));
  if (h <= 0 || w <= 0) return REF;
  const r1 = from.row + dr, c1 = from.col + dc;
  if (r1 < 1 || c1 < 1) return REF; // off the top / left of the sheet
  const sheet = base.sheet;
  if (h === 1 && w === 1) return { ref: { row: r1, col: c1, ...(sheet ? { sheet } : {}) } };
  return { ref: { from: { row: r1, col: c1 }, to: { row: r1 + h - 1, col: c1 + w - 1 }, ...(sheet ? { sheet } : {}) } };
}

/** Parse "Sheet1!A1", "A1:B9" or an R1C1 address into a reference. */
function parseAddress(addr: string, a1Style: boolean): CellRef | RangeRef | null {
  let s = addr.trim();
  if (!s) return null;
  let sheet: string | undefined;
  const bang = s.lastIndexOf("!");
  if (bang >= 0) { sheet = s.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'"); s = s.slice(bang + 1); }
  const body = s.replace(/\$/g, "");
  const mk = (r: number, c: number): CellRef => ({ row: r, col: c, ...(sheet ? { sheet } : {}) });
  const parseOne = (part: string): { row: number; col: number } | null => {
    if (a1Style) { const p = parseA1Ref(part); return p ? { row: p.row, col: p.col } : null; }
    const m = /^R(\d+)C(\d+)$/i.exec(part.trim());
    return m ? { row: Number(m[1]), col: Number(m[2]) } : null;
  };
  const [a, b] = body.split(":");
  const p1 = parseOne(a ?? "");
  if (!p1) return null;
  if (b === undefined) return mk(p1.row, p1.col);
  const p2 = parseOne(b);
  if (!p2) return null;
  return {
    from: { row: Math.min(p1.row, p2.row), col: Math.min(p1.col, p2.col) },
    to: { row: Math.max(p1.row, p2.row), col: Math.max(p1.col, p2.col) },
    ...(sheet ? { sheet } : {}),
  };
}

/** INDIRECT(refText, [a1]) -> the reference the text names. */
function INDIRECT(ctx: Ctx, textArg: unknown, a1Arg: unknown): unknown {
  const raw = val(ctx, textArg);
  const s = raw == null ? "" : String(Array.isArray(raw) ? (raw as unknown[][])[0]?.[0] ?? "" : raw);
  const a1 = a1Arg == null ? true : val(ctx, a1Arg) !== false;
  const ref = parseAddress(s, a1);
  return ref ? { ref } : REF;
}

/** The reference-returning function map (registered like any other custom function). */
export function referenceFunctions(): Record<string, (...args: unknown[]) => unknown> {
  return {
    OFFSET: (ctx, r, rows, cols, h, w) => OFFSET(ctx as Ctx, r, rows, cols, h, w),
    INDIRECT: (ctx, t, a1) => INDIRECT(ctx as Ctx, t, a1),
  };
}

/** OFFSET needs its first argument's reference, which the parser only preserves for functions on
    its no-data-retrieve list. Call this on each freshly built parser instance. */
export function allowRawRefs(parser: unknown): void {
  const p = parser as { funsNeedContextAndNoDataRetrieve?: string[] };
  if (Array.isArray(p.funsNeedContextAndNoDataRetrieve) && !p.funsNeedContextAndNoDataRetrieve.includes("OFFSET")) {
    p.funsNeedContextAndNoDataRetrieve.push("OFFSET");
  }
}

/** The sheet an INDEX source argument belongs to: a cell/range reference, or the chosen area of a
    multi-area reference like INDEX((A1:A9,Other!C1:C9), n, 1, 2). */
function sourceSheet(src: unknown, areaNum: unknown): string | undefined {
  const areas = (src as { refs?: RefArg[] })?.refs;
  if (Array.isArray(areas)) {
    const n = Math.trunc(Number((areaNum as { value?: unknown })?.value ?? areaNum ?? 1));
    return areas[(Number.isFinite(n) && n > 0 ? n : 1) - 1]?.ref?.sheet;
  }
  return (src as RefArg)?.ref?.sheet;
}

/**
 * fast-formula-parser's INDEX builds its result reference out of the source range's row and
 * column and drops the sheet, so `INDEX(Other!$D$5:$D$35, n)` is read back off whichever sheet
 * the formula lives on. Every INDEX+MATCH lookup into another sheet silently returns the wrong
 * cell. Wrap the built-in and put the sheet back.
 */
export function fixIndexSheet(parser: unknown): void {
  const p = parser as { functions?: Record<string, (...args: unknown[]) => unknown> };
  const orig = p.functions?.INDEX;
  if (typeof orig !== "function" || (orig as { sheetFixed?: boolean }).sheetFixed) return;
  const fixed = (...args: unknown[]): unknown => {
    const out = orig(...args);
    const ref = (out as RefArg)?.ref;
    if (ref && ref.sheet == null) {
      const sheet = sourceSheet(args[1], args[4]); // (context, ranges, rowNum, colNum, areaNum)
      if (sheet) ref.sheet = sheet;
    }
    return out;
  };
  fixed.sheetFixed = true;
  p.functions!.INDEX = fixed;
}

const COMPARISONS = new Set(["=", "<>", "<", ">", "<=", ">="]);

/**
 * Excel's blank-cell comparison. An empty cell takes the TYPE of whatever it is compared against,
 * so `A1=""` and `A1=0` are both TRUE when A1 is empty. fast-formula-parser turns a blank into 0
 * unconditionally, so `A1=""` compared a number with a string and came out FALSE. That breaks
 * `IF(A1="","",...)`, the standard way to leave a row blank until it is filled in: instead of
 * nothing, the cell showed the result of the arithmetic in the other branch.
 */
export function fixBlankCompare(parser: unknown): void {
  const utils = (parser as { utils?: Record<string, unknown> }).utils;
  const orig = utils?.["_applyInfix"];
  if (typeof orig !== "function" || (orig as { blankFixed?: boolean }).blankFixed) return;
  type Operand = { val?: unknown; isArray?: boolean };
  const asType = (blank: Operand, other: Operand): Operand => {
    const t = typeof other.val;
    if (t === "string") return { ...blank, val: "" };
    if (t === "boolean") return { ...blank, val: false };
    return blank; // against a number (or another blank), 0 is already right
  };
  const fixed = function (this: unknown, a: Operand, infix: string, b: Operand): unknown {
    if (COMPARISONS.has(infix) && !a?.isArray && !b?.isArray) {
      if (a?.val == null && b?.val != null) a = asType(a, b);
      else if (b?.val == null && a?.val != null) b = asType(b, a);
    }
    return (orig as (a: Operand, i: string, b: Operand) => unknown).call(this, a, infix, b);
  };
  fixed.blankFixed = true;
  utils!["_applyInfix"] = fixed;
}
