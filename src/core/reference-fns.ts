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
