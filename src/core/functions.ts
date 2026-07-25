import FormulaParser from "fast-formula-parser";
import { asMatrix, asNumber, asScalar, cmp } from "./dynamic-arrays";

// Excel functions fast-formula-parser ships as empty stubs (they throw "not implemented"), supplied
// here so real workbooks recalculate instead of showing a stale cached value. Grouped by family:
// lookup (MATCH / CHOOSE / XLOOKUP / XMATCH), multi-criteria aggregates (SUMIFS / COUNTIFS / ...),
// statistics (MEDIAN / STDEV / PERCENTILE / ...), text (UPPER / SUBSTITUTE / TEXTJOIN / VALUE) and
// SWITCH. Semantics follow Excel for the documented cases; see FUNCTIONS.md for known deviations.

const FE = (FormulaParser as unknown as { FormulaError: { NA: unknown; VALUE: unknown; DIV0: unknown; NUM: unknown; new (n: string): object } }).FormulaError;
const NA = FE.NA, VALUE = FE.VALUE, DIV0 = FE.DIV0, NUM = FE.NUM;

/** Every cell value of an argument, flattened row-major (blanks kept as null). */
function flat(arg: unknown): unknown[] {
  return asMatrix(arg).flat();
}
/** The numbers in an argument (text / blanks / booleans skipped, as Excel does inside a range). */
function nums(arg: unknown): number[] {
  const out: number[] = [];
  for (const v of flat(arg)) if (typeof v === "number" && Number.isFinite(v)) out.push(v);
  return out;
}
const text = (v: unknown): string => (v == null ? "" : String(v));

/** The parser calls functions on its `funsNeedContext` list (CHOOSE, IF, INDEX, OFFSET, ...) with
    its own context as an extra leading argument. Drop it so the real arguments line up. */
function dropContext(args: unknown[]): unknown[] {
  const first = args[0] as { utils?: unknown; retrieveRef?: unknown } | null;
  return first && typeof first === "object" && (first.utils !== undefined || typeof first.retrieveRef === "function") ? args.slice(1) : args;
}

/** Excel wildcard text match (* any run, ? any single char), case-insensitive. */
function wildcardMatch(value: string, pattern: string): boolean {
  const rx = new RegExp("^" + pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
  return rx.test(value);
}

/** Does a cell value satisfy an Excel criteria string (">1", "<>x", "ab*", 5)? */
function matchCriteria(value: unknown, criteria: unknown): boolean {
  if (typeof criteria === "number") return typeof value === "number" && value === criteria;
  const raw = text(criteria).trim();
  const m = /^(<>|>=|<=|=|>|<)(.*)$/.exec(raw);
  const op = m ? m[1]! : "=";
  const operand = m ? m[2]!.trim() : raw;
  const opNum = operand === "" ? NaN : Number(operand);
  const bothNum = typeof value === "number" && Number.isFinite(opNum);
  if (bothNum) {
    const v = value;
    switch (op) {
      case ">": return v > opNum; case ">=": return v >= opNum;
      case "<": return v < opNum; case "<=": return v <= opNum;
      case "<>": return v !== opNum; default: return v === opNum;
    }
  }
  const sv = text(value);
  if (op === "=" || op === "<>") {
    // An empty criteria ("=") matches blank cells; otherwise compare with wildcards.
    const hit = operand === "" ? (value == null || sv === "") : wildcardMatch(sv, operand);
    return op === "<>" ? !hit : hit;
  }
  // A relational operator against text compares case-insensitively.
  const c = sv.toLowerCase() < operand.toLowerCase() ? -1 : sv.toLowerCase() > operand.toLowerCase() ? 1 : 0;
  switch (op) {
    case ">": return c > 0; case ">=": return c >= 0;
    case "<": return c < 0; default: return c <= 0;
  }
}

/** Indices of the criteria-range cells satisfying every (range, criteria) pair. */
function matchingIndices(pairs: unknown[]): number[] | null {
  if (pairs.length < 2) return null;
  const ranges: unknown[][] = [], crits: unknown[] = [];
  for (let i = 0; i + 1 < pairs.length; i += 2) { ranges.push(flat(pairs[i])); crits.push(asScalar(pairs[i + 1])); }
  const n = ranges[0]!.length;
  const out: number[] = [];
  for (let i = 0; i < n; i++) if (ranges.every((r, k) => matchCriteria(r[i], crits[k]))) out.push(i);
  return out;
}

// --- lookup -----------------------------------------------------------------

/** MATCH(lookup, array, [type]) -> 1-based position, or #N/A. */
function MATCH(lookup: unknown, array: unknown, type?: unknown): unknown {
  const v = asScalar(lookup);
  const arr = flat(array);
  const mt = type === undefined ? 1 : asNumber(type, 1);
  if (mt === 0) {
    const isText = typeof v === "string";
    for (let i = 0; i < arr.length; i++) {
      const c = arr[i];
      if (isText ? (typeof c === "string" && wildcardMatch(c, v as string)) : c === v) return i + 1;
    }
    return NA;
  }
  // Ascending (1): the largest value <= lookup. Descending (-1): the smallest value >= lookup.
  let best = -1;
  for (let i = 0; i < arr.length; i++) {
    const c = arr[i];
    if (c == null || c === "") continue;
    const d = cmp(c, v);
    if (mt === 1 ? d <= 0 : d >= 0) best = i; else break;
  }
  return best >= 0 ? best + 1 : NA;
}

/** XMATCH(lookup, array, [matchMode], [searchMode]). matchMode 0 exact, -1 next smaller, 1 next larger. */
function XMATCH(lookup: unknown, array: unknown, matchMode?: unknown, searchMode?: unknown): unknown {
  const v = asScalar(lookup);
  const arr = flat(array);
  const mm = asNumber(matchMode, 0);
  const reverse = asNumber(searchMode, 1) < 0;
  const order = arr.map((_, i) => i);
  if (reverse) order.reverse();
  let bestIdx = -1, bestDiff = Infinity;
  for (const i of order) {
    const c = arr[i];
    if (mm === 2 ? (typeof c === "string" && wildcardMatch(c, text(v))) : c === v) return i + 1;
    if (mm === -1 || mm === 1) {
      if (c == null || c === "") continue;
      const d = cmp(c, v);
      const wanted = mm === -1 ? d <= 0 : d >= 0;
      if (wanted) { const dist = Math.abs(typeof c === "number" && typeof v === "number" ? c - v : d); if (dist < bestDiff) { bestDiff = dist; bestIdx = i; } }
    }
  }
  return bestIdx >= 0 ? bestIdx + 1 : NA;
}

/** XLOOKUP(lookup, lookupArray, returnArray, [ifNotFound], [matchMode], [searchMode]). */
function XLOOKUP(lookup: unknown, lookupArr: unknown, returnArr: unknown, ifNotFound?: unknown, matchMode?: unknown, searchMode?: unknown): unknown {
  const pos = XMATCH(lookup, lookupArr, matchMode, searchMode);
  if (typeof pos !== "number") return ifNotFound === undefined ? NA : asScalar(ifNotFound);
  const m = asMatrix(returnArr);
  // A single-column return array yields one value; a multi-column one yields that whole row.
  if (m.length > 1 && m[0]!.length > 1) return [m[pos - 1] ?? []];
  const flatR = m.flat();
  return flatR[pos - 1] ?? NA;
}

/** CHOOSE(index, v1, v2, ...). */
function CHOOSE(index: unknown, ...vals: unknown[]): unknown {
  const i = Math.trunc(asNumber(index, 0));
  if (i < 1 || i > vals.length) return VALUE;
  const picked = vals[i - 1];
  const m = asMatrix(picked);
  return m.length === 1 && m[0]!.length === 1 ? m[0]![0] : m;
}

// --- statistics -------------------------------------------------------------

function percentileOf(sorted: number[], k: number): unknown {
  if (!sorted.length || !(k >= 0 && k <= 1)) return NUM;
  const idx = k * (sorted.length - 1);
  const lo = Math.floor(idx), hi = Math.ceil(idx);
  return sorted[lo]! + (sorted[hi]! - sorted[lo]!) * (idx - lo);
}
function variance(v: number[], sample: boolean): unknown {
  const n = v.length;
  if (n < (sample ? 2 : 1)) return DIV0;
  const mean = v.reduce((a, b) => a + b, 0) / n;
  return v.reduce((a, b) => a + (b - mean) ** 2, 0) / (sample ? n - 1 : n);
}

// --- text -------------------------------------------------------------------

/** SUBSTITUTE(text, old, new, [instance]) - replaces all occurrences, or only the nth. */
function SUBSTITUTE(t: unknown, oldT: unknown, newT: unknown, instance?: unknown): unknown {
  const s = text(asScalar(t)), o = text(asScalar(oldT)), n = text(asScalar(newT));
  if (o === "") return s;
  if (instance === undefined) return s.split(o).join(n);
  const nth = Math.trunc(asNumber(instance, 1));
  if (nth < 1) return VALUE;
  let idx = -1;
  for (let k = 0; k < nth; k++) { idx = s.indexOf(o, idx + (k === 0 ? 0 : 1)); if (idx < 0) return s; }
  return s.slice(0, idx) + n + s.slice(idx + o.length);
}

/** TEXTJOIN(delimiter, ignoreEmpty, ...items). */
function TEXTJOIN(delim: unknown, ignoreEmpty: unknown, ...items: unknown[]): unknown {
  const d = text(asScalar(delim));
  const skip = asScalar(ignoreEmpty) !== false;
  const parts: string[] = [];
  for (const it of items) for (const v of flat(it)) { const s = v == null ? "" : text(v); if (skip && s === "") continue; parts.push(s); }
  return parts.join(d);
}

/** The aggregate behind SUBTOTAL / AGGREGATE, selected by Excel's function number. */
function subtotalLike(fn: number, ranges: unknown[]): unknown {
  const v = ranges.flatMap(nums);
  const sum = v.reduce((a, b) => a + b, 0);
  switch (Math.trunc(fn)) {
    case 1: return v.length ? sum / v.length : DIV0;            // AVERAGE
    case 2: return v.length;                                     // COUNT
    case 3: return ranges.flatMap(flat).filter((x) => x != null && x !== "").length; // COUNTA
    case 4: return v.length ? Math.max(...v) : 0;                // MAX
    case 5: return v.length ? Math.min(...v) : 0;                // MIN
    case 6: return v.reduce((a, b) => a * b, 1);                 // PRODUCT
    case 7: { const q = variance(v, true); return typeof q === "number" ? Math.sqrt(q) : q; }   // STDEV
    case 8: { const q = variance(v, false); return typeof q === "number" ? Math.sqrt(q) : q; }  // STDEVP
    case 9: return sum;                                          // SUM
    case 10: return variance(v, true);                           // VAR
    case 11: return variance(v, false);                          // VARP
    case 12: { const s = [...v].sort((a, b) => a - b); const m = s.length >> 1; return s.length ? (s.length % 2 ? s[m]! : (s[m - 1]! + s[m]!) / 2) : NUM; } // MEDIAN
    case 13: { const c = new Map<number, number>(); for (const x of v) c.set(x, (c.get(x) ?? 0) + 1); let best: number | null = null, bc = 1; for (const x of v) { const k = c.get(x)!; if (k > bc) { bc = k; best = x; } } return best ?? NA; } // MODE
    default: return VALUE;
  }
}

/** SWITCH(expr, val1, res1, [val2, res2, ...], [default]). */
function SWITCH(expr: unknown, ...rest: unknown[]): unknown {
  const v = asScalar(expr);
  let i = 0;
  for (; i + 1 < rest.length; i += 2) if (asScalar(rest[i]) === v) return asScalar(rest[i + 1]);
  return i < rest.length ? asScalar(rest[i]) : NA; // trailing odd arg = default
}

/** The custom-function map to merge into the parser config. */
export function extraFunctions(): Record<string, (...args: unknown[]) => unknown> {
  const statAgg = (fn: (v: number[]) => unknown) => (...args: unknown[]): unknown => fn(args.flatMap(nums));
  const sortedAgg = (fn: (v: number[], k: number) => unknown) => (a: unknown, k: unknown): unknown => fn(nums(a).sort((x, y) => x - y), asNumber(k, 0));
  // SUMIFS/AVERAGEIFS/MAXIFS/MINIFS take the value range first, then (range, criteria) pairs.
  const overPairs = (fn: (picked: number[]) => unknown) => (valueRange: unknown, ...pairs: unknown[]): unknown => {
    const idx = matchingIndices(pairs);
    if (!idx) return VALUE;
    const vals = flat(valueRange);
    const picked: number[] = [];
    for (const i of idx) { const v = vals[i]; if (typeof v === "number" && Number.isFinite(v)) picked.push(v); }
    return fn(picked);
  };
  return {
    // lookup
    MATCH: (a, b, c) => MATCH(a, b, c),
    XMATCH: (a, b, c, d) => XMATCH(a, b, c, d),
    XLOOKUP: (a, b, c, d, e, f) => XLOOKUP(a, b, c, d, e, f),
    // CHOOSE is on the parser's needs-context list, so its args arrive shifted by one.
    CHOOSE: (...args: unknown[]) => { const [a, ...rest] = dropContext(args); return CHOOSE(a, ...rest); },
    // multi-criteria aggregates
    SUMIFS: overPairs((v) => v.reduce((a, b) => a + b, 0)),
    AVERAGEIFS: overPairs((v) => (v.length ? v.reduce((a, b) => a + b, 0) / v.length : DIV0)),
    MAXIFS: overPairs((v) => (v.length ? Math.max(...v) : 0)),
    MINIFS: overPairs((v) => (v.length ? Math.min(...v) : 0)),
    COUNTIFS: (...pairs: unknown[]) => matchingIndices(pairs)?.length ?? VALUE,
    // statistics
    MEDIAN: statAgg((v) => { if (!v.length) return NUM; const s = [...v].sort((a, b) => a - b); const m = s.length >> 1; return s.length % 2 ? s[m]! : (s[m - 1]! + s[m]!) / 2; }),
    MODE: statAgg((v) => { const c = new Map<number, number>(); for (const n of v) c.set(n, (c.get(n) ?? 0) + 1); let best: number | null = null, bc = 1; for (const n of v) { const k = c.get(n)!; if (k > bc) { bc = k; best = n; } } return best ?? NA; }),
    "MODE.SNGL": statAgg((v) => { const c = new Map<number, number>(); for (const n of v) c.set(n, (c.get(n) ?? 0) + 1); let best: number | null = null, bc = 1; for (const n of v) { const k = c.get(n)!; if (k > bc) { bc = k; best = n; } } return best ?? NA; }),
    STDEV: statAgg((v) => { const q = variance(v, true); return typeof q === "number" ? Math.sqrt(q) : q; }),
    "STDEV.S": statAgg((v) => { const q = variance(v, true); return typeof q === "number" ? Math.sqrt(q) : q; }),
    STDEVP: statAgg((v) => { const q = variance(v, false); return typeof q === "number" ? Math.sqrt(q) : q; }),
    "STDEV.P": statAgg((v) => { const q = variance(v, false); return typeof q === "number" ? Math.sqrt(q) : q; }),
    VAR: statAgg((v) => variance(v, true)), "VAR.S": statAgg((v) => variance(v, true)),
    VARP: statAgg((v) => variance(v, false)), "VAR.P": statAgg((v) => variance(v, false)),
    LARGE: sortedAgg((v, k) => (k >= 1 && k <= v.length ? v[v.length - Math.trunc(k)]! : NUM)),
    SMALL: sortedAgg((v, k) => (k >= 1 && k <= v.length ? v[Math.trunc(k) - 1]! : NUM)),
    PERCENTILE: sortedAgg((v, k) => percentileOf(v, k)),
    "PERCENTILE.INC": sortedAgg((v, k) => percentileOf(v, k)),
    QUARTILE: sortedAgg((v, q) => (q >= 0 && q <= 4 ? percentileOf(v, q / 4) : NUM)),
    "QUARTILE.INC": sortedAgg((v, q) => (q >= 0 && q <= 4 ? percentileOf(v, q / 4) : NUM)),
    RANK: (n: unknown, ref: unknown, order?: unknown) => {
      const v = asNumber(n, NaN), arr = nums(ref);
      if (!Number.isFinite(v) || !arr.includes(v)) return NA;
      const asc = asNumber(order, 0) !== 0;
      return arr.filter((x) => (asc ? x < v : x > v)).length + 1;
    },
    "RANK.EQ": (n: unknown, ref: unknown, order?: unknown) => {
      const v = asNumber(n, NaN), arr = nums(ref);
      if (!Number.isFinite(v) || !arr.includes(v)) return NA;
      const asc = asNumber(order, 0) !== 0;
      return arr.filter((x) => (asc ? x < v : x > v)).length + 1;
    },
    COUNTBLANK: (...args: unknown[]) => args.flatMap(flat).filter((v) => v == null || v === "").length,
    // text
    UPPER: (t: unknown) => text(asScalar(t)).toUpperCase(),
    SUBSTITUTE: (a, b, c, d) => SUBSTITUTE(a, b, c, d),
    TEXTJOIN: (a, b, ...rest) => TEXTJOIN(a, b, ...rest),
    VALUE: (t: unknown) => { const s = text(asScalar(t)).trim().replace(/,/g, ""); if (s === "") return 0; const n = Number(s.endsWith("%") ? s.slice(0, -1) : s); return Number.isFinite(n) ? (s.endsWith("%") ? n / 100 : n) : VALUE; },
    // Excel's SEARCH is case-INsensitive (FIND is the case-sensitive one); the library's is not.
    SEARCH: (needle: unknown, hay: unknown, start?: unknown) => {
      const n = text(asScalar(needle)).toLowerCase(), h = text(asScalar(hay)).toLowerCase();
      const from = Math.max(1, Math.trunc(asNumber(start, 1)));
      if (from > h.length + 1) return VALUE;
      // SEARCH also honours the * and ? wildcards.
      if (/[*?]/.test(n)) {
        const rx = new RegExp(n.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, "."), "i");
        const m = rx.exec(h.slice(from - 1));
        return m ? m.index + from : VALUE;
      }
      const i = h.indexOf(n, from - 1);
      return i < 0 ? VALUE : i + 1;
    },
    // TEXTSPLIT / TEXTBEFORE / TEXTAFTER (Excel 365 text splitting).
    TEXTSPLIT: (t: unknown, colDelim: unknown, rowDelim: unknown, ignoreEmpty?: unknown) => {
      const s = text(asScalar(t));
      const cds = flat(colDelim).map(text).filter((d) => d !== "");
      const rds = rowDelim === undefined ? [] : flat(rowDelim).map(text).filter((d) => d !== "");
      const skip = !!asScalar(ignoreEmpty);
      const split = (str: string, ds: string[]): string[] => {
        if (!ds.length) return [str];
        const rx = new RegExp(ds.map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|"));
        return str.split(new RegExp(rx, "g"));
      };
      const rows = rds.length ? split(s, rds) : [s];
      const out = rows.filter((r) => !skip || r !== "").map((r) => split(r, cds).filter((c) => !skip || c !== ""));
      const w = Math.max(1, ...out.map((r) => r.length));
      return out.map((r) => (r.length === w ? r : [...r, ...Array<unknown>(w - r.length).fill(NA)]));
    },
    TEXTBEFORE: (t: unknown, delim: unknown, instance?: unknown) => {
      const s = text(asScalar(t)), d = text(asScalar(delim));
      if (d === "") return "";
      const nth = Math.trunc(asNumber(instance, 1));
      if (nth < 0) { const i = s.lastIndexOf(d); return i < 0 ? NA : s.slice(0, i); }
      let idx = -1;
      for (let k = 0; k < Math.max(1, nth); k++) { idx = s.indexOf(d, k === 0 ? 0 : idx + d.length); if (idx < 0) return NA; }
      return s.slice(0, idx);
    },
    TEXTAFTER: (t: unknown, delim: unknown, instance?: unknown) => {
      const s = text(asScalar(t)), d = text(asScalar(delim));
      if (d === "") return s;
      const nth = Math.trunc(asNumber(instance, 1));
      if (nth < 0) { const i = s.lastIndexOf(d); return i < 0 ? NA : s.slice(i + d.length); }
      let idx = -1;
      for (let k = 0; k < Math.max(1, nth); k++) { idx = s.indexOf(d, k === 0 ? 0 : idx + d.length); if (idx < 0) return NA; }
      return s.slice(idx + d.length);
    },
    /** LOOKUP(value, vector, [result]) - approximate match over an ascending vector. */
    LOOKUP: (lookup: unknown, vector: unknown, result?: unknown) => {
      const v = asScalar(lookup);
      const src = asMatrix(vector);
      // The array form (no result vector) searches the first column/row and returns the last one.
      const wide = src.length === 1 || (src[0]?.length ?? 0) > src.length;
      const keys = result !== undefined ? flat(vector) : (wide ? src[0]! : src.map((r) => r[0]));
      const outs = result !== undefined ? flat(result) : (wide ? src[src.length - 1]! : src.map((r) => r[r.length - 1]));
      let best = -1;
      for (let i = 0; i < keys.length; i++) { const c = keys[i]; if (c == null || c === "") continue; if (cmp(c, v) <= 0) best = i; else break; }
      return best >= 0 ? outs[best] ?? NA : NA;
    },
    /** SUBTOTAL(fnNum, ...ranges) - 1..11 (and 101..111, which ignore hidden rows we do not model). */
    SUBTOTAL: (fnNum: unknown, ...ranges: unknown[]) => subtotalLike(asNumber(fnNum, 9) % 100, ranges),
    /** AGGREGATE(fnNum, options, ...) - options (ignore errors / hidden) are accepted and ignored. */
    AGGREGATE: (fnNum: unknown, _opts: unknown, ...rest: unknown[]) => {
      const f = asNumber(fnNum, 9);
      // 14..19 take an extra k argument (LARGE / SMALL / PERCENTILE / QUARTILE).
      if (f >= 14 && f <= 19) {
        const v = nums(rest[0]).sort((a, b) => a - b), k = asNumber(rest[1], 1);
        switch (f) {
          case 14: return k >= 1 && k <= v.length ? v[v.length - Math.trunc(k)]! : NUM;
          case 15: return k >= 1 && k <= v.length ? v[Math.trunc(k) - 1]! : NUM;
          case 16: return percentileOf(v, k);
          case 17: return k >= 0 && k <= 4 ? percentileOf(v, k / 4) : NUM;
          default: return NUM;
        }
      }
      return subtotalLike(f, rest);
    },
    // logical
    SWITCH: (a, ...rest) => SWITCH(a, ...rest),
  };
}
