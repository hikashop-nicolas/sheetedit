import { MAX_COL, MAX_ROW, colToLetters, lettersToCol } from "./model";
import { looksLikeCellRef, quoteSheetName } from "./names";

// 3D references: Sheet1:Sheet3!A1 means A1 on every sheet from Sheet1 to Sheet3 in tab order.
// fast-formula-parser has no such thing: it read "Sheet1:" as a range operator against a
// defined name and produced a whole-sheet range (17 billion cells). The engine therefore
// rewrites each span into the per-sheet references it stands for before parsing.

/** A sheet prefix found in a formula, with the reference after it when there is one. */
export interface SheetPrefix {
  /** Offsets in the formula: the prefix starts at `start`, the "!" ends at `bang`, the ref at `end`. */
  start: number;
  bang: number;
  end: number;
  first: string;
  /** The second sheet of a span; undefined for a plain Sheet!A1 prefix. */
  last?: string;
  /** The reference after the "!", empty when it is not one the engine recognizes. */
  ref: string;
  /** Upper-cased name of the innermost function call it sits in, if any. */
  fn?: string;
  /** True when the reference is a whole argument of that call (a 3D-capable position). */
  wholeArg: boolean;
}

const NAME = "[\\p{L}_][\\p{L}\\p{N}_.]*";
const QUOTED = "'((?:[^']|'')+)'";
const PART = `(?:${QUOTED}|(${NAME}))`;
const PREFIX_RE = new RegExp(`${PART}(?::${PART})?!`, "uy");
const CELL = "\\$?[A-Za-z]{1,3}\\$?\\d{1,7}";
const REF_RE = new RegExp(
  `(?:${CELL}(?:\\s*:\\s*${CELL})?|\\$?[A-Za-z]{1,3}\\s*:\\s*\\$?[A-Za-z]{1,3}|\\$?\\d{1,7}\\s*:\\s*\\$?\\d{1,7}|#REF!)(?![\\p{L}\\p{N}_(!])`,
  "uy",
);
const BOUNDARY = /[\p{L}\p{N}_.$\]!'#]/u;

/** Functions Excel accepts a 3D reference in; each takes any number of arguments. */
export const SPAN_FUNCTIONS = new Set([
  "SUM", "AVERAGE", "AVERAGEA", "COUNT", "COUNTA", "MAX", "MAXA", "MIN", "MINA", "PRODUCT",
  "STDEV", "STDEVA", "STDEVP", "STDEVPA", "STDEV.S", "STDEV.P",
  "VAR", "VARA", "VARP", "VARPA", "VAR.S", "VAR.P",
]);

/** Cheap pre-check: a formula without both ":" and "!" holds no 3D reference. */
export const mayHaveSheetSpan = (formula: string): boolean => formula.includes(":") && formula.includes("!");

const unq = (s: string): string => s.replace(/''/g, "'");

/** Every sheet prefix in the formula, in order; string literals are skipped. */
export function scanSheetPrefixes(formula: string): SheetPrefix[] {
  const out: SheetPrefix[] = [];
  // Function name per open parenthesis (undefined for a grouping one).
  const stack: (string | undefined)[] = [];
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i]!;
    if (ch === '"') {
      i++;
      while (i < formula.length) {
        if (formula[i] === '"' && formula[i + 1] === '"') i += 2;
        else if (formula[i++] === '"') break;
      }
      continue;
    }
    if (ch === "(") {
      const m = /[A-Za-z0-9_.]+$/.exec(formula.slice(0, i));
      stack.push(m ? m[0].toUpperCase().replace(/^_XLFN\./, "").replace(/^_XLWS\./, "") : undefined);
      i++;
      continue;
    }
    if (ch === ")") {
      stack.pop();
      i++;
      continue;
    }
    if (i > 0 && BOUNDARY.test(formula[i - 1]!)) {
      i++;
      continue;
    }
    PREFIX_RE.lastIndex = i;
    const m = PREFIX_RE.exec(formula);
    if (!m) {
      if (ch === "'") {
        // A quoted name with no "!" after it: skip it whole so its text is not scanned.
        let j = i + 1;
        while (j < formula.length) {
          if (formula[j] === "'" && formula[j + 1] === "'") j += 2;
          else if (formula[j++] === "'") break;
        }
        i = j;
      } else i++;
      continue;
    }
    let first = m[1] != null ? unq(m[1]) : m[2]!;
    let last: string | undefined = m[3] != null ? unq(m[3]) : m[4];
    // An unquoted name shaped like a cell (A1:Sheet3!B2) is a range operator, not a span.
    if (m[2] != null && last != null && looksLikeCellRef(m[2])) {
      i++;
      continue;
    }
    // Excel's own form quotes the whole span: 'Jan 2026:Mar 2026'!A1 (":" is illegal in a name).
    if (last == null && m[1] != null && first.includes(":")) {
      const k = first.indexOf(":");
      last = first.slice(k + 1);
      first = first.slice(0, k);
    }
    const bang = i + m[0].length;
    REF_RE.lastIndex = bang;
    const r = REF_RE.exec(formula);
    const ref = r ? r[0] : "";
    const end = bang + ref.length;
    const before = formula.slice(0, i).trimEnd().slice(-1);
    const after = formula.slice(end).trimStart().charAt(0);
    const fn = stack[stack.length - 1];
    const wholeArg = fn != null && (before === "(" || before === ",") && (after === ")" || after === ",");
    out.push({ start: i, bang, end, first, last, ref, fn, wholeArg });
    i = end > bang ? end : bang;
  }
  return out;
}

/** A span prefix as Excel writes it: Jan:Mar, or 'Jan 2026:Mar 2026' when either needs quotes. */
export function spanPrefix(first: string, last: string): string {
  const bare = quoteSheetName(first) === first && quoteSheetName(last) === last;
  return bare ? `${first}:${last}` : `'${`${first}:${last}`.replace(/'/g, "''")}'`;
}

function rebuild(formula: string, prefixes: SheetPrefix[], fn: (p: SheetPrefix) => string | null): string {
  let out = "";
  let at = 0;
  for (const p of prefixes) {
    const rep = fn(p);
    if (rep == null) continue;
    out += formula.slice(at, p.start) + rep;
    at = p.end;
  }
  return at ? out + formula.slice(at) : formula;
}

/** The sheets a span covers, by position; null when either end is not a sheet. */
export function spanSheets(first: string, last: string, sheetNames: readonly string[]): string[] | null {
  const lower = sheetNames.map((n) => n.toLowerCase());
  const a = lower.indexOf(first.toLowerCase());
  const b = lower.indexOf(last.toLowerCase());
  if (a < 0 || b < 0) return null;
  return sheetNames.slice(Math.min(a, b), Math.max(a, b) + 1);
}

/** How far a sheet's cells reach; a 3D range is cut to it, as nothing past it can count. */
export type SheetExtent = (sheetName: string) => { rows: number; cols: number } | undefined;

/** The ref cut to the sheet's used extent (A:A on a 40-row sheet reads A1:A40). */
function clampRef(ref: string, ext: { rows: number; cols: number }): string {
  const t = ref.replace(/[\s$]/g, "").toUpperCase();
  let m: RegExpExecArray | null;
  let r1: number, c1: number, r2: number, c2: number;
  if ((m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(t))) {
    [c1, r1, c2, r2] = [lettersToCol(m[1]!), Number(m[2]), lettersToCol(m[3]!), Number(m[4])];
  } else if ((m = /^([A-Z]+):([A-Z]+)$/.exec(t))) {
    [c1, r1, c2, r2] = [lettersToCol(m[1]!), 1, lettersToCol(m[2]!), MAX_ROW];
  } else if ((m = /^(\d+):(\d+)$/.exec(t))) {
    [c1, r1, c2, r2] = [1, Number(m[1]), MAX_COL, Number(m[2])];
  } else return ref;
  [r1, r2] = [Math.min(r1, r2), Math.max(r1, r2)];
  [c1, c2] = [Math.min(c1, c2), Math.max(c1, c2)];
  // An empty intersection keeps one blank corner cell, which no 3D function counts.
  r2 = Math.max(r1, Math.min(r2, ext.rows));
  c2 = Math.max(c1, Math.min(c2, ext.cols));
  const a = colToLetters(c1) + r1;
  return r1 === r2 && c1 === c2 ? a : `${a}:${colToLetters(c2)}${r2}`;
}

/**
 * The formula with each 3D reference replaced by the per-sheet references it covers:
 * SUM(Sheet1:Sheet3!A1) evaluates as SUM(Sheet1!A1,Sheet2!A1,Sheet3!A1). A span outside an
 * argument of a 3D-capable function is #VALUE! (as in LibreOffice); a missing end sheet is #REF!.
 */
export function expandSheetSpans(formula: string, sheetNames: readonly string[], extentOf?: SheetExtent): string {
  if (!mayHaveSheetSpan(formula)) return formula;
  const spans = scanSheetPrefixes(formula).filter((p) => p.last != null);
  if (!spans.length) return formula;
  return rebuild(formula, spans, (p) => {
    const sheets = spanSheets(p.first, p.last!, sheetNames);
    if (!sheets || !p.ref || p.ref === "#REF!") return "#REF!";
    if (sheets.length === 1) return `${quoteSheetName(sheets[0]!)}!${p.ref}`;
    if (!p.wholeArg || !SPAN_FUNCTIONS.has(p.fn!)) return "#VALUE!";
    return sheets
      .map((s) => {
        const ext = extentOf?.(s);
        return `${quoteSheetName(s)}!${ext ? clampRef(p.ref, ext) : p.ref}`;
      })
      .join(",");
  });
}

/** Rewrite the 3D spans in Excel's canonical form ('A b':'C d'!A1 becomes 'A b:C d'!A1). */
export function normalizeSheetSpans(formula: string): string {
  if (!mayHaveSheetSpan(formula)) return formula;
  const spans = scanSheetPrefixes(formula).filter((p) => p.last != null);
  return rebuild(formula, spans, (p) => {
    const text = spanPrefix(p.first, p.last!) + "!" + p.ref;
    return text === formula.slice(p.start, p.end) ? null : text;
  });
}

/** The formula with every reference to sheet `from` (plain or a span end) naming `to` instead. */
export function renameSheetInFormula(formula: string, from: string, to: string): string {
  if (!formula.includes("!")) return formula;
  const low = from.toLowerCase();
  const hits = scanSheetPrefixes(formula).filter((p) => p.first.toLowerCase() === low || p.last?.toLowerCase() === low);
  return rebuild(formula, hits, (p) => {
    const a = p.first.toLowerCase() === low ? to : p.first;
    if (p.last == null) return `${quoteSheetName(a)}!${p.ref}`;
    const b = p.last.toLowerCase() === low ? to : p.last;
    return `${spanPrefix(a, b)}!${p.ref}`;
  });
}

/**
 * The formula after the sheet at `index` of `sheetNames` (the order before the delete) goes:
 * a span ending on it shrinks to the next sheet inward, as in Excel, and a span of that sheet
 * alone becomes #REF!.
 */
export function deleteSheetInFormula(formula: string, sheetNames: readonly string[], index: number): string {
  if (!mayHaveSheetSpan(formula)) return formula;
  const lower = sheetNames.map((n) => n.toLowerCase());
  return rebuild(formula, scanSheetPrefixes(formula), (p) => {
    if (p.last == null) return null;
    let a = lower.indexOf(p.first.toLowerCase());
    let b = lower.indexOf(p.last.toLowerCase());
    if (a < 0 || b < 0 || (a !== index && b !== index)) return null;
    if (a > b) [a, b] = [b, a];
    if (a === b) return "#REF!";
    if (a === index) a++;
    else b--;
    const pre = a === b ? quoteSheetName(sheetNames[a]!) : spanPrefix(sheetNames[a]!, sheetNames[b]!);
    return `${pre}!${p.ref}`;
  });
}
