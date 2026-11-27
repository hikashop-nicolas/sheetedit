import { colToLetters, parseA1Ref } from "./model";
import { formulaFromLocal, formulaToLocal } from "./formula-locale";

// Defined names: Excel's naming rules, references to and from a selection, and a token-aware
// scan of formulas so a renamed name follows into every formula and a constant name evaluates.

export type NameError = "empty" | "chars" | "cellRef" | "duplicate";
export type Rect = { r1: number; c1: number; r2: number; c2: number };

const NAME_RE = /^[\p{L}_\\][\p{L}\p{N}_.]*$/u;

/** True when the text would read as a cell reference (A1, XFD1048576, R1C1, R, C). */
export function looksLikeCellRef(name: string): boolean {
  const a1 = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(name);
  if (a1) {
    const p = parseA1Ref(`${a1[1]!.toUpperCase()}${a1[2]}`);
    if (p && p.col <= 16384 && p.row >= 1 && p.row <= 1048576) return true;
  }
  return /^[Rr]\d*([Cc]\d*)?$/.test(name) || /^[Cc]\d*$/.test(name);
}

/** Why a name is refused, as Excel would refuse it, or null when it is fine. */
export function nameError(name: string, existing: Iterable<string>, self?: string): NameError | null {
  if (!name) return "empty";
  if (name.length > 255 || !NAME_RE.test(name) || /^_xlnm\./i.test(name)) return "chars";
  if (looksLikeCellRef(name)) return "cellRef";
  const low = name.toLowerCase();
  for (const n of existing) if (n.toLowerCase() === low && n.toLowerCase() !== self?.toLowerCase()) return "duplicate";
  return null;
}

/** A name's key in the map, matched case-insensitively as Excel does. */
export function findName(names: Map<string, string> | undefined, name: string): string | undefined {
  if (!names) return undefined;
  if (names.has(name)) return name;
  const low = name.toLowerCase();
  for (const k of names.keys()) if (k.toLowerCase() === low) return k;
  return undefined;
}

/** A sheet name as a reference prefix needs it: quoted unless it is a bare identifier. */
export const quoteSheetName = (name: string): string =>
  /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) && !looksLikeCellRef(name) ? name : `'${name.replace(/'/g, "''")}'`;

/** An absolute, sheet-qualified reference to a rectangle ("Sheet1!$A$1:$B$5"). */
export function absoluteRangeRef(sheetName: string, r: Rect): string {
  const a = `$${colToLetters(r.c1)}$${r.r1}`;
  const b = `$${colToLetters(r.c2)}$${r.r2}`;
  return `${quoteSheetName(sheetName)}!${a}${r.r1 === r.r2 && r.c1 === r.c2 ? "" : ":" + b}`;
}

const RANGE_RE = /^(?:('(?:[^']|'')+'|[^!'"(),:;\s]+)!)?(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?$/;

/** A plain cell or range reference ("Sheet1!$A$1:$B$5", "B2"), or null for anything else. */
export function parseRangeRef(ref: string): { sheet?: string; rect: Rect } | null {
  const m = RANGE_RE.exec(ref.trim());
  if (!m) return null;
  const a = parseA1Ref(m[2]!.replace(/\$/g, "").toUpperCase());
  const b = m[3] ? parseA1Ref(m[3].replace(/\$/g, "").toUpperCase()) : a;
  if (!a || !b) return null;
  const sheet = m[1] ? m[1].replace(/^'|'$/g, "").replace(/''/g, "'") : undefined;
  return { sheet, rect: { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) } };
}

/** What a name refers to, as the user reads it: "=Feuil1!$A$1:$B$5", "=0,5", "=SOMME(...)". */
export const refersToLocal = (stored: string): string => "=" + formulaToLocal(stored);

/** A refers-to typed locally, made invariant for the workbook (leading "=" optional). */
export const refersFromLocal = (text: string): string => formulaFromLocal(text.trim().replace(/^=/, "").trim());

const WORD_START = /[\p{L}_\\]/u;
const WORD = /[\p{L}\p{N}_.\\]/u;

/** Rewrite each bare name token (not a function, sheet prefix, string or table item) via fn. */
export function mapNameTokens(f: string, fn: (word: string) => string | null): string {
  let out = "";
  let i = 0;
  while (i < f.length) {
    const ch = f[i]!;
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < f.length) {
        if (f[j] === ch) {
          if (f[j + 1] === ch) j += 2;
          else break;
        } else j++;
      }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "[") {
      let depth = 0;
      let j = i;
      for (; j < f.length; j++) {
        if (f[j] === "'") { j++; continue; }
        if (f[j] === "[") depth++;
        else if (f[j] === "]" && --depth === 0) break;
      }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (/\d/.test(ch)) {
      let j = i + 1;
      while (j < f.length && WORD.test(f[j]!)) j++;
      out += f.slice(i, j);
      i = j;
      continue;
    }
    if (WORD_START.test(ch)) {
      let j = i + 1;
      while (j < f.length && WORD.test(f[j]!)) j++;
      const word = f.slice(i, j);
      const prev = f[i - 1] ?? "", next = f[j] ?? "";
      const bare = next !== "(" && next !== "!" && next !== "[" && prev !== "!" && prev !== "$" && next !== "$" && next !== ":" && prev !== ":";
      const rep = bare ? fn(word) : null;
      out += rep ?? word;
      i = j;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

/** True when the formula refers to this name (case-insensitively). */
export function formulaUsesName(formula: string, name: string): boolean {
  const low = name.toLowerCase();
  if (!formula.toLowerCase().includes(low)) return false;
  let hit = false;
  mapNameTokens(formula, (w) => {
    if (w.toLowerCase() === low) hit = true;
    return null;
  });
  return hit;
}

/** The formula with one name renamed, every occurrence, matched case-insensitively. */
export function renameInFormula(formula: string, from: string, to: string): string {
  const low = from.toLowerCase();
  return mapNameTokens(formula, (w) => (w.toLowerCase() === low ? to : null));
}

/** True when a refers-to is something the engine resolves as a reference (cell or range). */
export const isRangeName = (ref: string): boolean => parseRangeRef(ref) != null;

const valueCache = new WeakMap<Map<string, string>, { size: number; values: Map<string, string> | null }>();

/** Drop the cached value names, which a name pointed elsewhere (same count) would keep stale. */
export const forgetValueNames = (names: Map<string, string> | undefined): void => {
  if (names) valueCache.delete(names);
};
function valueNames(names: Map<string, string>): Map<string, string> | null {
  const hit = valueCache.get(names);
  if (hit && hit.size === names.size) return hit.values;
  let values: Map<string, string> | null = null;
  for (const [k, v] of names) if (!isRangeName(v)) (values ??= new Map()).set(k.toLowerCase(), v);
  valueCache.set(names, { size: names.size, values });
  return values;
}

/**
 * Inline names that hold a constant or a formula ("Rate" = 0.05 -> "(0.05)"), so the engine,
 * which only resolves names to ranges, still evaluates them. Nested names are followed a few levels.
 */
export function expandValueNames(formula: string, names: Map<string, string> | undefined): string {
  if (!names?.size) return formula;
  const values = valueNames(names);
  if (!values) return formula;
  let f = formula;
  for (let depth = 0; depth < 6; depth++) {
    let changed = false;
    f = mapNameTokens(f, (w) => {
      const v = values!.get(w.toLowerCase());
      if (v == null) return null;
      changed = true;
      return `(${v})`;
    });
    if (!changed) break;
  }
  return f;
}
