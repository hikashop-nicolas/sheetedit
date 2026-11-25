// Formulas in the user's language: "SOMME(A1;1,5)" on screen, "SUM(A1,1.5)" in the workbook.
// A small tokenizer swaps function names, booleans, error literals, decimal and argument
// separators, leaving strings, quoted sheet names and structured-reference items untouched.

import { isInvariant, parseLocalNumber, regional, type Regional } from "./regional";

const INV_ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"];

interface Tables {
  fn: Map<string, string>;
  errors: [string, string][];
  bool?: [string, string];
}

const cache = new Map<string, { to: Tables; from: Tables }>();

function tables(r: Regional): { to: Tables; from: Tables } {
  let t = cache.get(r.code);
  if (t) return t;
  const to = new Map<string, string>(), from = new Map<string, string>();
  for (const [en, local] of Object.entries(r.fn)) {
    to.set(en, local);
    from.set(local.toUpperCase(), en);
  }
  const errTo = INV_ERRORS.map((e) => [e, r.errors[e] ?? e] as [string, string]);
  const errFrom = errTo.map(([a, b]) => [b, a] as [string, string]).sort((a, b) => b[0].length - a[0].length);
  t = {
    to: { fn: to, errors: errTo, bool: r.bool },
    from: { fn: from, errors: errFrom, bool: r.bool ? ["TRUE", "FALSE"] : undefined },
  };
  cache.set(r.code, t);
  return t;
}

const WORD_START = /[\p{L}_\\]/u;
const WORD = /[\p{L}\p{N}_.]/u;

function translate(f: string, r: Regional, toLocal: boolean): string {
  const tb = toLocal ? tables(r).to : tables(r).from;
  const srcDec = toLocal ? "." : r.decimal, dstDec = toLocal ? r.decimal : ".";
  const srcArg = toLocal ? "," : r.argSep, dstArg = toLocal ? r.argSep : ",";
  const srcCol = toLocal ? "," : r.arrayCol, dstCol = toLocal ? r.arrayCol : ",";
  const srcRow = toLocal ? ";" : r.arrayRow, dstRow = toLocal ? r.arrayRow : ";";
  const numRe = new RegExp(`^(?:\\d+(?:${srcDec === "." ? "\\." : srcDec}\\d*)?|${srcDec === "." ? "\\." : srcDec}\\d+)(?:[eE][-+]?\\d+)?`);
  const boolSrc = toLocal ? ["TRUE", "FALSE"] : r.bool, boolDst = toLocal ? r.bool : ["TRUE", "FALSE"];
  let out = "";
  let brace = 0;
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
      // Structured reference: only the separator between its items changes.
      let depth = 0;
      let j = i;
      for (; j < f.length; j++) {
        const c = f[j]!;
        if (c === "'") { out += c + (f[j + 1] ?? ""); j++; continue; }
        if (c === "[") depth++;
        else if (c === "]") depth--;
        out += depth === 1 && c === srcArg ? dstArg : c;
        if (depth === 0) break;
      }
      i = j + 1;
      continue;
    }
    if (ch === "{") brace++;
    else if (ch === "}") brace = Math.max(0, brace - 1);
    if (ch === "#") {
      const hit = tb.errors.find(([src]) => f.startsWith(src, i));
      if (hit) {
        out += hit[1];
        i += hit[0].length;
        continue;
      }
    }
    const prev = f[i - 1] ?? "";
    if ((/\d/.test(ch) || (ch === srcDec && /\d/.test(f[i + 1] ?? ""))) && !WORD.test(prev)) {
      const m = numRe.exec(f.slice(i));
      if (m) {
        out += m[0].replace(srcDec, dstDec);
        i += m[0].length;
        continue;
      }
    }
    if (brace > 0 && (ch === srcCol || (!toLocal && ch === "\\"))) { out += dstCol; i++; continue; }
    if (brace > 0 && ch === srcRow) { out += dstRow; i++; continue; }
    if (WORD_START.test(ch)) {
      let j = i + 1;
      while (j < f.length && WORD.test(f[j]!)) j++;
      const name = f.slice(i, j);
      if (f[j] === "(") {
        const m = /^((?:_xl[a-z]+\.)*)(.+)$/i.exec(name)!;
        const mapped = tb.fn.get(m[2]!.toUpperCase());
        out += mapped != null ? m[1]! + mapped : name;
      } else if (boolSrc && boolDst && f[j] !== "!" && prev !== "!" && boolSrc.includes(name.toUpperCase())) {
        out += boolDst[boolSrc.indexOf(name.toUpperCase())];
      } else out += name;
      i = j;
      continue;
    }
    out += brace === 0 && ch === srcArg ? dstArg : ch;
    i++;
  }
  return out;
}

/** An invariant formula (no "=") as the user reads it in their language. */
export function formulaToLocal(formula: string, r: Regional = regional()): string {
  return isInvariant(r) ? formula : translate(formula, r, true);
}

/** A formula typed in the user's language (no "=") back to the invariant form stored in files. */
export function formulaFromLocal(formula: string, r: Regional = regional()): string {
  return isInvariant(r) ? formula : translate(formula, r, false);
}

/**
 * What the user typed in a cell, made invariant for setCellInput: a local formula, number
 * ("1 234,5") or boolean ("VRAI"). Anything else passes through unchanged.
 */
export function inputFromLocal(raw: string, r: Regional = regional()): string {
  if (isInvariant(r)) return raw;
  if (raw.startsWith("=")) return "=" + formulaFromLocal(raw.slice(1), r);
  const num = parseLocalNumber(raw, r);
  if (num != null) return num;
  const up = raw.trim().toUpperCase();
  if (r.bool && up === r.bool[0]) return "TRUE";
  if (r.bool && up === r.bool[1]) return "FALSE";
  return raw;
}
