// Formulas in the user's language: "SOMME(A1;1,5)" on screen, "SUM(A1,1.5)" in the workbook.
// A small tokenizer swaps function names, booleans, error literals, decimal and argument
// separators, leaving strings, quoted sheet names and structured-reference items untouched.

import { isInvariant, localDateEditText, localNumberInput, parseLocalNumber, regional, type Regional } from "./regional";
import { parseDateInput, serialToEditText } from "./dates";
import { numToStr } from "./model";

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

/** How a rule operand reads: a date or time bound, or a value that may be plain text. */
export interface OperandKind {
  date?: boolean;
  time?: boolean;
  /** Plain words become a quoted string, as Excel's dialogs do ("oui" -> "\"oui\""). */
  text?: boolean;
}

const CELL_REF = /^(?:(?:'[^']+'|[\p{L}\p{N}_.]+)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/u;
const INVARIANT_NUM = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;

/**
 * A rule operand typed in a dialog (validation bound, highlight value), made invariant: a local
 * formula, a number ("3,5" -> "3.5"), a date ("25/11/2026" -> its serial) or a quoted text.
 */
export function operandFromLocal(text: string, kind: OperandKind = {}, r: Regional = regional()): string {
  const t = text.trim();
  if (!t) return "";
  if (t.startsWith("=")) return formulaFromLocal(t.slice(1).trim(), r);
  if (kind.date || kind.time) {
    const d = parseDateInput(t, r.dayOrder);
    if (d) return numToStr(d.serial);
  }
  const n = parseLocalNumber(t, r) ?? (INVARIANT_NUM.test(t) ? t : null);
  if (n != null) return n;
  if (kind.text && !t.startsWith('"') && !CELL_REF.test(t)) return `"${t.replace(/"/g, '""')}"`;
  return formulaFromLocal(t, r);
}

/** The other way, to show a stored operand in a dialog: "46351" -> "25/11/2026", "3.5" -> "3,5". */
export function operandToLocal(stored: string | undefined, kind: OperandKind = {}, r: Regional = regional()): string {
  const s = (stored ?? "").trim();
  if (!s) return "";
  if (INVARIANT_NUM.test(s)) {
    const serial = Number(s);
    if (kind.time && !kind.date && serial >= 0 && serial < 1) {
      const secs = Math.round(serial * 86400);
      const p2 = (n: number) => String(n).padStart(2, "0");
      return `${p2(Math.floor(secs / 3600))}:${p2(Math.floor((secs % 3600) / 60))}${secs % 60 ? ":" + p2(secs % 60) : ""}`;
    }
    if (kind.date) {
      const withTime = !Number.isInteger(serial);
      const d = localDateEditText(serial, withTime, r) ?? serialToEditText(serial, withTime);
      if (d != null) return d;
    }
    return localNumberInput(s, r);
  }
  if (kind.text && /^"(?:[^"]|"")*"$/.test(s)) return s.slice(1, -1).replace(/""/g, '"');
  return "=" + formulaToLocal(s, r);
}

/** A validation list typed locally ("oui;non" in French) -> its items, numbers made invariant. */
export function listFromLocal(text: string, r: Regional = regional()): string[] {
  return text.split(r.argSep).map((x) => x.trim()).filter(Boolean).map((x) => parseLocalNumber(x, r) ?? x);
}

/** A validation list's items as the user types them back. */
export function listToLocal(values: string[], r: Regional = regional()): string {
  return values.map((v) => localNumberInput(v, r)).join(r.argSep);
}
