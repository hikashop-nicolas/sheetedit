import { colToLetters, getCell, type Workbook } from "./model";
import { listWorkbookTables, type WorkbookTable } from "../adapters/xlsx/tables";

// Structured (table) references: Table1[Units], Table1[[#Headers],[Units]], Table1[@Units].
//
// A formula engine that only knows A1 cannot see these at all, so a workbook whose totals are
// written against its table - which is what Excel's own UI produces the moment a range becomes a
// table - showed the file's stored value and a "could not evaluate" note for every one of them.
// They are rewritten to plain ranges before the formula is parsed, the same way LET is expanded.
//
// A selector we cannot honestly resolve (a totals row, which the model does not track; a column
// name that is not there) is replaced by a call to a function that does not exist, so the parse
// stops and the cell keeps the file's stored value under a "could not be evaluated" note. Leaving
// the text alone was worse: the engine read the bare table name and quietly produced a number.

/** Quick reject so the common formula never pays for the scan. */
export const hasTableRef = (formula: string): boolean => formula.includes("[");

const quoteSheet = (name: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);

/** A table's column names, from its header row (or Column1..N when it has none). */
function columnsOf(wb: Workbook, tbl: WorkbookTable): string[] {
  const sheet = wb.sheets[tbl.sheetIndex];
  const out: string[] = [];
  for (let c = tbl.c1; c <= tbl.c2; c++) {
    const header = sheet && tbl.headerRows > 0 ? getCell(sheet, tbl.r1, c)?.value : undefined;
    out.push(header != null && header !== "" ? String(header) : `Column${c - tbl.c1 + 1}`);
  }
  return out;
}

/** Split the inside of a [...] on top-level commas, honouring nested brackets and 'escapes. */
function splitParts(inner: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let cur = "";
  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i]!;
    if (ch === "'") { cur += inner[++i] ?? ""; continue; } // ' escapes the next character
    if (ch === "[") { depth++; cur += ch; continue; }
    if (ch === "]") { depth--; cur += ch; continue; }
    if (ch === "," && depth === 0) { parts.push(cur); cur = ""; continue; }
    cur += ch;
  }
  parts.push(cur);
  return parts.map((p) => p.trim()).filter((p) => p !== "");
}

/** "[Units]" / "Units" -> Units, with ' escapes removed. */
const bare = (s: string): string => {
  const body = s.startsWith("[") && s.endsWith("]") ? s.slice(1, -1) : s;
  return body.replace(/'(.)/g, "$1").trim();
};

/** The 0-based column offsets a column spec covers, or null when a name is not in the table. */
function columnBand(spec: string, columns: string[]): { from: number; to: number } | null {
  // "[A]:[B]" is a span; a single name is one column.
  const colon = spec.match(/^(\[[^\]]*\]|[^:]+):(\[[^\]]*\]|.+)$/);
  const find = (name: string): number => columns.findIndex((c) => c.toLowerCase() === name.toLowerCase());
  if (colon) {
    const a = find(bare(colon[1]!)), b = find(bare(colon[2]!));
    if (a < 0 || b < 0) return null;
    return { from: Math.min(a, b), to: Math.max(a, b) };
  }
  const i = find(bare(spec));
  return i < 0 ? null : { from: i, to: i };
}

/** One `Name[...]` resolved to an A1 range, or null when it cannot be honestly resolved. */
function resolveOne(wb: Workbook, tables: WorkbookTable[], name: string, inner: string, row: number): string | null {
  const key = name.toLowerCase();
  const tbl = tables.find((t) => t.displayName.toLowerCase() === key || t.name.toLowerCase() === key);
  const sheet = tbl && wb.sheets[tbl.sheetIndex];
  if (!tbl || !sheet) return null;
  const columns = columnsOf(wb, tbl);

  let r1 = tbl.r1 + tbl.headerRows;
  let r2 = tbl.r2;
  let band: { from: number; to: number } = { from: 0, to: columns.length - 1 };

  for (const part of splitParts(inner)) {
    const sel = bare(part).toLowerCase();
    if (sel === "#all") { r1 = tbl.r1; r2 = tbl.r2; continue; }
    if (sel === "#data") { r1 = tbl.r1 + tbl.headerRows; r2 = tbl.r2; continue; }
    if (sel === "#headers") {
      if (!tbl.headerRows) return null; // no header row to point at
      r1 = tbl.r1; r2 = tbl.r1;
      continue;
    }
    // A totals row is not in the model, so pointing at one would be a guess.
    if (sel === "#totals") return null;
    if (sel === "#this row" || part.startsWith("@")) {
      if (row < tbl.r1 + tbl.headerRows || row > tbl.r2) return null; // the formula is outside the table
      r1 = row; r2 = row;
      const rest = part.startsWith("@") ? part.slice(1).trim() : "";
      if (rest) {
        const b = columnBand(rest, columns);
        if (!b) return null;
        band = b;
      }
      continue;
    }
    const b = columnBand(part, columns);
    if (!b) return null;
    band = b;
  }
  if (r2 < r1) return null; // an empty table body has no range to name
  const c1 = tbl.c1 + band.from, c2 = tbl.c1 + band.to;
  return `${quoteSheet(sheet.name)}!$${colToLetters(c1)}$${r1}:$${colToLetters(c2)}$${r2}`;
}

/** What an unresolvable structured reference becomes: no such function, so the parse stops. */
const UNRESOLVED = "_SHEETEDIT_UNRESOLVED_TABLE_REF()";

/**
 * Rewrite every structured reference in `formula` to the plain range it names. One that cannot be
 * resolved becomes a call the engine will reject, so the cell reports itself rather than showing a
 * number read from the wrong place.
 */
export function expandTableRefs(formula: string, wb: Workbook, row: number): string {
  let tables: WorkbookTable[] | undefined;
  let out = "";
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i]!;
    if (ch === '"') {
      // Copy a string literal across untouched; "" is an escaped quote inside one.
      const start = i++;
      while (i < formula.length && !(formula[i] === '"' && formula[i + 1] !== '"')) i += formula[i] === '"' ? 2 : 1;
      out += formula.slice(start, ++i);
      continue;
    }
    // A table name is an identifier immediately followed by "[" - a function call has "(" instead.
    const m = /^[A-Za-z_\\][A-Za-z0-9_.\\]*(?=\[)/.exec(formula.slice(i));
    if (!m) { out += ch; i++; continue; }
    const name = m[0];
    let depth = 0, j = i + name.length;
    for (; j < formula.length; j++) {
      const c = formula[j]!;
      if (c === "'") { j++; continue; }
      if (c === "[") depth++;
      else if (c === "]" && --depth === 0) break;
    }
    if (j >= formula.length) { out += formula.slice(i); break; } // unbalanced: leave it alone
    const inner = formula.slice(i + name.length + 1, j);
    tables ??= listWorkbookTables(wb);
    const range = resolveOne(wb, tables, name, inner, row);
    out += range ?? UNRESOLVED;
    i = j + 1;
  }
  return out;
}

/** A bare table name used as a reference means its data body, as a "Sheet!$A$2:$C$9" string. */
export function tableBodyRef(wb: Workbook, name: string): string | undefined {
  const key = name.toLowerCase();
  const tbl = listWorkbookTables(wb).find((t) => t.displayName.toLowerCase() === key || t.name.toLowerCase() === key);
  const sheet = tbl && wb.sheets[tbl.sheetIndex];
  if (!tbl || !sheet) return undefined;
  const r1 = tbl.r1 + tbl.headerRows;
  if (tbl.r2 < r1) return undefined;
  return `${quoteSheet(sheet.name)}!$${colToLetters(tbl.c1)}$${r1}:$${colToLetters(tbl.c2)}$${tbl.r2}`;
}
