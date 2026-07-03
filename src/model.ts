import FormulaParser from "fast-formula-parser";
import { strFromU8, strToU8 } from "fflate";
// ---------------------------------------------------------------------------
// Model
// ---------------------------------------------------------------------------

export type CellKind = "n" | "s" | "b" | "e" | "blank";

/** Resolved visual formatting for a cell (read from the file's style pools). */
export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  color?: string; // CSS text colour
  bg?: string; // CSS fill colour
  align?: "left" | "center" | "right";
  /** Border presence + CSS colour per side. */
  borders?: { top?: string; right?: string; bottom?: string; left?: string };
}

export interface Cell {
  row: number; // 1-based
  col: number; // 1-based
  /** Canonical/editable value: the literal, or the cached result for a formula cell. */
  value: string;
  /** Serialization hint for `value`. */
  kind: CellKind;
  /** Formatted text for the grid (number format applied). Falls back to `value`. */
  display?: string;
  /** Number format for this cell: an xlsx format code, a built-in numFmtId, or none. */
  numFmt?: string | number;
  /** Formula text in A1 syntax, without the leading "=". Undefined if not a formula. */
  formula?: string;
  /** xlsx: the <c> element in the worksheet DOM (for surgical edits). */
  el?: Element;
  /** ods: the original <table:table-cell> element (cloned verbatim if untouched). */
  odfFormula?: string;
  /** xlsx @s style index / ods @table:style-name, preserved across edits. */
  style?: string;
  /** Resolved visual formatting (fonts/fills/borders/alignment) for the grid. */
  cellStyle?: CellStyle;
  /** User changed the value/formula (forces regeneration). */
  edited?: boolean;
  /** Recalc changed the cached value. */
  recomputed?: boolean;
  /** xlsx: shared-formula group (@si) this cell belongs to. */
  sharedSi?: string;
  /** The formula itself was added/changed/removed (not just its cached value). */
  fDirty?: boolean;
}

export interface Sheet {
  name: string;
  cells: Map<string, Cell>;
  maxRow: number; // 1-based extent of used cells (0 = empty)
  maxCol: number;
  /** 1-based column -> width in px (from the file's <cols>), when specified. */
  colWidths?: Map<number, number>;
  /** 1-based row -> height in px (from the file's <row ht>), when specified. */
  rowHeights?: Map<number, number>;
  /** Merged ranges (1-based, inclusive); the top-left cell holds the value. */
  merges?: { r1: number; c1: number; r2: number; c2: number }[];
  // xlsx
  doc?: Document;
  sheetData?: Element;
  path?: string;
  /** Column widths or row heights changed and the sheet XML must be re-serialized. */
  layoutDirty?: boolean;
  // ods
  tableEl?: Element;
  /** ods: 1-based row -> its table:style-name, preserved/updated so writeOds can re-emit it. */
  odsRowStyles?: Map<number, string>;
  /** ods: original <table:table-row> per expanded row index, so row attributes
      (visibility, default-cell-style-name, ...) survive a rebuild. */
  odsRowEls?: Map<number, Element>;
  /** ods: repeated content rows beyond REPEAT_CAP, preserved verbatim as runs. */
  odsRowRuns?: { from: number; to: number; el: Element }[];
  /** ods: original covered-table-cell elements that carry content or attributes, by "r:c". */
  odsCoveredEls?: Map<string, Element>;
  /** ods: expanded row range of the original table:table-header-rows group. */
  odsHeaderRows?: { el: Element; from: number; to: number };
  /** ods: rows must be re-emitted (cell/merge/row-height edits); untouched sheets keep their XML. */
  odsDirty?: boolean;
}

export interface Workbook {
  kind: "xlsx" | "ods";
  sheets: Sheet[];
  files: Record<string, Uint8Array>;
  contentDoc?: Document; // ods
  contentPath?: string; // ods
  stylesDoc?: Document; // xlsx xl/styles.xml, kept for style writes
  stylesDirty?: boolean; // xlsx styles.xml changed and must be re-serialized
}

/** A style change to apply to a cell (only the set fields change). */
export interface StyleChange {
  bold?: boolean;
  italic?: boolean;
  color?: string; // CSS "#rrggbb" text colour
  bg?: string; // CSS "#rrggbb" fill colour
  align?: "left" | "center" | "right";
  border?: boolean; // all-sides box border on/off (convenience)
  /** Per-side borders to set; each specified side is turned on/off, others kept. */
  borderSides?: { top?: boolean; right?: boolean; bottom?: boolean; left?: boolean };
}

export const key = (row: number, col: number): string => `${row}:${col}`;

export function getCell(sheet: Sheet, row: number, col: number): Cell | undefined {
  return sheet.cells.get(key(row, col));
}

export function ensureCell(sheet: Sheet, row: number, col: number): Cell {
  let c = sheet.cells.get(key(row, col));
  if (!c) {
    c = { row, col, value: "", kind: "blank" };
    sheet.cells.set(key(row, col), c);
  }
  if (row > sheet.maxRow) sheet.maxRow = row;
  if (col > sheet.maxCol) sheet.maxCol = col;
  return c;
}

export function noteExtent(sheet: Sheet, row: number, col: number): void {
  if (row > sheet.maxRow) sheet.maxRow = row;
  if (col > sheet.maxCol) sheet.maxCol = col;
}

/** Typed value of a cell for the formula engine: number, boolean, string or null. */
export function typedValue(cell: Cell | undefined): number | boolean | string | null {
  if (!cell) return null;
  if (cell.value === "") return cell.formula != null ? null : null;
  switch (cell.kind) {
    case "n": {
      const n = Number(cell.value);
      return Number.isFinite(n) ? n : null;
    }
    case "b":
      return cell.value === "TRUE" || cell.value === "1" || cell.value === "true";
    case "blank":
      return null;
    default:
      return cell.value;
  }
}

export const isNumeric = (s: string): boolean => {
  const t = s.trim();
  return t !== "" && /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t);
};

export function numToStr(n: number): string {
  if (Number.isInteger(n)) return String(n);
  return String(parseFloat(n.toPrecision(15)));
}

/** Apply a number format (code or built-in id) to a numeric value via SSF. */
export function formatNumber(fmt: string | number, value: string): string | undefined {
  const n = Number(value);
  if (!Number.isFinite(n)) return undefined;
  try {
    return FormulaParser.SSF.format(fmt, n);
  } catch {
    return undefined;
  }
}

/** Text shown in the grid for a cell: the formatted display, else the raw value. */
export const cellDisplay = (cell: Cell | undefined): string => (cell ? cell.display ?? cell.value : "");

// ---------------------------------------------------------------------------
// A1 helpers
// ---------------------------------------------------------------------------

export function colToLetters(col: number): string {
  let s = "";
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

export function lettersToCol(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function parseA1Ref(ref: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
  if (!m) return null;
  return { col: lettersToCol(m[1]!), row: Number(m[2]) };
}

export const MAX_COL = 16384; // XFD
export const MAX_ROW = 1048576;
export const REF_TOKEN = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\dA-Za-z_(])/g;

export function shiftRefsInChunk(chunk: string, dr: number, dc: number): string {
  return chunk.replace(REF_TOKEN, (m, cAbs: string, letters: string, rAbs: string, digits: string, off: number) => {
    const prev = off > 0 ? chunk[off - 1]! : "";
    if (/[A-Za-z0-9_$.[\]]/.test(prev)) return m; // inside a longer name (LOG10, tax2026, Table1[...])
    const col0 = lettersToCol(letters.toUpperCase());
    const row0 = Number(digits);
    if (col0 > MAX_COL || row0 > MAX_ROW) return m; // not a real cell ref (defined name)
    const col = cAbs ? col0 : col0 + dc;
    const row = rAbs ? row0 : row0 + dr;
    if (col < 1 || row < 1 || col > MAX_COL || row > MAX_ROW) return "#REF!";
    return cAbs + colToLetters(col) + rAbs + String(row);
  });
}

/**
 * Shift the relative A1 references in a formula by (dr, dc): the translation a
 * shared-formula child applies to its master's formula. Skips string literals
 * and quoted sheet names; absolute parts ($) stay put.
 */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  let out = "";
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i]!;
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === ch) {
          if (formula[j + 1] === ch) j += 2; // doubled quote = escaped
          else {
            j++;
            break;
          }
        } else j++;
      }
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    let next = formula.length;
    for (const q of ['"', "'"]) {
      const p = formula.indexOf(q, i);
      if (p !== -1 && p < next) next = p;
    }
    out += shiftRefsInChunk(formula.slice(i, next), dr, dc);
    i = next;
  }
  return out;
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

export const parseXml = (file: Uint8Array): Document => {
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("corrupt workbook XML");
  return doc;
};
// Optional parts (styles, rels): a broken one degrades instead of failing the open.
export const parseXmlOpt = (file: Uint8Array): Document | undefined => {
  try {
    return parseXml(file);
  } catch {
    return undefined;
  }
};

export function serializeXml(doc: Document): Uint8Array {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith("<?xml")) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return strToU8(s);
}

export const firstByLocal = (parent: Element, local: string): Element | undefined => {
  for (const ch of Array.from(parent.children)) if (ch.localName === local) return ch;
  return undefined;
};

export const removeByLocal = (parent: Element, local: string): void => {
  for (const ch of Array.from(parent.children)) if (ch.localName === local) parent.removeChild(ch);
};

