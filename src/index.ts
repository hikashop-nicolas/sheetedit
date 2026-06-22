import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import FormulaParser from "fast-formula-parser";
import { t } from "./i18n";

// sheetedit: a standalone, framework-agnostic, client-side spreadsheet editor for
// .xlsx (OOXML SpreadsheetML) and .ods (ODF spreadsheet). Both are zips of XML.
//
// Philosophy (same as the docx/odt siblings): edit in place and preserve everything
// untouched. For .xlsx we surgically update only the changed <c> elements in each
// worksheet's DOM, so styles, number formats, charts, other sheets and untouched cells
// survive byte-for-byte. For .ods we regenerate the table body from the model but clone
// untouched cells verbatim and preserve every other part of the archive.
//
// Formulas are first-class: they are kept on save, and a dependency-ordered recalc
// engine (fast-formula-parser, MIT) recomputes cached results when cells change. The
// model always stores formulas in A1 syntax; .ods formulas are translated to/from the
// ODF (`of:=[.A1]`) form at the edges.

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

const key = (row: number, col: number): string => `${row}:${col}`;

function getCell(sheet: Sheet, row: number, col: number): Cell | undefined {
  return sheet.cells.get(key(row, col));
}

function ensureCell(sheet: Sheet, row: number, col: number): Cell {
  let c = sheet.cells.get(key(row, col));
  if (!c) {
    c = { row, col, value: "", kind: "blank" };
    sheet.cells.set(key(row, col), c);
  }
  if (row > sheet.maxRow) sheet.maxRow = row;
  if (col > sheet.maxCol) sheet.maxCol = col;
  return c;
}

function noteExtent(sheet: Sheet, row: number, col: number): void {
  if (row > sheet.maxRow) sheet.maxRow = row;
  if (col > sheet.maxCol) sheet.maxCol = col;
}

/** Typed value of a cell for the formula engine: number, boolean, string or null. */
function typedValue(cell: Cell | undefined): number | boolean | string | null {
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

const isNumeric = (s: string): boolean => {
  const t = s.trim();
  return t !== "" && /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t);
};

function numToStr(n: number): string {
  if (Number.isInteger(n)) return String(n);
  return String(parseFloat(n.toPrecision(15)));
}

/** Apply a number format (code or built-in id) to a numeric value via SSF. */
function formatNumber(fmt: string | number, value: string): string | undefined {
  const n = Number(value);
  if (!Number.isFinite(n)) return undefined;
  try {
    return FormulaParser.SSF.format(fmt, n);
  } catch {
    return undefined;
  }
}

/** Text shown in the grid for a cell: the formatted display, else the raw value. */
const cellDisplay = (cell: Cell | undefined): string => (cell ? cell.display ?? cell.value : "");

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

function lettersToCol(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function parseA1Ref(ref: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
  if (!m) return null;
  return { col: lettersToCol(m[1]!), row: Number(m[2]) };
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

const parseXml = (file: Uint8Array): Document =>
  new DOMParser().parseFromString(strFromU8(file), "application/xml");

function serializeXml(doc: Document): Uint8Array {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith("<?xml")) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return strToU8(s);
}

const firstByLocal = (parent: Element, local: string): Element | undefined => {
  for (const ch of Array.from(parent.children)) if (ch.localName === local) return ch;
  return undefined;
};

const removeByLocal = (parent: Element, local: string): void => {
  for (const ch of Array.from(parent.children)) if (ch.localName === local) parent.removeChild(ch);
};

// ---------------------------------------------------------------------------
// xlsx (OOXML SpreadsheetML)
// ---------------------------------------------------------------------------

const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function readSharedStrings(file: Uint8Array | undefined): string[] {
  if (!file) return [];
  const doc = parseXml(file);
  return Array.from(doc.getElementsByTagName("si")).map((si) =>
    Array.from(si.getElementsByTagName("t"))
      .map((t) => t.textContent ?? "")
      .join(""),
  );
}

interface XlsxStyles {
  customFmt: Map<number, string>; // numFmtId -> format code (custom, id >= 164)
  xfNumFmtIds: number[]; // cellXfs index (the cell @s) -> numFmtId
  xfStyles: (CellStyle | undefined)[]; // cellXfs index (the cell @s) -> resolved style
}

// ARGB ("FFRRGGBB" or "RRGGBB") -> CSS "#rrggbb".
function argbToCss(argb: string | null | undefined): string | undefined {
  if (!argb) return undefined;
  const h = argb.length === 8 ? argb.slice(2) : argb;
  return /^[0-9a-fA-F]{6}$/.test(h) ? "#" + h.toLowerCase() : undefined;
}

// Excel tint: negative darkens toward black, positive lightens toward white.
function applyTint(hex: string, tint: number): string {
  const ch = (i: number) => {
    const c = parseInt(hex.slice(i, i + 2), 16);
    const v = tint < 0 ? c * (1 + tint) : c * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(v)))
      .toString(16)
      .padStart(2, "0");
  };
  return "#" + ch(1) + ch(3) + ch(5);
}

const findByLocal = (doc: Document, local: string): Element | undefined =>
  Array.from(doc.getElementsByTagName("*")).find((e) => e.localName === local);

// theme1.xml <clrScheme> -> array indexed by a <color theme="N"> index.
function readTheme(file: Uint8Array | undefined): string[] {
  const fallback = ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#0563c1", "#954f72"];
  if (!file) return fallback;
  try {
    const scheme = findByLocal(parseXml(file), "clrScheme");
    if (!scheme) return fallback;
    const byName: Record<string, string> = {};
    for (const el of Array.from(scheme.children)) {
      const c = el.firstElementChild;
      const css = c && (c.localName === "srgbClr" ? argbToCss(c.getAttribute("val")) : argbToCss(c.getAttribute("lastClr")));
      if (el.localName && css) byName[el.localName] = css;
    }
    // theme index order swaps dk/lt 1 and 2 vs the clrScheme element order.
    const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    return order.map((nm, i) => byName[nm] ?? fallback[i]!);
  } catch {
    return fallback;
  }
}

// Resolve a <color> element (rgb, or theme + tint) to a CSS colour.
function resolveColor(el: Element | undefined, theme: string[]): string | undefined {
  if (!el) return undefined;
  const rgb = el.getAttribute("rgb");
  if (rgb) return argbToCss(rgb);
  const t = el.getAttribute("theme");
  if (t != null) {
    const base = theme[Number(t)] ?? "#000000";
    const tint = Number(el.getAttribute("tint") || "0");
    return tint ? applyTint(base, tint) : base;
  }
  return undefined;
}

function readXlsxStyles(doc: Document | undefined, theme: string[]): XlsxStyles {
  const customFmt = new Map<number, string>();
  const xfNumFmtIds: number[] = [];
  const xfStyles: (CellStyle | undefined)[] = [];
  if (!doc) return { customFmt, xfNumFmtIds, xfStyles };
  for (const nf of Array.from(doc.getElementsByTagName("numFmt"))) {
    const id = Number(nf.getAttribute("numFmtId"));
    const code = nf.getAttribute("formatCode");
    if (Number.isFinite(id) && code != null) customFmt.set(id, code);
  }

  const pool = (local: string) => {
    const parent = firstByLocal(doc.documentElement, local);
    return parent ? Array.from(parent.children).filter((e) => e.localName === local.replace(/s$/, "")) : [];
  };
  const fonts = pool("fonts").map((f) => ({
    bold: !!firstByLocal(f, "b"),
    italic: !!firstByLocal(f, "i"),
    color: resolveColor(firstByLocal(f, "color"), theme),
  }));
  const fills = pool("fills").map((fl) => {
    const pat = firstByLocal(fl, "patternFill");
    return pat?.getAttribute("patternType") === "solid" ? resolveColor(firstByLocal(pat, "fgColor"), theme) : undefined;
  });
  const borders = pool("borders").map((bd) => {
    const side = (name: string): string | undefined => {
      const s = firstByLocal(bd, name);
      return s?.getAttribute("style") ? (resolveColor(firstByLocal(s, "color"), theme) ?? "#444") : undefined;
    };
    const b = { top: side("top"), right: side("right"), bottom: side("bottom"), left: side("left") };
    return b.top || b.right || b.bottom || b.left ? b : undefined;
  });

  // The cell @s indexes <cellXfs>, not <cellStyleXfs>; read that list specifically.
  const cellXfs = doc.getElementsByTagName("cellXfs")[0];
  if (cellXfs) {
    for (const xf of Array.from(cellXfs.children)) {
      if (xf.localName !== "xf") continue;
      xfNumFmtIds.push(Number(xf.getAttribute("numFmtId") || "0"));
      const st: CellStyle = {};
      const font = fonts[Number(xf.getAttribute("fontId") || "0")];
      if (font) {
        if (font.bold) st.bold = true;
        if (font.italic) st.italic = true;
        if (font.color) st.color = font.color;
      }
      const fill = fills[Number(xf.getAttribute("fillId") || "0")];
      if (fill) st.bg = fill;
      const border = borders[Number(xf.getAttribute("borderId") || "0")];
      if (border) st.borders = border;
      const align = firstByLocal(xf, "alignment")?.getAttribute("horizontal");
      if (align === "center" || align === "right" || align === "left") st.align = align;
      xfStyles.push(Object.keys(st).length ? st : undefined);
    }
  }
  return { customFmt, xfNumFmtIds, xfStyles };
}

/** Resolve a cell's number format (code or built-in id), or undefined for General. */
function resolveXlsxFmt(styles: XlsxStyles, s: string | undefined): string | number | undefined {
  if (s == null) return undefined;
  const numFmtId = styles.xfNumFmtIds[Number(s)];
  if (numFmtId == null || numFmtId === 0) return undefined; // 0 = General
  const custom = styles.customFmt.get(numFmtId);
  if (custom != null) return custom === "General" ? undefined : custom;
  return numFmtId; // built-in id; SSF resolves it
}

function readXlsx(files: Record<string, Uint8Array>): Workbook {
  const wb: Workbook = { kind: "xlsx", sheets: [], files };
  const wbXml = files["xl/workbook.xml"];
  if (!wbXml) throw new Error("not an .xlsx: xl/workbook.xml missing");
  const wbDoc = parseXml(wbXml);
  const rels = new Map<string, string>();
  const relsFile = files["xl/_rels/workbook.xml.rels"];
  if (relsFile) {
    for (const r of Array.from(parseXml(relsFile).getElementsByTagName("Relationship"))) {
      const id = r.getAttribute("Id");
      const target = r.getAttribute("Target");
      if (id && target) rels.set(id, target);
    }
  }
  const shared = readSharedStrings(files["xl/sharedStrings.xml"]);
  const theme = readTheme(files["xl/theme/theme1.xml"]);
  wb.stylesDoc = files["xl/styles.xml"] ? parseXml(files["xl/styles.xml"]) : undefined;
  const styles = readXlsxStyles(wb.stylesDoc, theme);

  let n = 0;
  for (const sheetEl of Array.from(wbDoc.getElementsByTagName("sheet"))) {
    n++;
    const name = sheetEl.getAttribute("name") ?? `Sheet${n}`;
    const rid = sheetEl.getAttribute("r:id") ?? sheetEl.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "id",
    );
    let target = (rid && rels.get(rid)) || `worksheets/sheet${n}.xml`;
    const path = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
    const wsFile = files[path];
    const sheet: Sheet = { name, cells: new Map(), maxRow: 0, maxCol: 0, path };
    if (wsFile) {
      const doc = parseXml(wsFile);
      const sheetData = doc.getElementsByTagName("sheetData")[0];
      sheet.doc = doc;
      sheet.sheetData = sheetData;
      // Column widths: <cols><col min max width/></cols>. Width is in character units;
      // convert to px (~7px per char + padding for the default font).
      const colsEl = doc.getElementsByTagName("cols")[0];
      if (colsEl) {
        const cw = new Map<number, number>();
        for (const col of Array.from(colsEl.children)) {
          if (col.localName !== "col") continue;
          const min = Number(col.getAttribute("min") || "0");
          const max = Number(col.getAttribute("max") || "0");
          const width = Number(col.getAttribute("width") || "0");
          if (!min || !width) continue;
          const px = Math.round(width * 7 + 5);
          for (let c = min; c <= Math.min(max || min, min + 1000); c++) cw.set(c, px);
        }
        if (cw.size) sheet.colWidths = cw;
      }
      // Row heights: <row r ht customHeight/>. ht is in points; convert to px (~4/3 px/pt).
      if (sheetData) {
        const rh = new Map<number, number>();
        for (const rowEl of Array.from(sheetData.children)) {
          if (rowEl.localName !== "row") continue;
          const r = Number(rowEl.getAttribute("r") || "0");
          const ht = Number(rowEl.getAttribute("ht") || "0");
          if (r && ht) rh.set(r, Math.round((ht * 4) / 3));
        }
        if (rh.size) sheet.rowHeights = rh;
      }
      // Merged ranges: <mergeCells><mergeCell ref="B1:C1"/></mergeCells>.
      const mergeEls = doc.getElementsByTagName("mergeCell");
      if (mergeEls.length) {
        const merges: { r1: number; c1: number; r2: number; c2: number }[] = [];
        for (const m of Array.from(mergeEls)) {
          const ref = m.getAttribute("ref");
          const [a, b] = (ref ?? "").split(":");
          const p1 = a ? parseA1Ref(a) : null;
          const p2 = b ? parseA1Ref(b) : null;
          if (p1 && p2) merges.push({ r1: p1.row, c1: p1.col, r2: p2.row, c2: p2.col });
        }
        if (merges.length) sheet.merges = merges;
      }
      if (sheetData) readSheetData(sheet, sheetData, shared, styles);
    }
    wb.sheets.push(sheet);
  }
  return wb;
}

function readSheetData(sheet: Sheet, sheetData: Element, shared: string[], styles: XlsxStyles): void {
  for (const rowEl of Array.from(sheetData.getElementsByTagName("row"))) {
    const rAttr = rowEl.getAttribute("r");
    let rowNum = rAttr ? Number(rAttr) : 0;
    let colCursor = 0;
    for (const c of Array.from(rowEl.children)) {
      if (c.localName !== "c") continue;
      const ref = c.getAttribute("r");
      let row = rowNum;
      let col: number;
      if (ref) {
        const p = parseA1Ref(ref);
        if (!p) continue;
        row = p.row;
        col = p.col;
        colCursor = col;
      } else {
        col = ++colCursor;
      }
      if (!row) continue;
      const t = c.getAttribute("t");
      const fEl = firstByLocal(c, "f");
      const vEl = firstByLocal(c, "v");
      const isEl = firstByLocal(c, "is");
      const formulaText = fEl?.textContent?.trim();
      const formula = formulaText ? formulaText : undefined;

      let value = "";
      let kind: CellKind = "blank";
      if (t === "s") {
        value = shared[Number(vEl?.textContent ?? "0")] ?? "";
        kind = "s";
      } else if (t === "inlineStr") {
        value = isEl ? Array.from(isEl.getElementsByTagName("t")).map((x) => x.textContent ?? "").join("") : "";
        kind = "s";
      } else if (t === "str") {
        value = vEl?.textContent ?? "";
        kind = "s";
      } else if (t === "b") {
        value = vEl?.textContent === "1" ? "TRUE" : "FALSE";
        kind = "b";
      } else if (t === "e") {
        value = vEl?.textContent ?? "";
        kind = "e";
      } else {
        value = vEl?.textContent ?? "";
        kind = value === "" ? "blank" : "n";
      }

      const cell: Cell = {
        row,
        col,
        value,
        kind,
        formula,
        el: c,
        style: c.getAttribute("s") ?? undefined,
      };
      if (kind === "n") {
        const fmt = resolveXlsxFmt(styles, cell.style);
        if (fmt != null) {
          cell.numFmt = fmt;
          const d = formatNumber(fmt, value);
          if (d != null) cell.display = d;
        }
      }
      if (cell.style != null) cell.cellStyle = styles.xfStyles[Number(cell.style)];
      sheet.cells.set(key(row, col), cell);
      noteExtent(sheet, row, col);
    }
  }
}

function ensureXlsxCellEl(sheet: Sheet, cell: Cell): Element {
  if (cell.el) return cell.el;
  const doc = sheet.doc!;
  const sheetData = sheet.sheetData!;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  // find or create the <row>
  let rowEl: Element | undefined;
  let insertRowBefore: Element | null = null;
  for (const r of Array.from(sheetData.children)) {
    if (r.localName !== "row") continue;
    const rn = Number(r.getAttribute("r") || "0");
    if (rn === cell.row) {
      rowEl = r;
      break;
    }
    if (rn > cell.row) {
      insertRowBefore = r;
      break;
    }
  }
  if (!rowEl) {
    rowEl = doc.createElementNS(ns, "row");
    rowEl.setAttribute("r", String(cell.row));
    sheetData.insertBefore(rowEl, insertRowBefore);
  }
  // find or create the <c> in column order
  const ref = colToLetters(cell.col) + cell.row;
  let insertCellBefore: Element | null = null;
  for (const c of Array.from(rowEl.children)) {
    if (c.localName !== "c") continue;
    const cref = c.getAttribute("r");
    const p = cref ? parseA1Ref(cref) : null;
    if (p && p.col === cell.col) return (cell.el = c);
    if (p && p.col > cell.col) {
      insertCellBefore = c;
      break;
    }
  }
  const cEl = doc.createElementNS(ns, "c");
  cEl.setAttribute("r", ref);
  if (cell.style) cEl.setAttribute("s", cell.style);
  rowEl.insertBefore(cEl, insertCellBefore);
  cell.el = cEl;
  return cEl;
}

function writeXlsxCell(sheet: Sheet, cell: Cell): void {
  const doc = sheet.doc!;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const c = ensureXlsxCellEl(sheet, cell);
  removeByLocal(c, "f");
  removeByLocal(c, "v");
  removeByLocal(c, "is");
  const addV = (text: string) => {
    const v = doc.createElementNS(ns, "v");
    v.textContent = text;
    c.appendChild(v);
  };
  if (cell.formula != null) {
    const f = doc.createElementNS(ns, "f");
    f.textContent = cell.formula;
    c.appendChild(f);
    if (cell.kind === "n") {
      c.removeAttribute("t");
      if (cell.value !== "") addV(cell.value);
    } else if (cell.kind === "b") {
      c.setAttribute("t", "b");
      addV(cell.value === "TRUE" || cell.value === "1" ? "1" : "0");
    } else if (cell.kind === "e") {
      c.setAttribute("t", "e");
      addV(cell.value);
    } else if (cell.kind === "blank" || cell.value === "") {
      c.removeAttribute("t");
    } else {
      c.setAttribute("t", "str");
      addV(cell.value);
    }
    return;
  }
  // literal
  if (cell.value === "" || cell.kind === "blank") {
    c.removeAttribute("t");
  } else if (cell.kind === "n") {
    c.removeAttribute("t");
    addV(cell.value);
  } else if (cell.kind === "b") {
    c.setAttribute("t", "b");
    addV(cell.value === "TRUE" || cell.value === "1" ? "1" : "0");
  } else if (cell.kind === "e") {
    c.setAttribute("t", "e");
    addV(cell.value);
  } else {
    c.setAttribute("t", "inlineStr");
    const is = doc.createElementNS(ns, "is");
    const t = doc.createElementNS(ns, "t");
    t.setAttribute("xml:space", "preserve");
    t.textContent = cell.value;
    is.appendChild(t);
    c.appendChild(is);
  }
}

// Set a single column's width (px) in the worksheet's <cols>, creating <cols>/<col>
// as needed and splitting any existing run that covers the column. Keeps colWidths in sync.
export function setXlsxColWidth(sheet: Sheet, col: number, px: number): void {
  if (!sheet.colWidths) sheet.colWidths = new Map();
  sheet.colWidths.set(col, px);
  const doc = sheet.doc;
  if (!doc) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const width = Math.max(0, (px - 5) / 7);
  let colsEl = doc.getElementsByTagName("cols")[0] as Element | undefined;
  if (!colsEl) {
    colsEl = doc.createElementNS(ns, "cols");
    // <cols> must precede <sheetData> per the schema.
    sheet.sheetData?.parentNode?.insertBefore(colsEl, sheet.sheetData);
  }
  // Narrow any run that spans `col` so we can give `col` its own entry.
  for (const c of Array.from(colsEl.children)) {
    if (c.localName !== "col") continue;
    const min = Number(c.getAttribute("min") || "0");
    const max = Number(c.getAttribute("max") || String(min));
    if (col < min || col > max) continue;
    if (min === max) {
      c.setAttribute("width", String(width));
      c.setAttribute("customWidth", "1");
      sheet.layoutDirty = true;
      return;
    }
    // Split: left part [min..col-1], right part [col+1..max], plus the single col.
    if (col > min) {
      const left = c.cloneNode(true) as Element;
      left.setAttribute("min", String(min));
      left.setAttribute("max", String(col - 1));
      colsEl.insertBefore(left, c);
    }
    if (col < max) {
      const right = c.cloneNode(true) as Element;
      right.setAttribute("min", String(col + 1));
      right.setAttribute("max", String(max));
      colsEl.insertBefore(right, c);
    }
    c.setAttribute("min", String(col));
    c.setAttribute("max", String(col));
    c.setAttribute("width", String(width));
    c.setAttribute("customWidth", "1");
    sheet.layoutDirty = true;
    return;
  }
  const colEl = doc.createElementNS(ns, "col");
  colEl.setAttribute("min", String(col));
  colEl.setAttribute("max", String(col));
  colEl.setAttribute("width", String(width));
  colEl.setAttribute("customWidth", "1");
  colsEl.appendChild(colEl);
  sheet.layoutDirty = true;
}

// Set a single row's height (px) on its <row>, creating the row element if absent.
export function setXlsxRowHeight(sheet: Sheet, row: number, px: number): void {
  if (!sheet.rowHeights) sheet.rowHeights = new Map();
  sheet.rowHeights.set(row, px);
  const doc = sheet.doc;
  const sd = sheet.sheetData;
  if (!doc || !sd) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const pt = (px * 3) / 4;
  let rowEl: Element | undefined;
  for (const re of Array.from(sd.children)) {
    if (re.localName === "row" && Number(re.getAttribute("r") || "0") === row) {
      rowEl = re;
      break;
    }
  }
  if (!rowEl) {
    rowEl = doc.createElementNS(ns, "row");
    rowEl.setAttribute("r", String(row));
    // Insert keeping rows in ascending order.
    let next: Element | null = null;
    for (const re of Array.from(sd.children)) {
      if (re.localName === "row" && Number(re.getAttribute("r") || "0") > row) {
        next = re;
        break;
      }
    }
    sd.insertBefore(rowEl, next);
  }
  rowEl.setAttribute("ht", String(pt));
  rowEl.setAttribute("customHeight", "1");
  sheet.layoutDirty = true;
}

// Add or remove a merged range (1-based, inclusive). The top-left cell shows through;
// any cells the merge hides keep their data (so unmerging restores it). Updates the
// worksheet's <mergeCells> element and the in-memory merge list.
export function setXlsxMerge(
  sheet: Sheet,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  merge: boolean,
): void {
  const top = Math.min(r1, r2),
    left = Math.min(c1, c2),
    bottom = Math.max(r1, r2),
    right = Math.max(c1, c2);
  const ref = `${colToLetters(left)}${top}:${colToLetters(right)}${bottom}`;
  const merges = (sheet.merges ??= []);
  const idx = merges.findIndex((m) => m.r1 === top && m.c1 === left && m.r2 === bottom && m.c2 === right);
  if (merge) {
    if (idx === -1) merges.push({ r1: top, c1: left, r2: bottom, c2: right });
  } else if (idx !== -1) {
    merges.splice(idx, 1);
  }

  const doc = sheet.doc;
  if (!doc) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  let mcEl = doc.getElementsByTagName("mergeCells")[0] as Element | undefined;
  if (merge) {
    if (!mcEl) {
      mcEl = doc.createElementNS(ns, "mergeCells");
      // <mergeCells> follows <sheetData> in the schema.
      sheet.sheetData?.parentNode?.insertBefore(mcEl, sheet.sheetData.nextSibling);
    }
    const exists = Array.from(mcEl.children).some((m) => m.getAttribute("ref") === ref);
    if (!exists) {
      const mc = doc.createElementNS(ns, "mergeCell");
      mc.setAttribute("ref", ref);
      mcEl.appendChild(mc);
    }
  } else if (mcEl) {
    for (const m of Array.from(mcEl.children))
      if (m.getAttribute("ref") === ref) mcEl.removeChild(m);
  }
  if (mcEl) {
    if (mcEl.children.length === 0) mcEl.parentNode?.removeChild(mcEl);
    else mcEl.setAttribute("count", String(mcEl.children.length));
  }
  sheet.layoutDirty = true;
}

const xmlOf = (el: Element): string => new XMLSerializer().serializeToString(el);

// Find a matching child in a style pool (deduped by serialized form) or append it;
// returns its index and keeps the pool's count attribute in sync.
function poolIndex(parent: Element, candidate: Element): number {
  const want = xmlOf(candidate);
  const kids = Array.from(parent.children);
  for (let i = 0; i < kids.length; i++) if (xmlOf(kids[i]!) === want) return i;
  parent.appendChild(candidate);
  parent.setAttribute("count", String(parent.children.length));
  return parent.children.length - 1;
}

const argbOf = (css: string): string => "FF" + css.replace("#", "").toUpperCase();

// Compute the resulting CellStyle from the current one plus a change. Shared by the
// xlsx and ods style writers so both stay consistent; affected borders become black.
function mergeCellStyle(cur: CellStyle, change: StyleChange): CellStyle {
  const sides = {
    top: !!cur.borders?.top,
    right: !!cur.borders?.right,
    bottom: !!cur.borders?.bottom,
    left: !!cur.borders?.left,
  };
  if (change.border !== undefined) sides.top = sides.right = sides.bottom = sides.left = change.border;
  if (change.borderSides) Object.assign(sides, change.borderSides);
  const any = sides.top || sides.right || sides.bottom || sides.left;
  return {
    bold: change.bold ?? cur.bold,
    italic: change.italic ?? cur.italic,
    color: change.color ?? cur.color,
    bg: change.bg ?? cur.bg,
    align: change.align ?? cur.align,
    borders: any
      ? {
          top: sides.top ? "#000000" : undefined,
          right: sides.right ? "#000000" : undefined,
          bottom: sides.bottom ? "#000000" : undefined,
          left: sides.left ? "#000000" : undefined,
        }
      : undefined,
  };
}

/**
 * Apply a style change to a cell, managing the xlsx style pools: derive a new font /
 * fill / border from the cell's current format plus the change, find-or-create each in
 * styles.xml, find-or-create the combined <xf>, and point the cell at it.
 */
export function setXlsxCellStyle(wb: Workbook, sheet: Sheet, cell: Cell, change: StyleChange): void {
  const doc = wb.stylesDoc;
  if (!doc) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const ce = (name: string) => doc.createElementNS(ns, name);
  const root = doc.documentElement;
  const pool = (name: string): Element => firstByLocal(root, name) ?? (root.appendChild(ce(name)) as Element);
  const fontsEl = pool("fonts");
  const fillsEl = pool("fills");
  const bordersEl = pool("borders");
  const cellXfsEl = pool("cellXfs");

  const curXf = cellXfsEl.children[cell.style ? Number(cell.style) : 0];
  const numFmtId = curXf?.getAttribute("numFmtId") || "0";
  const curFontId = Number(curXf?.getAttribute("fontId") || "0");
  const curFillId = Number(curXf?.getAttribute("fillId") || "0");
  const curBorderId = Number(curXf?.getAttribute("borderId") || "0");

  const cur = cell.cellStyle ?? {};
  const bold = change.bold ?? cur.bold;
  const italic = change.italic ?? cur.italic;
  const color = change.color ?? cur.color;
  const bg = change.bg ?? cur.bg;
  const align = change.align ?? cur.align;
  // Border sides: start from the current borders, apply the all-sides toggle and/or per-side change.
  const curSides = {
    top: !!cur.borders?.top,
    right: !!cur.borders?.right,
    bottom: !!cur.borders?.bottom,
    left: !!cur.borders?.left,
  };
  let sides = curSides;
  let borderChanged = false;
  if (change.border !== undefined) {
    sides = { top: change.border, right: change.border, bottom: change.border, left: change.border };
    borderChanged = true;
  }
  if (change.borderSides) {
    sides = { ...sides, ...change.borderSides };
    borderChanged = true;
  }

  // Font: clone the current one and toggle bold/italic/colour.
  const baseFont = fontsEl.children[curFontId];
  const font = baseFont ? (baseFont.cloneNode(true) as Element) : ce("font");
  const flag = (tag: string, on: boolean | undefined) => {
    const ex = firstByLocal(font, tag);
    if (on && !ex) font.appendChild(ce(tag));
    else if (!on && ex) font.removeChild(ex);
  };
  flag("b", bold);
  flag("i", italic);
  if (color) {
    const col = firstByLocal(font, "color") ?? (font.appendChild(ce("color")) as Element);
    col.removeAttribute("theme");
    col.removeAttribute("tint");
    col.setAttribute("rgb", argbOf(color));
  }
  const fontId = poolIndex(fontsEl, font);

  // Fill (solid) when set; else keep the current fill.
  let fillId = curFillId;
  if (bg) {
    const fill = ce("fill");
    const pat = ce("patternFill");
    pat.setAttribute("patternType", "solid");
    const fg = ce("fgColor");
    fg.setAttribute("rgb", argbOf(bg));
    pat.appendChild(fg);
    fill.appendChild(pat);
    fillId = poolIndex(fillsEl, fill);
  }

  // Border: rebuild the per-side border element only when the change touches borders.
  let borderId = curBorderId;
  if (borderChanged) {
    const bd = ce("border");
    for (const side of ["left", "right", "top", "bottom"] as const) {
      const s = ce(side);
      if (sides[side]) {
        s.setAttribute("style", "thin");
        const cc = ce("color");
        cc.setAttribute("rgb", "FF000000");
        s.appendChild(cc);
      }
      bd.appendChild(s);
    }
    borderId = poolIndex(bordersEl, bd);
  }

  const xf = ce("xf");
  xf.setAttribute("numFmtId", numFmtId);
  xf.setAttribute("fontId", String(fontId));
  xf.setAttribute("fillId", String(fillId));
  xf.setAttribute("borderId", String(borderId));
  xf.setAttribute("xfId", "0");
  xf.setAttribute("applyFont", "1");
  if (bg) xf.setAttribute("applyFill", "1");
  if (borderId) xf.setAttribute("applyBorder", "1");
  if (align) {
    xf.setAttribute("applyAlignment", "1");
    const a = ce("alignment");
    a.setAttribute("horizontal", align);
    xf.appendChild(a);
  }
  const sIdx = poolIndex(cellXfsEl, xf);

  cell.style = String(sIdx);
  ensureXlsxCellEl(sheet, cell).setAttribute("s", String(sIdx));
  cell.cellStyle = mergeCellStyle(cur, change);
  cell.edited = true;
  wb.stylesDirty = true;
}

function writeXlsx(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    if (!sheet.doc || !sheet.sheetData) continue;
    let touched = false;
    for (const cell of sheet.cells.values()) {
      if (cell.edited || cell.recomputed) {
        writeXlsxCell(sheet, cell);
        touched = true;
      }
    }
    if ((touched || sheet.layoutDirty) && sheet.path) wb.files[sheet.path] = serializeXml(sheet.doc);
  }
  if (wb.stylesDirty && wb.stylesDoc) wb.files["xl/styles.xml"] = serializeXml(wb.stylesDoc);
}

// ---------------------------------------------------------------------------
// ods (ODF spreadsheet)
// ---------------------------------------------------------------------------

const ODS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
};
const REPEAT_CAP = 1024;

/** Replace text outside single-quoted string literals. */
function replaceOutsideStrings(s: string, fn: (chunk: string) => string): string {
  let out = "";
  let i = 0;
  while (i < s.length) {
    const q = s.indexOf('"', i);
    const qq = s.indexOf("'", i);
    let next = -1;
    let quote = '"';
    if (q === -1 && qq === -1) next = -1;
    else if (q === -1) {
      next = qq;
      quote = "'";
    } else if (qq === -1) {
      next = q;
      quote = '"';
    } else if (q < qq) {
      next = q;
      quote = '"';
    } else {
      next = qq;
      quote = "'";
    }
    if (next === -1) {
      out += fn(s.slice(i));
      break;
    }
    out += fn(s.slice(i, next));
    const end = s.indexOf(quote, next + 1);
    if (end === -1) {
      out += s.slice(next);
      break;
    }
    out += s.slice(next, end + 1);
    i = end + 1;
  }
  return out;
}

/** ODF formula (`of:=[.A1]+[.B1]`) -> A1 (`A1+B1`). */
export function odfToA1(odf: string): string {
  let core = odf.replace(/^of:=/, "").replace(/^=/, "");
  core = core.replace(/\[([^\]]*)\]/g, (_, inner: string) => {
    if (!inner || inner.includes("#")) return inner.replace(/\./g, ""); // #REF! etc.
    const parts = inner.split(":");
    const mapped = parts.map((part) => {
      const dot = part.lastIndexOf(".");
      const sheet = dot >= 0 ? part.slice(0, dot) : "";
      const ref = dot >= 0 ? part.slice(dot + 1) : part;
      return { sheet, ref };
    });
    const sheet = mapped[0]!.sheet;
    const cells = mapped.map((m) => m.ref).join(":");
    return (sheet ? sheet + "!" : "") + cells;
  });
  return replaceOutsideStrings(core, (chunk) => chunk.replace(/;/g, ","));
}

/** A1 (`A1+B1`) -> ODF formula (`of:=[.A1]+[.B1]`). Used only for user-typed formulas. */
export function a1ToOdf(a1: string): string {
  const refRe =
    /(?<![A-Za-z0-9_.$])(?:('[^']*'|[A-Za-z_][A-Za-z0-9_]*)!)?(\$?[A-Z]{1,3}\$?[0-9]+)(?::(\$?[A-Z]{1,3}\$?[0-9]+))?(?![A-Za-z0-9_(])/g;
  const converted = replaceOutsideStrings(a1, (chunk) => {
    const semi = chunk.replace(/,/g, ";");
    return semi.replace(refRe, (_m, sheet: string | undefined, c1: string, c2?: string) => {
      const sh = sheet ?? "";
      const range = c2 ? `${c1}:.${c2}` : c1;
      return `[${sh}.${range}]`;
    });
  });
  return "of:=" + converted;
}

function odsCellText(cell: Element): string {
  return Array.from(cell.getElementsByTagName("text:p"))
    .map((p) => p.textContent ?? "")
    .join("\n");
}

// Convert an ODF length ("2.5cm", "96pt", "1in", "0.45mm", "12px") to CSS px.
function odsLenToPx(len: string | null | undefined): number | undefined {
  if (!len) return undefined;
  const m = /^([\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(len.trim());
  if (!m) return undefined;
  const v = parseFloat(m[1]!);
  switch (m[2]) {
    case "cm":
      return Math.round((v * 96) / 2.54);
    case "mm":
      return Math.round((v * 96) / 25.4);
    case "in":
      return Math.round(v * 96);
    case "pt":
      return Math.round((v * 96) / 72);
    case "pc":
      return Math.round(v * 16);
    default:
      return Math.round(v); // px or unitless
  }
}

const odsColorOf = (v: string | null): string | undefined =>
  v && v !== "transparent" && v !== "none" ? (/^#[0-9a-fA-F]{6}$/.test(v) ? v.toLowerCase() : v) : undefined;

// A border value is like "0.5pt solid #000000" or "none"; keep the colour if it draws.
function odsBorderColor(v: string | null): string | undefined {
  if (!v || v === "none" || /(^|\s)0(\.0+)?(pt|cm|mm|in|px)?\s/.test(" " + v + " ")) return undefined;
  const m = /#[0-9a-fA-F]{6}/.exec(v);
  return m ? m[0].toLowerCase() : "#000000";
}

interface OdsStyles {
  cell: Map<string, CellStyle>; // family table-cell -> resolved style
  colW: Map<string, number>; // family table-column -> width px
  rowH: Map<string, number>; // family table-row -> height px
}

// Parse <style:style> from the given docs (content.xml automatic styles + styles.xml),
// resolving table-cell parent chains, and the column/row dimension styles.
function parseOdsStyles(docs: Document[]): OdsStyles {
  const raw = new Map<string, { el: Element; parent?: string }>();
  const colW = new Map<string, number>();
  const rowH = new Map<string, number>();
  for (const doc of docs) {
    for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
      const name = st.getAttribute("style:name");
      if (!name) continue;
      const family = st.getAttribute("style:family");
      if (family === "table-cell") {
        raw.set(name, { el: st, parent: st.getAttribute("style:parent-style-name") || undefined });
      } else if (family === "table-column") {
        const p = st.getElementsByTagName("style:table-column-properties")[0];
        const w = odsLenToPx(p?.getAttribute("style:column-width"));
        if (w) colW.set(name, w);
      } else if (family === "table-row") {
        const p = st.getElementsByTagName("style:table-row-properties")[0];
        const h = odsLenToPx(p?.getAttribute("style:row-height"));
        if (h) rowH.set(name, h);
      }
    }
  }
  const ownStyle = (el: Element): CellStyle => {
    const s: CellStyle = {};
    const cp = el.getElementsByTagName("style:table-cell-properties")[0];
    if (cp) {
      const bg = odsColorOf(cp.getAttribute("fo:background-color"));
      if (bg) s.bg = bg;
      const all = cp.getAttribute("fo:border");
      const sides: CellStyle["borders"] = {};
      const top = odsBorderColor(cp.getAttribute("fo:border-top") ?? all);
      const right = odsBorderColor(cp.getAttribute("fo:border-right") ?? all);
      const bottom = odsBorderColor(cp.getAttribute("fo:border-bottom") ?? all);
      const left = odsBorderColor(cp.getAttribute("fo:border-left") ?? all);
      if (top) sides.top = top;
      if (right) sides.right = right;
      if (bottom) sides.bottom = bottom;
      if (left) sides.left = left;
      if (top || right || bottom || left) s.borders = sides;
    }
    const tp = el.getElementsByTagName("style:text-properties")[0];
    if (tp) {
      if (tp.getAttribute("fo:font-weight") === "bold") s.bold = true;
      if (tp.getAttribute("fo:font-style") === "italic") s.italic = true;
      const col = odsColorOf(tp.getAttribute("fo:color"));
      if (col) s.color = col;
    }
    const pp = el.getElementsByTagName("style:paragraph-properties")[0];
    const ta = pp?.getAttribute("fo:text-align");
    if (ta === "center") s.align = "center";
    else if (ta === "end" || ta === "right") s.align = "right";
    else if (ta === "start" || ta === "left") s.align = "left";
    return s;
  };
  const cell = new Map<string, CellStyle>();
  const resolve = (name: string, depth = 0): CellStyle => {
    const cached = cell.get(name);
    if (cached) return cached;
    const entry = raw.get(name);
    if (!entry || depth > 8) return {};
    const base = entry.parent ? resolve(entry.parent, depth + 1) : {};
    const merged = { ...base, ...ownStyle(entry.el) };
    cell.set(name, merged);
    return merged;
  };
  for (const name of raw.keys()) resolve(name);
  return { cell, colW, rowH };
}

function readOds(files: Record<string, Uint8Array>): Workbook {
  const contentFile = files["content.xml"];
  if (!contentFile) throw new Error("not an .ods: content.xml missing");
  const contentDoc = parseXml(contentFile);
  const docs = [contentDoc];
  if (files["styles.xml"]) docs.push(parseXml(files["styles.xml"]));
  const styles = parseOdsStyles(docs);
  const wb: Workbook = { kind: "ods", sheets: [], files, contentDoc, contentPath: "content.xml" };
  for (const table of Array.from(contentDoc.getElementsByTagName("table:table"))) {
    const name = table.getAttribute("table:name") ?? `Sheet${wb.sheets.length + 1}`;
    const sheet: Sheet = { name, cells: new Map(), maxRow: 0, maxCol: 0, tableEl: table };
    readOdsTable(sheet, table, styles);
    wb.sheets.push(sheet);
  }
  return wb;
}

function readOdsTable(sheet: Sheet, table: Element, styles: OdsStyles): void {
  // Column widths: walk <table:table-column> (each may repeat) and map to px.
  const cols = new Map<number, number>();
  let colIdx = 0;
  const collectCols = (parent: Element) => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-column") {
        const rep = Math.max(1, Number(ch.getAttribute("table:number-columns-repeated") || "1"));
        const w = styles.colW.get(ch.getAttribute("table:style-name") ?? "");
        for (let i = 0; i < Math.min(rep, REPEAT_CAP); i++) {
          colIdx++;
          if (w) cols.set(colIdx, w);
        }
      } else if (ch.localName === "table-header-columns" || ch.localName === "table-columns") {
        collectCols(ch);
      }
    }
  };
  collectCols(table);
  if (cols.size) sheet.colWidths = cols;

  let rowNum = 0;
  const rows: Element[] = [];
  const collect = (parent: Element) => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-row") rows.push(ch);
      else if (ch.localName === "table-header-rows" || ch.localName === "table-rows") collect(ch);
    }
  };
  collect(table);
  const rowHeights = new Map<number, number>();
  const rowStyles = new Map<number, string>();
  const merges: { r1: number; c1: number; r2: number; c2: number }[] = [];
  for (const rowEl of rows) {
    const rrep = Math.max(1, Number(rowEl.getAttribute("table:number-rows-repeated") || "1"));
    const rowStyle = rowEl.getAttribute("table:style-name") ?? undefined;
    const rh = styles.rowH.get(rowStyle ?? "");
    const parsedCells = parseOdsRow(rowEl, styles);
    const rowHasContent = parsedCells.some((c) => c.has);
    const copies = rowHasContent ? Math.min(rrep, REPEAT_CAP) : 0;
    for (let k = 0; k < copies; k++) {
      const r = rowNum + 1 + k;
      if (rh) rowHeights.set(r, rh);
      if (rowStyle) rowStyles.set(r, rowStyle);
      for (const pc of parsedCells) {
        if (!pc.has) continue;
        const c = pc.cell!;
        if ((pc.colSpan ?? 1) > 1 || (pc.rowSpan ?? 1) > 1) {
          merges.push({ r1: r, c1: c.col, r2: r + (pc.rowSpan ?? 1) - 1, c2: c.col + (pc.colSpan ?? 1) - 1 });
        }
        for (let j = 0; j < pc.span; j++) {
          sheet.cells.set(key(r, c.col + j), { ...c, row: r, col: c.col + j });
          noteExtent(sheet, r, c.col + j);
        }
      }
    }
    rowNum += rrep;
  }
  if (rowHeights.size) sheet.rowHeights = rowHeights;
  if (rowStyles.size) sheet.odsRowStyles = rowStyles;
  if (merges.length) sheet.merges = merges;
}

interface ParsedOdsCell {
  has: boolean;
  span: number;
  colSpan?: number;
  rowSpan?: number;
  cell?: Cell;
}

function parseOdsRow(rowEl: Element, styles: OdsStyles): ParsedOdsCell[] {
  const out: ParsedOdsCell[] = [];
  let col = 0;
  for (const cellEl of Array.from(rowEl.children)) {
    const local = cellEl.localName;
    if (local !== "table-cell" && local !== "covered-table-cell") continue;
    const crep = Math.max(1, Number(cellEl.getAttribute("table:number-columns-repeated") || "1"));
    const startCol = col + 1;
    col += crep;
    if (local === "covered-table-cell") continue; // merged-away cell
    const colSpan = Math.max(1, Number(cellEl.getAttribute("table:number-columns-spanned") || "1"));
    const rowSpan = Math.max(1, Number(cellEl.getAttribute("table:number-rows-spanned") || "1"));
    const valueType = cellEl.getAttribute("office:value-type");
    const formulaRaw = cellEl.getAttribute("table:formula") ?? undefined;
    const style = cellEl.getAttribute("table:style-name") ?? undefined;
    const text = odsCellText(cellEl);
    let value = "";
    let kind: CellKind = "blank";
    let display: string | undefined;
    if (valueType === "float" || valueType === "percentage" || valueType === "currency") {
      value = cellEl.getAttribute("office:value") ?? text;
      // ODF stores the producer's formatted text in <text:p>; use it as the display.
      if (text !== "" && text !== value) display = text;
      kind = "n";
    } else if (valueType === "boolean") {
      value = cellEl.getAttribute("office:boolean-value") === "true" ? "TRUE" : "FALSE";
      kind = "b";
    } else if (valueType === "string") {
      value = cellEl.getAttribute("office:string-value") ?? odsCellText(cellEl);
      kind = "s";
    } else if (valueType === "date") {
      value = cellEl.getAttribute("office:date-value") ?? text;
      if (text !== "" && text !== value) display = text;
      kind = "s";
    } else if (valueType === "time") {
      value = cellEl.getAttribute("office:time-value") ?? text;
      if (text !== "" && text !== value) display = text;
      kind = "s";
    } else {
      value = odsCellText(cellEl);
      kind = value === "" ? "blank" : "s";
    }
    const has = value !== "" || formulaRaw != null || style != null;
    if (!has) {
      out.push({ has: false, span: crep });
      continue;
    }
    const cell: Cell = {
      row: 0,
      col: startCol,
      value,
      kind,
      display,
      formula: formulaRaw ? odfToA1(formulaRaw) : undefined,
      odfFormula: formulaRaw,
      style,
      cellStyle: style ? styles.cell.get(style) : undefined,
      el: cellEl,
    };
    out.push({ has: true, span: Math.min(crep, REPEAT_CAP), colSpan, rowSpan, cell });
  }
  return out;
}

function makeOdsCell(doc: Document, cell: Cell, edited: boolean): Element {
  // Untouched cell: clone the original verbatim (preserves dates, formats, rich text).
  if (cell.el && !cell.edited && !cell.recomputed) {
    const clone = cell.el.cloneNode(true) as Element;
    clone.removeAttribute("table:number-columns-repeated");
    clone.removeAttribute("table:number-rows-repeated");
    return clone;
  }
  const c = doc.createElementNS(ODS.table, "table:table-cell");
  if (cell.style) c.setAttributeNS(ODS.table, "table:style-name", cell.style);
  const formulaToWrite = edited && cell.formula != null ? a1ToOdf(cell.formula) : cell.odfFormula;
  if (formulaToWrite) c.setAttributeNS(ODS.table, "table:formula", formulaToWrite);
  const addText = (text: string) => {
    if (text === "") return;
    const p = doc.createElementNS(ODS.text, "text:p");
    p.textContent = text;
    c.appendChild(p);
  };
  if (cell.kind === "n") {
    c.setAttributeNS(ODS.office, "office:value-type", "float");
    c.setAttributeNS(ODS.office, "office:value", cell.value);
    addText(cell.value);
  } else if (cell.kind === "b") {
    c.setAttributeNS(ODS.office, "office:value-type", "boolean");
    c.setAttributeNS(ODS.office, "office:boolean-value", cell.value === "TRUE" ? "true" : "false");
    addText(cell.value);
  } else if (cell.kind === "s" || cell.kind === "e") {
    c.setAttributeNS(ODS.office, "office:value-type", "string");
    c.setAttributeNS(ODS.office, "office:string-value", cell.value);
    addText(cell.value);
  }
  return c;
}

// --- ods style write-back -------------------------------------------------

function ensureOdsAutoStyles(doc: Document): Element {
  let el = doc.getElementsByTagName("office:automatic-styles")[0] as Element | undefined;
  if (!el) {
    el = doc.createElementNS(ODS.office, "office:automatic-styles");
    const body = doc.getElementsByTagName("office:body")[0];
    doc.documentElement.insertBefore(el, body ?? null);
  }
  return el;
}

function findOdsStyleByName(doc: Document, name: string): Element | undefined {
  for (const s of Array.from(doc.getElementsByTagName("style:style")))
    if (s.getAttribute("style:name") === name) return s;
  return undefined;
}

// Add a built style element to <office:automatic-styles>, reusing an existing one with
// the same family + serialized properties. Returns the (existing or new) style name.
function internOdsStyle(doc: Document, autoStyles: Element, family: string, prefix: string, styleEl: Element): string {
  styleEl.setAttributeNS(ODS.style, "style:family", family);
  const sig = Array.from(styleEl.children).map(xmlOf).join("");
  for (const ex of Array.from(autoStyles.children)) {
    if (ex.localName !== "style" || ex.getAttribute("style:family") !== family) continue;
    if (Array.from(ex.children).map(xmlOf).join("") === sig) return ex.getAttribute("style:name")!;
  }
  const used = new Set(Array.from(doc.getElementsByTagName("style:style")).map((s) => s.getAttribute("style:name")));
  let n = 1;
  while (used.has(prefix + n)) n++;
  const name = prefix + n;
  styleEl.setAttributeNS(ODS.style, "style:name", name);
  autoStyles.appendChild(styleEl);
  return name;
}

const odsSetOrRemove = (el: Element, qn: string, v: string | undefined) => {
  if (v == null) el.removeAttribute(qn);
  else el.setAttributeNS(ODS.fo, qn, v);
};

// Apply a resolved CellStyle onto an ods cell style element (cloned from the original
// so number formats / parents survive), creating the property children as needed.
function applyCellStyleToOds(doc: Document, st: Element, cs: CellStyle): void {
  const child = (tag: string): Element => {
    const ex = st.getElementsByTagName(tag)[0];
    if (ex) return ex;
    const el = doc.createElementNS(ODS.style, tag);
    st.appendChild(el);
    return el;
  };
  const cp = child("style:table-cell-properties");
  odsSetOrRemove(cp, "fo:background-color", cs.bg);
  cp.removeAttribute("fo:border"); // use per-side so partial borders are exact
  const bv = (c?: string) => (c ? `0.5pt solid ${c}` : undefined);
  odsSetOrRemove(cp, "fo:border-top", bv(cs.borders?.top));
  odsSetOrRemove(cp, "fo:border-right", bv(cs.borders?.right));
  odsSetOrRemove(cp, "fo:border-bottom", bv(cs.borders?.bottom));
  odsSetOrRemove(cp, "fo:border-left", bv(cs.borders?.left));
  const tp = child("style:text-properties");
  odsSetOrRemove(tp, "fo:font-weight", cs.bold ? "bold" : undefined);
  odsSetOrRemove(tp, "fo:font-style", cs.italic ? "italic" : undefined);
  odsSetOrRemove(tp, "fo:color", cs.color);
  const pp = child("style:paragraph-properties");
  odsSetOrRemove(pp, "fo:text-align", cs.align === "center" ? "center" : cs.align === "right" ? "end" : cs.align === "left" ? "start" : undefined);
  // Drop property children that ended up empty so dedup stays tight.
  for (const el of [cp, tp, pp]) if (el.attributes.length === 0 && el.children.length === 0) st.removeChild(el);
}

export function setOdsCellStyle(wb: Workbook, _sheet: Sheet, cell: Cell, change: StyleChange): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  const autoStyles = ensureOdsAutoStyles(doc);
  const desired = mergeCellStyle(cell.cellStyle ?? {}, change);
  const orig = cell.style ? findOdsStyleByName(doc, cell.style) : undefined;
  const st = orig
    ? (orig.cloneNode(true) as Element)
    : doc.createElementNS(ODS.style, "style:style");
  st.removeAttribute("style:name");
  applyCellStyleToOds(doc, st, desired);
  cell.style = internOdsStyle(doc, autoStyles, "table-cell", "ce", st);
  cell.cellStyle = desired;
  cell.edited = true;
}

// Build (or reuse) a table-column style of the given width and return its name.
function odsColStyle(doc: Document, autoStyles: Element, px: number): string {
  const st = doc.createElementNS(ODS.style, "style:style");
  const p = doc.createElementNS(ODS.style, "style:table-column-properties");
  p.setAttributeNS(ODS.fo, "fo:break-before", "auto");
  p.setAttributeNS(ODS.style, "style:column-width", `${(px / 96) * 2.54}cm`);
  st.appendChild(p);
  return internOdsStyle(doc, autoStyles, "table-column", "co", st);
}

// Set one column's width (px), splitting the <table:table-column> run that covers it.
export function setOdsColWidth(wb: Workbook, sheet: Sheet, col: number, px: number): void {
  (sheet.colWidths ??= new Map()).set(col, px);
  const doc = wb.contentDoc;
  const table = sheet.tableEl;
  if (!doc || !table) return;
  const styleName = odsColStyle(doc, ensureOdsAutoStyles(doc), px);
  // Walk the column elements, tracking the running column index, and split the run at `col`.
  let idx = 0;
  for (const ch of Array.from(table.children)) {
    if (ch.localName !== "table-column") continue;
    const rep = Math.max(1, Number(ch.getAttribute("table:number-columns-repeated") || "1"));
    const start = idx + 1,
      end = idx + rep;
    idx = end;
    if (col < start || col > end) continue;
    const mk = (from: number, to: number, style?: string) => {
      const c = ch.cloneNode(false) as Element;
      const n = to - from + 1;
      if (n > 1) c.setAttributeNS(ODS.table, "table:number-columns-repeated", String(n));
      else c.removeAttribute("table:number-columns-repeated");
      if (style) c.setAttributeNS(ODS.table, "table:style-name", style);
      return c;
    };
    const parent = ch.parentNode!;
    if (start < col) parent.insertBefore(mk(start, col - 1), ch);
    parent.insertBefore(mk(col, col, styleName), ch);
    if (end > col) parent.insertBefore(mk(col + 1, end), ch);
    parent.removeChild(ch);
    return;
  }
}

// Set one row's height (px) by giving its row a row style (re-emitted by writeOds).
export function setOdsRowHeight(wb: Workbook, sheet: Sheet, row: number, px: number): void {
  (sheet.rowHeights ??= new Map()).set(row, px);
  const doc = wb.contentDoc;
  if (!doc) return;
  const st = doc.createElementNS(ODS.style, "style:style");
  const p = doc.createElementNS(ODS.style, "style:table-row-properties");
  p.setAttributeNS(ODS.style, "style:row-height", `${(px / 96) * 2.54}cm`);
  p.setAttributeNS(ODS.fo, "fo:break-before", "auto");
  st.appendChild(p);
  const name = internOdsStyle(doc, ensureOdsAutoStyles(doc), "table-row", "ro", st);
  (sheet.odsRowStyles ??= new Map()).set(row, name);
}

// Add or remove a merged range; writeOds emits the spans and covered cells.
export function setOdsMerge(sheet: Sheet, r1: number, c1: number, r2: number, c2: number, merge: boolean): void {
  const top = Math.min(r1, r2),
    left = Math.min(c1, c2),
    bottom = Math.max(r1, r2),
    right = Math.max(c1, c2);
  const merges = (sheet.merges ??= []);
  const idx = merges.findIndex((m) => m.r1 === top && m.c1 === left && m.r2 === bottom && m.c2 === right);
  if (merge) {
    if (idx === -1) merges.push({ r1: top, c1: left, r2: bottom, c2: right });
    // Mark covered cells edited so writeOds regenerates the row with covered-table-cells.
    for (let r = top; r <= bottom; r++)
      for (let c = left; c <= right; c++) {
        const cell = sheet.cells.get(key(r, c));
        if (cell) cell.edited = true;
      }
    const tl = sheet.cells.get(key(top, left));
    if (tl) tl.edited = true;
  } else if (idx !== -1) {
    merges.splice(idx, 1);
  }
  const tl = sheet.cells.get(key(top, left));
  if (tl) tl.edited = true;
}

function writeOds(wb: Workbook): void {
  const doc = wb.contentDoc!;
  for (const sheet of wb.sheets) {
    const table = sheet.tableEl;
    if (!table) continue;
    // preserve structural children (column definitions etc.), drop existing rows
    const keep: Element[] = [];
    for (const ch of Array.from(table.children)) {
      if (ch.localName !== "table-row" && ch.localName !== "table-header-rows" && ch.localName !== "table-rows") {
        keep.push(ch);
      }
    }
    while (table.firstChild) table.removeChild(table.firstChild);
    for (const k of keep) table.appendChild(k);
    // Merge bookkeeping: covered positions and the span at each top-left.
    const covered = new Set<string>();
    const spanAt = new Map<string, { cs: number; rs: number }>();
    let mergeMaxCol = 0;
    for (const m of sheet.merges ?? []) {
      spanAt.set(key(m.r1, m.c1), { cs: m.c2 - m.c1 + 1, rs: m.r2 - m.r1 + 1 });
      mergeMaxCol = Math.max(mergeMaxCol, m.c2);
      for (let r = m.r1; r <= m.r2; r++)
        for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(key(r, c));
    }
    const maxRow = Math.max(1, sheet.maxRow);
    const maxCol = Math.max(1, sheet.maxCol, mergeMaxCol);
    for (let r = 1; r <= maxRow; r++) {
      const rowEl = doc.createElementNS(ODS.table, "table:table-row");
      const rowStyle = sheet.odsRowStyles?.get(r);
      if (rowStyle) rowEl.setAttributeNS(ODS.table, "table:style-name", rowStyle);
      // The last column needing an explicit cell: content or a merge edge on this row.
      let lastContent = 0;
      for (let c = maxCol; c >= 1; c--) {
        if (getCell(sheet, r, c) || covered.has(key(r, c)) || spanAt.has(key(r, c))) {
          lastContent = c;
          break;
        }
      }
      for (let c = 1; c <= lastContent; c++) {
        if (covered.has(key(r, c))) {
          rowEl.appendChild(doc.createElementNS(ODS.table, "table:covered-table-cell"));
          continue;
        }
        const cell = getCell(sheet, r, c);
        const el = cell ? makeOdsCell(doc, cell, !!cell.edited) : doc.createElementNS(ODS.table, "table:table-cell");
        const span = spanAt.get(key(r, c));
        if (span) {
          if (span.cs > 1) el.setAttributeNS(ODS.table, "table:number-columns-spanned", String(span.cs));
          if (span.rs > 1) el.setAttributeNS(ODS.table, "table:number-rows-spanned", String(span.rs));
        }
        rowEl.appendChild(el);
      }
      if (lastContent < maxCol) {
        const filler = doc.createElementNS(ODS.table, "table:table-cell");
        filler.setAttributeNS(ODS.table, "table:number-columns-repeated", String(maxCol - lastContent));
        rowEl.appendChild(filler);
      }
      table.appendChild(rowEl);
    }
  }
  wb.files["content.xml"] = serializeXml(doc);
}

// ---------------------------------------------------------------------------
// Recalc engine (shared)
// ---------------------------------------------------------------------------

interface FNode {
  sheet: Sheet;
  cell: Cell;
  id: string;
  deps: Set<string>;
}

function applyResult(cell: Cell, res: unknown): void {
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

/** Recompute every formula cell's cached value, in dependency order. */
export function recalc(wb: Workbook): void {
  const FP = FormulaParser as unknown as {
    new (config: unknown): { parse(f: string, pos: unknown): unknown };
    DepParser: new (config: unknown) => { parse(f: string, pos: unknown): Array<Record<string, unknown>> };
  };
  const byName = new Map<string, Sheet>();
  for (const s of wb.sheets) byName.set(s.name, s);
  const defaultSheet = wb.sheets[0]?.name;
  const lookup = (sheetName: string | undefined, r: number, c: number): Cell | undefined => {
    const sheet = (sheetName && byName.get(sheetName)) || (defaultSheet ? byName.get(defaultSheet) : undefined);
    return sheet ? getCell(sheet, r, c) : undefined;
  };

  const nodes: FNode[] = [];
  const index = new Map<string, FNode>();
  const idOf = (sheetName: string, r: number, c: number) => `${sheetName} ${r}:${c}`;
  for (const sheet of wb.sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.formula == null) continue;
      const node: FNode = { sheet, cell, id: idOf(sheet.name, cell.row, cell.col), deps: new Set() };
      nodes.push(node);
      index.set(node.id, node);
    }
  }
  if (!nodes.length) return;

  const depParser = new FP.DepParser({ onVariable: () => null });
  for (const node of nodes) {
    let refs: Array<Record<string, unknown>> = [];
    try {
      refs = depParser.parse(node.cell.formula!, { row: node.cell.row, col: node.cell.col, sheet: node.sheet.name });
    } catch {
      refs = [];
    }
    for (const ref of refs) {
      const sName = (ref.sheet as string) ?? node.sheet.name;
      if (ref.from) {
        const from = ref.from as { row: number; col: number };
        const to = ref.to as { row: number; col: number };
        for (const other of nodes) {
          if (other.sheet.name !== sName) continue;
          const { row, col } = other.cell;
          if (row >= from.row && row <= to.row && col >= from.col && col <= to.col) node.deps.add(other.id);
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
    for (const node of nodes) if (!seen.has(node.id)) order.push(node);
  }

  const parser = new FP({
    onCell: (ref: { sheet?: string; row: number; col: number }) => typedValue(lookup(ref.sheet, ref.row, ref.col)),
    onRange: (ref: { sheet?: string; from: { row: number; col: number }; to: { row: number; col: number } }) => {
      const out: unknown[][] = [];
      for (let r = ref.from.row; r <= ref.to.row; r++) {
        const rowArr: unknown[] = [];
        for (let c = ref.from.col; c <= ref.to.col; c++) rowArr.push(typedValue(lookup(ref.sheet, r, c)));
        out.push(rowArr);
      }
      return out;
    },
  });
  for (const node of order) {
    let res: unknown;
    try {
      res = parser.parse(node.cell.formula!, { row: node.cell.row, col: node.cell.col, sheet: node.sheet.name });
    } catch {
      continue; // unsupported function / parse error: keep the file's cached value
    }
    // A fresh recompute can error on blank inputs (e.g. DATEDIF on an empty date) even
    // though the file holds a valid cached result; keep that result rather than show an error.
    const isErr = res != null && typeof res === "object" && !Array.isArray(res);
    if (isErr && node.cell.value !== "" && node.cell.kind !== "e") continue;
    applyResult(node.cell, res);
  }
}

// ---------------------------------------------------------------------------
// Public read / write
// ---------------------------------------------------------------------------

export function readWorkbook(bytes: Uint8Array): Workbook {
  const files = unzipSync(bytes);
  if (files["xl/workbook.xml"]) return readXlsx(files);
  if (files["content.xml"]) {
    const mt = files["mimetype"] ? strFromU8(files["mimetype"]) : "";
    if (!mt || mt.includes("spreadsheet")) return readOds(files);
  }
  throw new Error("unrecognized workbook: expected .xlsx or .ods");
}

export function writeWorkbook(wb: Workbook): Uint8Array {
  recalc(wb);
  if (wb.kind === "xlsx") {
    writeXlsx(wb);
    return zipSync(wb.files);
  }
  writeOds(wb);
  // ODF requires the "mimetype" entry first and stored (uncompressed).
  const repacked: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {};
  if (wb.files["mimetype"]) repacked["mimetype"] = [wb.files["mimetype"], { level: 0 }];
  for (const [name, data] of Object.entries(wb.files)) {
    if (name === "mimetype") continue;
    repacked[name] = data;
  }
  return zipSync(repacked as Record<string, Uint8Array>);
}

/** Commit a raw grid edit (the text a user typed) into the model. */
export function setCellInput(sheet: Sheet, row: number, col: number, raw: string): void {
  const existing = getCell(sheet, row, col);
  if (raw.startsWith("=")) {
    const cell = ensureCell(sheet, row, col);
    cell.formula = raw.slice(1).trim();
    cell.odfFormula = undefined;
    cell.edited = true;
    return;
  }
  if (existing == null && raw === "") return;
  const cell = ensureCell(sheet, row, col);
  cell.formula = undefined;
  cell.odfFormula = undefined;
  cell.edited = true;
  if (raw === "") {
    cell.value = "";
    cell.kind = "blank";
  } else if (isNumeric(raw)) {
    cell.value = raw.trim();
    cell.kind = "n";
  } else if (raw.toUpperCase() === "TRUE" || raw.toUpperCase() === "FALSE") {
    cell.value = raw.toUpperCase();
    cell.kind = "b";
  } else {
    cell.value = raw;
    cell.kind = "s";
  }
  // Re-apply the cell's number format (a typed value keeps the cell's format, like Excel).
  cell.display =
    cell.kind === "n" && cell.numFmt != null ? formatNumber(cell.numFmt, cell.value) ?? undefined : undefined;
}

// ---------------------------------------------------------------------------
// Editor
// ---------------------------------------------------------------------------

export interface SheetEditorOptions {
  onChange?: () => void;
}
export interface SheetEditor {
  getBytes(): Promise<Uint8Array>;
  isDirty(): boolean;
  destroy(): void;
}

const ROWS_MIN = 24;
const COLS_MIN = 12;
const ROWS_CAP = 5000;
const COLS_CAP = 256;
const ROW_CHUNK = 20; // rows added per "+ Row" click
const COL_CHUNK = 6; // columns added per "+ Col" click
const STYLE_ID = "sheetedit-style";

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-wrap { display:flex; flex-direction:column; height:100%; background:#1f2227; color:#e6e6e6; font:13px system-ui, sans-serif; }
    .sheetedit-toolbar { display:flex; flex-wrap:wrap; align-items:center; gap:5px; padding:5px 8px; background:#2b2f36; border-bottom:1px solid #1c1f24; }
    .sheetedit-btn {
      font:inherit; font-size:13px; background:#3a3f47; color:#e6e6e6; border:1px solid #4a4f57;
      border-radius:6px; padding:4px 9px; cursor:pointer; min-width:32px; line-height:1.1;
    }
    .sheetedit-btn:hover { background:#454b54; }
    .sheetedit-btn:focus-visible { outline:2px solid #6e7bff; outline-offset:1px; }
    .sheetedit-tb-sep { width:1px; align-self:stretch; background:#4a4f57; margin:1px 3px; }
    .sheetedit-color { width:30px; height:28px; padding:0; border:1px solid #4a4f57; border-radius:6px; background:#3a3f47; cursor:pointer; }
    .sheetedit-btn svg { display:block; width:16px; height:16px; }
    .sheetedit-table th.colhead, .sheetedit-table th.rownum, .sheetedit-table th.corner { cursor:pointer; }
    .sheetedit-table th.colhead:hover, .sheetedit-table th.rownum:hover, .sheetedit-table th.corner:hover { background:#e3e3e8; }
    .sheetedit-table td.sheetedit-sel input { background:rgba(110,123,255,0.18); }
    .sheetedit-pop { position:fixed; z-index:30; background:#2b2f36; border:1px solid #4a4f57; border-radius:8px; padding:4px; box-shadow:0 6px 18px rgba(0,0,0,0.45); display:flex; flex-direction:column; min-width:130px; }
    .sheetedit-pop-item { font:inherit; font-size:13px; text-align:left; background:transparent; color:#e6e6e6; border:0; border-radius:5px; padding:7px 11px; cursor:pointer; }
    .sheetedit-pop-item:hover { background:#3a3f47; }
    /* The grid is a light canvas (like a real spreadsheet) so the file's fills and
       font colours render faithfully and stay readable; the chrome stays dark. */
    .sheetedit-grid { flex:1; min-height:0; overflow:auto; background:#e9e9ec; }
    table.sheetedit-table { border-collapse:collapse; table-layout:fixed; font:13px/1.3 ui-sans-serif, system-ui, sans-serif; }
    .sheetedit-table th, .sheetedit-table td { padding:0; margin:0; }
    .sheetedit-table th { border:1px solid #d4d4d8; }
    /* Cell gridlines as box-shadows (not borders) so the file's own borders sit flush
       against their neighbours and touch, like a real spreadsheet. */
    .sheetedit-table td { background:#fff; box-shadow: inset -1px -1px 0 0 #e3e3e6; }
    .sheetedit-table th {
      position:sticky; top:0; z-index:2; background:#f1f1f4; color:#555; font-weight:600;
      padding:3px 8px; text-align:center; user-select:none;
    }
    .sheetedit-table th.corner { left:0; z-index:3; }
    .sheetedit-table th.rownum { position:sticky; left:0; z-index:1; top:auto; text-align:right; background:#f1f1f4; }
    /* Resize grips: a thin strip on the header border, wide enough to grab on touch. */
    .sheetedit-colgrip { position:absolute; top:0; right:-4px; width:9px; height:100%; cursor:col-resize; z-index:4; touch-action:none; }
    .sheetedit-rowgrip { position:absolute; left:0; bottom:-4px; width:100%; height:9px; cursor:row-resize; z-index:4; touch-action:none; }
    .sheetedit-colgrip:hover { box-shadow:inset -2px 0 0 0 #6e7bff; }
    .sheetedit-rowgrip:hover { box-shadow:inset 0 -2px 0 0 #6e7bff; }
    .sheetedit-table input {
      border:0; background:transparent; color:#1a1a1a; font:inherit; padding:3px 8px;
      width:100%; box-sizing:border-box; outline:none;
    }
    .sheetedit-table td.num input { text-align:right; font-variant-numeric:tabular-nums; }
    .sheetedit-table input:focus { box-shadow:inset 0 0 0 2px #6e7bff; background:#eef0ff; }
    .sheetedit-tabs { display:flex; align-items:center; gap:2px; padding:5px 8px; background:#2b2f36; border-top:1px solid #1c1f24; overflow-x:auto; }
    .sheetedit-tab {
      font:inherit; background:#3a3f47; color:#cfd3da; border:1px solid #4a4f57; border-bottom:none;
      border-radius:5px 5px 0 0; padding:4px 12px; cursor:pointer; white-space:nowrap;
    }
    .sheetedit-tab[aria-selected="true"] { background:#6e7bff; color:#fff; border-color:#6e7bff; }
    .sheetedit-tab:focus-visible { outline:2px solid #fff; outline-offset:1px; }
  `;
  document.head.appendChild(s);
}

export function createSheetEditor(
  container: HTMLElement,
  bytes: Uint8Array,
  options: SheetEditorOptions = {},
): SheetEditor {
  const original = bytes.slice();
  let dirty = false;
  injectStyles();

  const wb = readWorkbook(bytes);
  // Trust the file's cached results on open (like Excel/LibreOffice); recalc only runs
  // after an edit. Recomputing on load would overwrite valid cached values whose inputs
  // are blank in this session (e.g. a DATEDIF age before a birthdate is entered).

  const wrap = document.createElement("div");
  wrap.className = "sheetedit-wrap";
  const toolbar = document.createElement("div");
  toolbar.className = "sheetedit-toolbar";
  const gridScroll = document.createElement("div");
  gridScroll.className = "sheetedit-grid";
  const tabs = document.createElement("div");
  tabs.className = "sheetedit-tabs";
  tabs.setAttribute("role", "tablist");
  tabs.setAttribute("aria-label", t("sheets"));
  wrap.append(toolbar, gridScroll, tabs);
  container.appendChild(wrap);

  let active = 0;
  let inputs = new Map<string, HTMLInputElement>();
  let tds = new Map<string, HTMLElement>();
  // Extra rows/columns the user added beyond the sheet's used extent (per active sheet).
  let extraRows = 0;
  let extraCols = 0;
  // Selection rectangle (1-based, inclusive) and the anchor for shift-extend.
  let sel: { r1: number; c1: number; r2: number; c2: number } | null = null;
  let anchor: { r: number; c: number } | null = null;

  const paintSel = () => {
    for (const td of tds.values()) td.classList.remove("sheetedit-sel");
    if (!sel) return;
    for (let r = sel.r1; r <= sel.r2; r++)
      for (let c = sel.c1; c <= sel.c2; c++) tds.get(key(r, c))?.classList.add("sheetedit-sel");
  };
  const setSel = (r1: number, c1: number, r2: number, c2: number) => {
    sel = { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
    paintSel();
  };
  const selectCell = (r: number, c: number, extend: boolean) => {
    if (extend && anchor) setSel(anchor.r, anchor.c, r, c);
    else {
      anchor = { r, c };
      setSel(r, c, r, c);
    }
  };

  // Rectangular range selection by dragging: mouse drag, or touch long-press then drag.
  // A plain tap still focuses a cell for editing; header/corner taps select a whole line.
  const cellAtPoint = (x: number, y: number): { r: number; c: number } | null => {
    const el = document.elementFromPoint(x, y) as HTMLElement | null;
    const td = el?.closest("td") as HTMLElement | null;
    const rc = td?.dataset.rc;
    if (!rc) return null;
    const [r, c] = rc.split(":").map(Number);
    return { r, c };
  };
  let dragAnchor: { r: number; c: number } | null = null;
  let dragActive = false;
  let justDragged = false;
  let lpTimer: number | null = null;
  let lpStart: { x: number; y: number } | null = null;
  gridScroll.addEventListener("pointerdown", (e) => {
    if (resizing) return;
    const cell = cellAtPoint(e.clientX, e.clientY);
    if (!cell) return;
    if (e.pointerType === "touch") {
      lpStart = { x: e.clientX, y: e.clientY };
      lpTimer = window.setTimeout(() => {
        lpTimer = null;
        dragActive = true;
        dragAnchor = cell;
        anchor = cell;
        (document.activeElement as HTMLElement | null)?.blur?.();
        setSel(cell.r, cell.c, cell.r, cell.c);
      }, 250);
    } else {
      dragAnchor = cell; // mouse: a click still edits; a drag selects
    }
  });
  gridScroll.addEventListener(
    "pointermove",
    (e) => {
      if (e.pointerType === "touch") {
        if (lpTimer != null && lpStart) {
          if (Math.abs(e.clientX - lpStart.x) > 8 || Math.abs(e.clientY - lpStart.y) > 8) {
            clearTimeout(lpTimer); // moved before the long-press fired: it is a scroll
            lpTimer = null;
          }
          return;
        }
        if (!dragActive || !dragAnchor) return;
        const cell = cellAtPoint(e.clientX, e.clientY);
        if (cell) {
          e.preventDefault();
          setSel(dragAnchor.r, dragAnchor.c, cell.r, cell.c);
        }
      } else {
        if (!dragAnchor || e.buttons === 0) return;
        const cell = cellAtPoint(e.clientX, e.clientY);
        if (!cell || (cell.r === dragAnchor.r && cell.c === dragAnchor.c && !dragActive)) return;
        if (!dragActive) {
          dragActive = true;
          anchor = dragAnchor;
          (document.activeElement as HTMLElement | null)?.blur?.();
        }
        e.preventDefault();
        setSel(dragAnchor.r, dragAnchor.c, cell.r, cell.c);
      }
    },
    { passive: false },
  );
  const endDrag = () => {
    if (lpTimer != null) {
      clearTimeout(lpTimer);
      lpTimer = null;
    }
    dragAnchor = null;
    if (dragActive) {
      justDragged = true; // swallow the trailing tap so it does not enter edit mode
      setTimeout(() => (justDragged = false), 350);
    }
    dragActive = false;
  };
  window.addEventListener("pointerup", endDrag);
  window.addEventListener("pointercancel", endDrag);

  // Column-width / row-height drag from the header borders. `resizing` suppresses the
  // cell drag-select while a resize is in progress.
  let resizing = false;
  const startColResize = (e: PointerEvent, col: number, colEl: HTMLElement, startW: number) => {
    e.preventDefault();
    e.stopPropagation();
    resizing = true;
    const x0 = e.clientX;
    const onMove = (ev: PointerEvent) => {
      const w = Math.max(24, Math.round(startW + (ev.clientX - x0)));
      colEl.style.width = `${w}px`;
    };
    const onUp = (ev: PointerEvent) => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      const w = Math.max(24, Math.round(startW + (ev.clientX - x0)));
      const sheet = wb.sheets[active];
      if (sheet && wb.kind === "xlsx") {
        setXlsxColWidth(sheet, col, w);
        mark();
      } else if (sheet && wb.kind === "ods") {
        setOdsColWidth(wb, sheet, col, w);
        mark();
      } else if (sheet) {
        (sheet.colWidths ??= new Map()).set(col, w);
      }
      renderGrid();
      setTimeout(() => (resizing = false), 0);
    };
    window.addEventListener("pointermove", onMove, { passive: false });
    window.addEventListener("pointerup", onUp);
  };
  const startRowResize = (e: PointerEvent, row: number, rowEl: HTMLElement, startH: number) => {
    e.preventDefault();
    e.stopPropagation();
    resizing = true;
    const y0 = e.clientY;
    const onMove = (ev: PointerEvent) => {
      const h = Math.max(16, Math.round(startH + (ev.clientY - y0)));
      rowEl.style.height = `${h}px`;
    };
    const onUp = (ev: PointerEvent) => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      const h = Math.max(16, Math.round(startH + (ev.clientY - y0)));
      const sheet = wb.sheets[active];
      if (sheet && wb.kind === "xlsx") {
        setXlsxRowHeight(sheet, row, h);
        mark();
      } else if (sheet && wb.kind === "ods") {
        setOdsRowHeight(wb, sheet, row, h);
        mark();
      } else if (sheet) {
        (sheet.rowHeights ??= new Map()).set(row, h);
      }
      renderGrid();
      setTimeout(() => (resizing = false), 0);
    };
    window.addEventListener("pointermove", onMove, { passive: false });
    window.addEventListener("pointerup", onUp);
  };

  const tbBtn = (label: string, title: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sheetedit-btn";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("click", onClick);
    return b;
  };
  const tbIcon = (svg: string, title: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sheetedit-btn";
    b.innerHTML = svg;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("click", onClick);
    return b;
  };

  // Style of the selection's top-left cell (used to toggle bold/italic/borders).
  const curStyle = () => (sel ? getCell(wb.sheets[active]!, sel.r1, sel.c1)?.cellStyle : undefined);
  // Apply a style change to every cell in the selection (xlsx only), then re-render.
  // Apply a style change to one cell using the active format's style writer.
  const setCellStyle = (sheet: Sheet, cell: Cell, change: StyleChange) => {
    if (wb.kind === "ods") setOdsCellStyle(wb, sheet, cell, change);
    else setXlsxCellStyle(wb, sheet, cell, change);
  };

  const applyStyle = (change: StyleChange) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    let n = 0;
    for (let r = sel.r1; r <= sel.r2 && n < 4000; r++)
      for (let c = sel.c1; c <= sel.c2 && n < 4000; c++, n++) setCellStyle(sheet, ensureCell(sheet, r, c), change);
    mark();
    renderGrid();
  };

  // Apply a border mode across the selection, computing per-cell sides from each cell's
  // position in the rectangle (e.g. "outer" only borders the perimeter).
  type BorderMode = "all" | "outer" | "top" | "bottom" | "left" | "right" | "none";
  const applyBorder = (mode: BorderMode) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const { r1, c1, r2, c2 } = sel;
    let n = 0;
    for (let r = r1; r <= r2 && n < 4000; r++)
      for (let c = c1; c <= c2 && n < 4000; c++, n++) {
        let sides: StyleChange["borderSides"] = {};
        if (mode === "all") sides = { top: true, right: true, bottom: true, left: true };
        else if (mode === "none") sides = { top: false, right: false, bottom: false, left: false };
        else {
          if ((mode === "outer" || mode === "top") && r === r1) sides.top = true;
          if ((mode === "outer" || mode === "bottom") && r === r2) sides.bottom = true;
          if ((mode === "outer" || mode === "left") && c === c1) sides.left = true;
          if ((mode === "outer" || mode === "right") && c === c2) sides.right = true;
        }
        if (Object.keys(sides).length) setCellStyle(sheet, ensureCell(sheet, r, c), { borderSides: sides });
      }
    mark();
    renderGrid();
  };

  let borderPop: HTMLElement | null = null;
  const openBorderPopover = (btn: HTMLElement) => {
    if (borderPop) {
      borderPop.remove();
      borderPop = null;
      return;
    }
    const pop = document.createElement("div");
    pop.className = "sheetedit-pop";
    const close = () => {
      pop.remove();
      borderPop = null;
      document.removeEventListener("pointerdown", onOutside, true);
    };
    const onOutside = (e: Event) => {
      const tgt = e.target as Node;
      if (!pop.contains(tgt) && !btn.contains(tgt)) close();
    };
    const opts: [string, BorderMode][] = [
      [t("borderAll"), "all"],
      [t("borderOuter"), "outer"],
      [t("borderTop"), "top"],
      [t("borderBottom"), "bottom"],
      [t("borderLeft"), "left"],
      [t("borderRight"), "right"],
      [t("borderNone"), "none"],
    ];
    for (const [label, mode] of opts) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-pop-item";
      b.textContent = label;
      b.addEventListener("click", () => {
        applyBorder(mode);
        close();
      });
      pop.appendChild(b);
    }
    document.body.appendChild(pop);
    borderPop = pop;
    const r = btn.getBoundingClientRect();
    pop.style.left = `${Math.round(r.left)}px`;
    pop.style.top = `${Math.round(r.bottom + 4)}px`;
    setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
  };

  // Merge the selection into one cell, or unmerge when the selection sits on an
  // existing merge. Merging a region first clears any merges it overlaps.
  const setMerge = (sheet: Sheet, mr1: number, mc1: number, mr2: number, mc2: number, on: boolean) => {
    if (wb.kind === "ods") setOdsMerge(sheet, mr1, mc1, mr2, mc2, on);
    else setXlsxMerge(sheet, mr1, mc1, mr2, mc2, on);
  };
  const toggleMerge = () => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const { r1, c1, r2, c2 } = sel;
    const merges = sheet.merges ?? [];
    const within = (m: { r1: number; c1: number; r2: number; c2: number }) =>
      r1 >= m.r1 && c1 >= m.c1 && r2 <= m.r2 && c2 <= m.c2;
    const intersects = (m: { r1: number; c1: number; r2: number; c2: number }) =>
      !(r2 < m.r1 || r1 > m.r2 || c2 < m.c1 || c1 > m.c2);
    const containing = merges.find(within); // selection inside (or equal to) a merge
    if (containing) {
      setMerge(sheet, containing.r1, containing.c1, containing.r2, containing.c2, false);
    } else if (r1 !== r2 || c1 !== c2) {
      for (const m of merges.filter(intersects)) setMerge(sheet, m.r1, m.c1, m.r2, m.c2, false);
      setMerge(sheet, r1, c1, r2, c2, true);
    } else {
      return; // a single, unmerged cell: nothing to do
    }
    mark();
    renderGrid();
  };

  const ICON = {
    left: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M2 8h7M2 12h10"/></svg>`,
    center: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M4.5 8h7M3 12h10"/></svg>`,
    right: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M7 8h7M4 12h10"/></svg>`,
    borders: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4"><rect x="2" y="2" width="12" height="12"/><path d="M8 2v12M2 8h12"/></svg>`,
    merge: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><rect x="1.5" y="4" width="13" height="8"/><path d="M5 6 7.5 8 5 10M11 6 8.5 8 11 10"/></svg>`,
  };
  const buildToolbar = () => {
    toolbar.innerHTML = "";
    const sep = () => {
      const d = document.createElement("div");
      d.className = "sheetedit-tb-sep";
      return d;
    };
    const colorInput = (title: string, def: string, apply: (v: string) => void) => {
      const i = document.createElement("input");
      i.type = "color";
      i.title = title;
      i.setAttribute("aria-label", title);
      i.className = "sheetedit-color";
      i.value = def;
      i.addEventListener("change", () => apply(i.value));
      return i;
    };
    toolbar.append(
      tbBtn(t("addRow"), t("addRows"), () => {
        extraRows += ROW_CHUNK;
        renderGrid();
      }),
      tbBtn(t("addCol"), t("addCols"), () => {
        extraCols += COL_CHUNK;
        renderGrid();
      }),
    );
    if (wb.kind !== "xlsx" && wb.kind !== "ods") return; // styling needs a known style model
    const bold = tbBtn("B", t("bold"), () => applyStyle({ bold: !curStyle()?.bold }));
    bold.style.fontWeight = "700";
    const italic = tbBtn("I", t("italic"), () => applyStyle({ italic: !curStyle()?.italic }));
    italic.style.fontStyle = "italic";
    toolbar.append(
      sep(),
      bold,
      italic,
      colorInput(t("textColour"), "#000000", (v) => applyStyle({ color: v })),
      colorInput(t("fillColour"), "#ffff00", (v) => applyStyle({ bg: v })),
      sep(),
      tbIcon(ICON.left, t("alignLeft"), () => applyStyle({ align: "left" })),
      tbIcon(ICON.center, t("alignCentre"), () => applyStyle({ align: "center" })),
      tbIcon(ICON.right, t("alignRight"), () => applyStyle({ align: "right" })),
      sep(),
    );
    const borderBtn = tbIcon(ICON.borders, t("borders"), () => openBorderPopover(borderBtn));
    toolbar.append(borderBtn, tbIcon(ICON.merge, t("merge"), toggleMerge));
  };

  const mark = () => {
    if (!dirty) {
      dirty = true;
    }
    options.onChange?.();
  };

  const displayValue = (sheet: Sheet, r: number, c: number): string => cellDisplay(getCell(sheet, r, c));

  const refreshDisplays = (sheet: Sheet, except?: HTMLInputElement) => {
    for (const [k, input] of inputs) {
      if (input === except) continue;
      const [r, c] = k.split(":").map(Number);
      input.value = displayValue(sheet, r!, c!);
      const cell = getCell(sheet, r!, c!);
      input.parentElement?.classList.toggle("num", cell?.kind === "n");
    }
  };

  const renderGrid = () => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    inputs = new Map();
    tds = new Map();
    gridScroll.innerHTML = "";
    const rows = Math.min(ROWS_CAP, Math.max(ROWS_MIN, sheet.maxRow + 6) + extraRows);
    const cols = Math.min(COLS_CAP, Math.max(COLS_MIN, sheet.maxCol + 2) + extraCols);

    const table = document.createElement("table");
    table.className = "sheetedit-table";

    // Column widths (table-layout is fixed, so these are authoritative). The table is
    // sized to the sum so columns keep their width and the grid scrolls horizontally,
    // rather than the table shrinking to the viewport and squashing every column.
    const colgroup = document.createElement("colgroup");
    const rnCol = document.createElement("col");
    rnCol.style.width = "44px";
    colgroup.appendChild(rnCol);
    let totalW = 44;
    const colEls: HTMLElement[] = [];
    for (let c = 1; c <= cols; c++) {
      const w = sheet.colWidths?.get(c) ?? 96;
      const col = document.createElement("col");
      col.style.width = `${w}px`;
      colgroup.appendChild(col);
      colEls.push(col);
      totalW += w;
    }
    table.appendChild(colgroup);
    table.style.width = `${totalW}px`;

    const head = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "corner";
    corner.title = t("selectAll");
    corner.addEventListener("click", () => setSel(1, 1, rows, cols));
    head.appendChild(corner);
    for (let c = 1; c <= cols; c++) {
      const th = document.createElement("th");
      th.className = "colhead";
      th.textContent = colToLetters(c);
      th.title = t("selectColumn", { col: colToLetters(c) });
      th.addEventListener("click", () => {
        if (resizing) return;
        anchor = { r: 1, c };
        setSel(1, c, rows, c);
      });
      const grip = document.createElement("div");
      grip.className = "sheetedit-colgrip";
      const colEl = colEls[c - 1]!;
      grip.addEventListener("pointerdown", (e) =>
        startColResize(e, c, colEl, sheet.colWidths?.get(c) ?? 96),
      );
      th.appendChild(grip);
      head.appendChild(th);
    }
    table.appendChild(head);

    // Merged ranges: the top-left cell spans; covered cells are not rendered.
    const covered = new Set<string>();
    const spanAt = new Map<string, { rs: number; cs: number }>();
    for (const m of sheet.merges ?? []) {
      spanAt.set(key(m.r1, m.c1), { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
      for (let r = m.r1; r <= m.r2; r++)
        for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(key(r, c));
    }

    for (let r = 1; r <= rows; r++) {
      const tr = document.createElement("tr");
      const rh = sheet.rowHeights?.get(r);
      if (rh) tr.style.height = `${rh}px`;
      const rn = document.createElement("th");
      rn.className = "rownum";
      rn.textContent = String(r);
      rn.title = t("selectRow", { row: r });
      rn.addEventListener("click", () => {
        if (resizing) return;
        anchor = { r, c: 1 };
        setSel(r, 1, r, cols);
      });
      const rgrip = document.createElement("div");
      rgrip.className = "sheetedit-rowgrip";
      rgrip.addEventListener("pointerdown", (e) => startRowResize(e, r, tr, rh ?? 22));
      rn.appendChild(rgrip);
      tr.appendChild(rn);
      for (let c = 1; c <= cols; c++) {
        if (covered.has(key(r, c))) continue; // part of a merge; the top-left cell spans it
        const td = document.createElement("td");
        td.dataset.rc = key(r, c);
        tds.set(key(r, c), td);
        const sp = spanAt.get(key(r, c));
        if (sp) {
          if (sp.rs > 1) td.rowSpan = sp.rs;
          if (sp.cs > 1) td.colSpan = sp.cs;
        }
        const cell = getCell(sheet, r, c);
        if (cell?.kind === "n") td.classList.add("num");
        const input = document.createElement("input");
        input.type = "text";
        input.value = cellDisplay(cell);
        input.setAttribute("aria-label", `${colToLetters(c)}${r}`);
        // Apply the file's visual style (fill/borders on the cell, font/colour/align on the text).
        const cs = cell?.cellStyle;
        if (cs) {
          if (cs.bg) td.style.background = cs.bg;
          if (cs.borders) {
            // Override the default gridline box-shadow: keep light right/bottom unless the
            // file specifies a border there, and add the file's top/left where present.
            const bd = cs.borders;
            const g = "#e3e3e6";
            const sh = [`inset -1px 0 0 0 ${bd.right ?? g}`, `inset 0 -1px 0 0 ${bd.bottom ?? g}`];
            if (bd.top) sh.push(`inset 0 1px 0 0 ${bd.top}`);
            if (bd.left) sh.push(`inset 1px 0 0 0 ${bd.left}`);
            td.style.boxShadow = sh.join(", ");
          }
          if (cs.bold) input.style.fontWeight = "700";
          if (cs.italic) input.style.fontStyle = "italic";
          if (cs.color) input.style.color = cs.color;
          if (cs.align) input.style.textAlign = cs.align;
        }
        const ki = key(r, c);
        // Shift-click extends the selection from the anchor (no caret/edit).
        input.addEventListener("mousedown", (e) => {
          if (e.shiftKey) {
            e.preventDefault();
            selectCell(r, c, true);
          }
        });
        input.addEventListener("focus", () => {
          if (justDragged) {
            input.blur(); // a range was just drag-selected; do not enter edit on the trailing tap
            return;
          }
          selectCell(r, c, false); // tapping a cell selects it; toolbar styles target the selection
          const live = getCell(sheet, r, c);
          if (!live) return;
          // Show the editable underlying value (formula or raw), not the formatted display.
          input.value = live.formula != null ? "=" + live.formula : live.value;
        });
        const commit = () => {
          const raw = input.value;
          const live = getCell(sheet, r, c);
          const before = live ? (live.formula != null ? "=" + live.formula : live.value) : "";
          if (raw === before) {
            input.value = displayValue(sheet, r, c);
            return;
          }
          setCellInput(sheet, r, c, raw);
          recalc(wb);
          mark();
          refreshDisplays(sheet);
          input.value = displayValue(sheet, r, c);
          input.parentElement?.classList.toggle("num", getCell(sheet, r, c)?.kind === "n");
        };
        input.addEventListener("blur", commit);
        input.addEventListener("keydown", (e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            input.blur();
            const below = inputs.get(key(r + 1, c));
            below?.focus();
          } else if (e.key === "Escape") {
            input.value = displayValue(sheet, r, c);
            input.blur();
          }
        });
        td.appendChild(input);
        tr.appendChild(td);
        inputs.set(ki, input);
      }
      table.appendChild(tr);
    }
    gridScroll.appendChild(table);
    paintSel(); // restore the selection highlight after a re-render
  };

  const renderTabs = () => {
    tabs.innerHTML = "";
    wb.sheets.forEach((sheet, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-tab";
      b.textContent = sheet.name;
      b.setAttribute("role", "tab");
      b.setAttribute("aria-selected", String(i === active));
      b.addEventListener("click", () => {
        if (i === active) return;
        active = i;
        extraRows = 0; // each sheet starts at its own extent
        extraCols = 0;
        sel = null;
        anchor = null;
        renderTabs();
        renderGrid();
      });
      tabs.appendChild(b);
    });
  };

  buildToolbar();
  renderTabs();
  renderGrid();

  return {
    isDirty() {
      return dirty;
    },
    async getBytes() {
      return dirty ? writeWorkbook(wb) : original.slice();
    },
    destroy() {
      wrap.remove();
    },
  };
}
