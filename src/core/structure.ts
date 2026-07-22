import type { Cell, Sheet, Workbook } from "./model";
import { colToLetters, key, lettersToCol, parseA1Ref, MAX_COL, MAX_ROW } from "./model";

// ---------------------------------------------------------------------------
// Row / column insertion and deletion with reference rewriting.
//
// The op mutates the model (cells map, merges, size maps, extents) and the
// format-specific storage: xlsx sheet XML is shifted surgically (<row r>, <c r>,
// <cols>, mergeCells, dimension); ods only shifts the model and its row metadata
// maps, because writeOds re-emits a dirty sheet's rows from the model. Formulas
// in EVERY sheet that reference the edited sheet are rewritten: refs past the
// edit shift, refs into deleted lines become #REF!, ranges grow and shrink.
// Known limits: row/col-anchored extras (hyperlinks, data validations,
// conditional formats, autofilters) are not re-anchored.
// ---------------------------------------------------------------------------

export interface LineOp {
  axis: "row" | "col";
  kind: "insert" | "delete";
  /** 1-based first line of the op. Insert places new lines BEFORE this index. */
  at: number;
  count: number;
}

/** Map a 1-based line index through the op; null = the line was deleted. */
export function mapPoint(i: number, op: LineOp): number | null {
  if (op.kind === "insert") return i >= op.at ? i + op.count : i;
  if (i < op.at) return i;
  if (i < op.at + op.count) return null;
  return i - op.count;
}

/** Map an inclusive 1-based span through the op; null = entirely deleted. */
export function mapSpan(a: number, b: number, op: LineOp): { a: number; b: number } | null {
  if (op.kind === "insert") {
    return { a: a >= op.at ? a + op.count : a, b: b >= op.at ? b + op.count : b };
  }
  const end = op.at + op.count; // exclusive
  const na = a < op.at ? a : a >= end ? a - op.count : op.at;
  const nb = b < op.at ? b : b >= end ? b - op.count : op.at - 1;
  return nb < na ? null : { a: na, b: nb };
}

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** Shift a rectangle by a line op: lines past the edit move; a range straddling an insertion
    grows; a range fully inside a deletion collapses (null). A resulting single cell is kept
    (unlike merges, secondary ranges may legitimately cover one cell). */
export function shiftRect(m: Rect, op: LineOp): Rect | null {
  const growRow = op.axis === "row" && op.kind === "insert" && m.r1 < op.at && m.r2 >= op.at;
  const growCol = op.axis === "col" && op.kind === "insert" && m.c1 < op.at && m.c2 >= op.at;
  const rs = op.axis === "row" ? (growRow ? { a: m.r1, b: m.r2 + op.count } : mapSpan(m.r1, m.r2, op)) : { a: m.r1, b: m.r2 };
  const cs = op.axis === "col" ? (growCol ? { a: m.c1, b: m.c2 + op.count } : mapSpan(m.c1, m.c2, op)) : { a: m.c1, b: m.c2 };
  if (!rs || !cs) return null;
  return { r1: rs.a, c1: cs.a, r2: rs.b, c2: cs.b };
}

const rectToA1 = (m: Rect): string =>
  m.r1 === m.r2 && m.c1 === m.c2 ? `${colToLetters(m.c1)}${m.r1}` : `${colToLetters(m.c1)}${m.r1}:${colToLetters(m.c2)}${m.r2}`;

function parseRangeA1(range: string): Rect | null {
  const [a, b] = range.split(":");
  const p1 = parseA1Ref(a ?? "");
  const p2 = b ? parseA1Ref(b) : p1;
  if (!p1 || !p2) return null;
  return { r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) };
}

/** Shift every range in a space-separated sqref/ref string; drops collapsed ranges. Ranges that
    do not parse as plain A1 (e.g. a cross-sheet ref) are left untouched. */
export function shiftSqref(sqref: string, op: LineOp): string {
  const out: string[] = [];
  for (const range of sqref.split(/\s+/).filter(Boolean)) {
    const rect = parseRangeA1(range);
    if (!rect) { out.push(range); continue; }
    const s = shiftRect(rect, op);
    if (s) out.push(rectToA1(s));
  }
  return out.join(" ");
}

function shiftSqrefAttr(el: Element, attr: string, op: LineOp): void {
  const cur = el.getAttribute(attr);
  if (!cur) return;
  const next = shiftSqref(cur, op);
  if (next) el.setAttribute(attr, next);
  else el.parentNode?.removeChild(el);
}

// --- formula rewriting -------------------------------------------------------

const NAME_CHAR = /[A-Za-z0-9_$.[\]]/;
// Optional sheet prefix: 'Quoted Name'! (with '' escapes) or an unquoted name!.
const PREFIX = "((?:'(?:[^']|'')*'|[A-Za-z_\\u00C0-\\uFFFF][A-Za-z0-9_.\\u00C0-\\uFFFF]*)!)?";
// A1 or A1:B2, with optional $ anchors (structural ops move absolute refs too).
const RECT_REF = new RegExp(PREFIX + "(\\$?)([A-Za-z]{1,3})(\\$?)(\\d{1,7})(?:(\\s*:\\s*)(\\$?)([A-Za-z]{1,3})(\\$?)(\\d{1,7}))?(?![A-Za-z0-9_(:])", "g");
const COL_RANGE = new RegExp(PREFIX + "(\\$?)([A-Za-z]{1,3})(\\s*:\\s*)(\\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(])", "g");
const ROW_RANGE = new RegExp(PREFIX + "(\\$?)(\\d{1,7})(\\s*:\\s*)(\\$?)(\\d{1,7})(?![A-Za-z0-9_:])", "g");

const unquoteSheet = (prefix: string | undefined): string | null => {
  if (!prefix) return null;
  let name = prefix.slice(0, -1); // drop "!"
  if (name.startsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
  return name;
};

/** Does this occurrence refer to the sheet the op runs on? */
const refersToTarget = (prefix: string | undefined, formulaSheet: string, targetSheet: string): boolean => {
  const name = unquoteSheet(prefix) ?? formulaSheet;
  return name.toLowerCase() === targetSheet.toLowerCase();
};

/** True when the match starts inside a longer identifier (LOG10, tax2026, Table1[...]). */
const insideName = (chunk: string, off: number, prefix: string | undefined): boolean => {
  const startCh = off > 0 ? chunk[off - 1]! : "";
  if (prefix && prefix.startsWith("'")) return false; // quoted prefixes are unambiguous
  return NAME_CHAR.test(startCh);
};

function rewriteChunk(chunk: string, formulaSheet: string, targetSheet: string, op: LineOp): string {
  const refErr = (prefix: string | undefined) => (prefix ?? "") + "#REF!";
  // Rectangular refs and single cells.
  chunk = chunk.replace(RECT_REF, (m, prefix: string | undefined, a1c: string, c1s: string, a1r: string, r1s: string, colon: string | undefined, a2c: string, c2s: string, a2r: string, r2s: string, off: number) => {
    if (insideName(chunk, off, prefix)) return m;
    if (!refersToTarget(prefix, formulaSheet, targetSheet)) return m;
    const c1 = lettersToCol(c1s.toUpperCase());
    const r1 = Number(r1s);
    if (c1 > MAX_COL || r1 > MAX_ROW) return m; // a defined name shaped like a ref
    const pre = prefix ?? "";
    if (!colon) {
      const ni = mapPoint(op.axis === "row" ? r1 : c1, op);
      if (ni == null) return refErr(prefix);
      const nc = op.axis === "col" ? ni : c1;
      const nr = op.axis === "row" ? ni : r1;
      return pre + a1c + colToLetters(nc) + a1r + nr;
    }
    const c2 = lettersToCol(c2s.toUpperCase());
    const r2 = Number(r2s);
    if (c2 > MAX_COL || r2 > MAX_ROW) return m;
    const span = op.axis === "row" ? mapSpan(Math.min(r1, r2), Math.max(r1, r2), op) : mapSpan(Math.min(c1, c2), Math.max(c1, c2), op);
    if (!span) return refErr(prefix);
    // Keep the endpoints' original order and anchors; only their numbers change.
    const lo = op.axis === "row" ? Math.min(r1, r2) : Math.min(c1, c2);
    const m1 = (op.axis === "row" ? r1 : c1) === lo ? span.a : span.b;
    const m2 = (op.axis === "row" ? r2 : c2) === lo ? span.a : span.b;
    const p1 = op.axis === "row" ? a1c + c1s.toUpperCase() + a1r + m1 : a1c + colToLetters(m1) + a1r + r1;
    const p2 = op.axis === "row" ? a2c + c2s.toUpperCase() + a2r + m2 : a2c + colToLetters(m2) + a2r + r2;
    return pre + p1 + colon + p2;
  });
  // Whole-column ranges (A:C), only meaningful for column ops.
  if (op.axis === "col") {
    chunk = chunk.replace(COL_RANGE, (m, prefix: string | undefined, s1: string, l1: string, colon: string, s2: string, l2: string, off: number) => {
      if (insideName(chunk, off, prefix)) return m;
      if (!refersToTarget(prefix, formulaSheet, targetSheet)) return m;
      const c1 = lettersToCol(l1.toUpperCase());
      const c2 = lettersToCol(l2.toUpperCase());
      if (c1 > MAX_COL || c2 > MAX_COL) return m;
      const span = mapSpan(Math.min(c1, c2), Math.max(c1, c2), op);
      if (!span) return (prefix ?? "") + "#REF!";
      const m1 = c1 <= c2 ? span.a : span.b;
      const m2 = c1 <= c2 ? span.b : span.a;
      return (prefix ?? "") + s1 + colToLetters(m1) + colon + s2 + colToLetters(m2);
    });
  }
  // Whole-row ranges (1:5), only meaningful for row ops.
  if (op.axis === "row") {
    chunk = chunk.replace(ROW_RANGE, (m, prefix: string | undefined, s1: string, d1: string, colon: string, s2: string, d2: string, off: number) => {
      if (insideName(chunk, off, prefix)) return m;
      if (!refersToTarget(prefix, formulaSheet, targetSheet)) return m;
      const r1 = Number(d1);
      const r2 = Number(d2);
      if (r1 > MAX_ROW || r2 > MAX_ROW || r1 < 1 || r2 < 1) return m;
      const span = mapSpan(Math.min(r1, r2), Math.max(r1, r2), op);
      if (!span) return (prefix ?? "") + "#REF!";
      const m1 = r1 <= r2 ? span.a : span.b;
      const m2 = r1 <= r2 ? span.b : span.a;
      return (prefix ?? "") + s1 + m1 + colon + s2 + m2;
    });
  }
  return chunk;
}

/**
 * Rewrite the A1 references in a formula for a structural op on `targetSheet`.
 * `formulaSheet` is the sheet the formula lives on (unprefixed refs belong to it).
 * Double-quoted string literals are left untouched.
 */
export function rewriteFormula(formula: string, formulaSheet: string, targetSheet: string, op: LineOp): string {
  let out = "";
  let i = 0;
  while (i < formula.length) {
    if (formula[i] === '"') {
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') j += 2;
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
    let next = formula.indexOf('"', i);
    if (next === -1) next = formula.length;
    out += rewriteChunk(formula.slice(i, next), formulaSheet, targetSheet, op);
    i = next;
  }
  return out;
}

// --- xlsx XML surgery ---------------------------------------------------------

/** Give every row and cell of the sheet an explicit r attribute (implicit
    positional refs would silently mis-place once neighbours shift). */
function materializeXlsxRefs(sheet: Sheet): void {
  const sd = sheet.sheetData;
  if (!sd) return;
  let rowIdx = 0;
  for (const rowEl of Array.from(sd.children)) {
    if (rowEl.localName !== "row") continue;
    rowIdx = Number(rowEl.getAttribute("r")) || rowIdx + 1;
    rowEl.setAttribute("r", String(rowIdx));
    let colIdx = 0;
    for (const c of Array.from(rowEl.children)) {
      if (c.localName !== "c") continue;
      const ref = c.getAttribute("r");
      const parsed = ref ? /^([A-Z]+)(\d+)$/.exec(ref) : null;
      colIdx = parsed ? lettersToCol(parsed[1]!) : colIdx + 1;
      c.setAttribute("r", colToLetters(colIdx) + rowIdx);
    }
  }
}

/** Rebuild <mergeCells> from the model's merge list. */
export function syncXlsxMerges(sheet: Sheet): void {
  const doc = sheet.doc;
  if (!doc || !sheet.sheetData) return;
  const ns = doc.documentElement.namespaceURI;
  let mcEl = doc.getElementsByTagName("mergeCells")[0] as Element | undefined;
  const merges = sheet.merges ?? [];
  if (!merges.length) {
    mcEl?.parentNode?.removeChild(mcEl);
    return;
  }
  if (!mcEl) {
    mcEl = doc.createElementNS(ns, "mergeCells");
    sheet.sheetData.parentNode?.insertBefore(mcEl, sheet.sheetData.nextSibling);
  }
  while (mcEl.firstChild) mcEl.removeChild(mcEl.firstChild);
  for (const m of merges) {
    const mc = doc.createElementNS(ns, "mergeCell");
    mc.setAttribute("ref", `${colToLetters(m.c1)}${m.r1}:${colToLetters(m.c2)}${m.r2}`);
    mcEl.appendChild(mc);
  }
  mcEl.setAttribute("count", String(merges.length));
}

function shiftXlsxXml(sheet: Sheet, op: LineOp): void {
  const sd = sheet.sheetData;
  const doc = sheet.doc;
  if (!sd || !doc) return;
  materializeXlsxRefs(sheet);
  if (op.axis === "row") {
    for (const rowEl of Array.from(sd.children)) {
      if (rowEl.localName !== "row") continue;
      const r = Number(rowEl.getAttribute("r"));
      const nr = mapPoint(r, op);
      if (nr == null) {
        sd.removeChild(rowEl);
        continue;
      }
      if (nr === r) continue;
      rowEl.setAttribute("r", String(nr));
      for (const c of Array.from(rowEl.children)) {
        if (c.localName !== "c") continue;
        const m = /^([A-Z]+)\d+$/.exec(c.getAttribute("r") ?? "");
        if (m) c.setAttribute("r", m[1]! + nr);
      }
    }
  } else {
    for (const rowEl of Array.from(sd.children)) {
      if (rowEl.localName !== "row") continue;
      for (const c of Array.from(rowEl.children)) {
        if (c.localName !== "c") continue;
        const m = /^([A-Z]+)(\d+)$/.exec(c.getAttribute("r") ?? "");
        if (!m) continue;
        const col = lettersToCol(m[1]!);
        const nc = mapPoint(col, op);
        if (nc == null) rowEl.removeChild(c);
        else if (nc !== col) c.setAttribute("r", colToLetters(nc) + m[2]);
      }
      rowEl.removeAttribute("spans"); // stale after a column shift
    }
    // <cols>: shift runs past the op; a run covering an insertion grows.
    const colsEl = doc.getElementsByTagName("cols")[0] as Element | undefined;
    if (colsEl) {
      for (const ce of Array.from(colsEl.children)) {
        if (ce.localName !== "col") continue;
        const min = Number(ce.getAttribute("min") || "0");
        const max = Number(ce.getAttribute("max") || String(min));
        const span = op.kind === "insert" && min < op.at && max >= op.at
          ? { a: min, b: max + op.count } // insertion inside the run: it grows
          : mapSpan(min, max, op);
        if (!span) {
          colsEl.removeChild(ce);
          continue;
        }
        ce.setAttribute("min", String(span.a));
        ce.setAttribute("max", String(span.b));
      }
      if (!colsEl.children.length) colsEl.parentNode?.removeChild(colsEl);
    }
  }
  // Secondary ranges that live in the sheet XML but are otherwise untouched: conditional
  // formatting, data validations, hyperlinks and the autofilter. Shift their sqref/ref so they
  // track the edit instead of going stale.
  for (const el of Array.from(doc.getElementsByTagName("conditionalFormatting"))) shiftSqrefAttr(el, "sqref", op);
  for (const el of Array.from(doc.getElementsByTagName("dataValidation"))) shiftSqrefAttr(el, "sqref", op);
  for (const el of Array.from(doc.getElementsByTagName("hyperlink"))) shiftSqrefAttr(el, "ref", op);
  const af = doc.getElementsByTagName("autoFilter")[0] as Element | undefined;
  if (af) shiftSqrefAttr(af, "ref", op);

  // Refresh <dimension> from the model extents (recomputed by the caller).
  const dim = doc.getElementsByTagName("dimension")[0] as Element | undefined;
  if (dim) dim.setAttribute("ref", `A1:${colToLetters(Math.max(1, sheet.maxCol))}${Math.max(1, sheet.maxRow)}`);
  sheet.layoutDirty = true;
}

// --- ods metadata shifting ------------------------------------------------------

function shiftNumMap<V>(map: Map<number, V> | undefined, op: LineOp): Map<number, V> | undefined {
  if (!map) return map;
  const next = new Map<number, V>();
  for (const [i, v] of map) {
    const ni = mapPoint(i, op);
    if (ni != null) next.set(ni, v);
  }
  return next;
}

function shiftOdsRowMeta(sheet: Sheet, op: LineOp): void {
  if (op.axis !== "row") {
    // Column op: shift the covered-cell keys' columns; row metadata is untouched.
    if (sheet.odsCoveredEls) {
      const next = new Map<string, Element>();
      for (const [k, el] of sheet.odsCoveredEls) {
        const [r, c] = k.split(":").map(Number);
        const nc = mapPoint(c!, op);
        if (nc != null) next.set(key(r!, nc), el);
      }
      sheet.odsCoveredEls = next;
    }
    return;
  }
  sheet.odsRowStyles = shiftNumMap(sheet.odsRowStyles, op);
  sheet.odsRowEls = shiftNumMap(sheet.odsRowEls, op);
  if (sheet.odsCoveredEls) {
    const next = new Map<string, Element>();
    for (const [k, el] of sheet.odsCoveredEls) {
      const [r, c] = k.split(":").map(Number);
      const nr = mapPoint(r!, op);
      if (nr != null) next.set(key(nr, c!), el);
    }
    sheet.odsCoveredEls = next;
  }
  if (sheet.odsHeaderRows) {
    const span = op.kind === "insert" && sheet.odsHeaderRows.from < op.at && sheet.odsHeaderRows.to >= op.at
      ? { a: sheet.odsHeaderRows.from, b: sheet.odsHeaderRows.to + op.count }
      : mapSpan(sheet.odsHeaderRows.from, sheet.odsHeaderRows.to, op);
    if (span) {
      sheet.odsHeaderRows.from = span.a;
      sheet.odsHeaderRows.to = span.b;
    } else {
      sheet.odsHeaderRows = undefined;
    }
  }
  if (sheet.odsRowRuns) {
    const next: { from: number; to: number; el: Element }[] = [];
    for (const run of sheet.odsRowRuns) {
      if (op.kind === "insert") {
        if (op.at <= run.from) next.push({ from: run.from + op.count, to: run.to + op.count, el: run.el });
        else if (op.at <= run.to) {
          // Inserting inside a repeated content run: split it around the new lines.
          next.push({ from: run.from, to: op.at - 1, el: run.el });
          next.push({ from: op.at + op.count, to: run.to + op.count, el: run.el });
        } else next.push(run);
      } else {
        const before = { from: run.from, to: Math.min(run.to, op.at - 1) };
        const after = { from: Math.max(run.from, op.at + op.count) - op.count, to: run.to - op.count };
        if (before.to >= before.from) next.push({ ...before, el: run.el });
        if (after.to >= after.from && run.to >= op.at + op.count) next.push({ ...after, el: run.el });
      }
    }
    sheet.odsRowRuns = next;
  }
}

/** Adjust the ods table's declared columns (table:table-column repeats). */
function shiftOdsColumns(sheet: Sheet, op: LineOp): void {
  const table = sheet.tableEl;
  if (!table || op.axis !== "col") return;
  const decls: Element[] = [];
  const collect = (parent: Element) => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-column") decls.push(ch);
      else if (ch.localName === "table-header-columns" || ch.localName === "table-columns") collect(ch);
    }
  };
  collect(table);
  let idx = 1;
  for (const d of decls) {
    const rep = Number(d.getAttribute("table:number-columns-repeated") || "1");
    const from = idx;
    const to = idx + rep - 1;
    idx = to + 1;
    let nrep = rep;
    if (op.kind === "insert") {
      // The run holding the column just before the insertion grows (at=1 grows the first run).
      const anchor = Math.max(1, op.at - 1);
      if (anchor >= from && anchor <= to) nrep = rep + op.count;
    } else {
      const overlap = Math.max(0, Math.min(to, op.at + op.count - 1) - Math.max(from, op.at) + 1);
      nrep = rep - overlap;
    }
    if (nrep === rep) continue;
    if (nrep <= 0) d.parentNode?.removeChild(d);
    else if (nrep === 1) d.removeAttribute("table:number-columns-repeated");
    else d.setAttribute("table:number-columns-repeated", String(nrep));
  }
}

// --- the workbook-level op ------------------------------------------------------

function recomputeExtent(sheet: Sheet): void {
  let mr = 0;
  let mc = 0;
  for (const cell of sheet.cells.values()) {
    if (cell.row > mr) mr = cell.row;
    if (cell.col > mc) mc = cell.col;
  }
  for (const m of sheet.merges ?? []) {
    if (m.r2 > mr) mr = m.r2;
    if (m.c2 > mc) mc = m.c2;
  }
  for (const run of sheet.odsRowRuns ?? []) if (run.to > mr) mr = run.to;
  sheet.maxRow = mr;
  sheet.maxCol = mc;
}

/**
 * Insert or delete rows/columns on a sheet. Rewrites references across the whole
 * workbook (unless `rewriteRefs` is false, used by undo which restores formulas
 * from a snapshot), shifts the model and the format-specific storage, and marks
 * everything needed for a correct save. Callers recalc and re-render after.
 */
export function applyLineOp(wb: Workbook, sheetIdx: number, op: LineOp, rewriteRefs = true): void {
  const sheet = wb.sheets[sheetIdx];
  if (!sheet || op.count < 1 || op.at < 1) return;

  if (rewriteRefs) {
    for (const s of wb.sheets) {
      for (const cell of s.cells.values()) {
        if (cell.formula == null) continue;
        const nf = rewriteFormula(cell.formula, s.name, sheet.name, op);
        if (nf !== cell.formula) {
          cell.formula = nf;
          cell.odfFormula = undefined;
          cell.fDirty = true;
          cell.edited = true;
        }
      }
    }
  }

  // Shift the cell map; cells on deleted lines drop out.
  const next = new Map<string, Cell>();
  const moved: Cell[] = [];
  for (const cell of sheet.cells.values()) {
    const i = op.axis === "row" ? cell.row : cell.col;
    const ni = mapPoint(i, op);
    if (ni == null) continue;
    if (ni !== i) {
      if (op.axis === "row") cell.row = ni;
      else cell.col = ni;
      moved.push(cell);
    }
    next.set(key(cell.row, cell.col), cell);
  }
  sheet.cells = next;

  if (wb.kind === "xlsx") {
    // Moved formula cells: force a plain rewrite so shared groups de-share safely
    // and the stale calcChain (which stores absolute cell addresses) gets dropped.
    for (const cell of moved) {
      if (cell.formula != null) {
        cell.fDirty = true;
        cell.edited = true;
      }
    }
    for (const cell of sheet.cells.values()) {
      if (cell.sharedSi != null) {
        cell.fDirty = true;
        cell.edited = true;
      }
    }
  }

  // Merges: shift; ranges shrink on delete and grow on inside-insertion; degenerate drop.
  if (sheet.merges?.length) {
    const remapped: { r1: number; c1: number; r2: number; c2: number }[] = [];
    for (const m of sheet.merges) {
      const growRow = op.axis === "row" && op.kind === "insert" && m.r1 < op.at && m.r2 >= op.at;
      const growCol = op.axis === "col" && op.kind === "insert" && m.c1 < op.at && m.c2 >= op.at;
      const rs = op.axis === "row" ? (growRow ? { a: m.r1, b: m.r2 + op.count } : mapSpan(m.r1, m.r2, op)) : { a: m.r1, b: m.r2 };
      const cs = op.axis === "col" ? (growCol ? { a: m.c1, b: m.c2 + op.count } : mapSpan(m.c1, m.c2, op)) : { a: m.c1, b: m.c2 };
      if (!rs || !cs) continue;
      if (rs.a === rs.b && cs.a === cs.b) continue;
      remapped.push({ r1: rs.a, c1: cs.a, r2: rs.b, c2: cs.b });
    }
    sheet.merges = remapped;
  }

  // Data-validation ranges (model, used for rendering the dropdowns) follow the edit too; the
  // xlsx sqref in the sheet XML is shifted separately in shiftXlsxXml.
  if (sheet.validations?.length) {
    sheet.validations = sheet.validations
      .map((v) => ({ ...v, ranges: v.ranges.map((g) => shiftRect(g, op)).filter((g): g is Rect => g != null) }))
      .filter((v) => v.ranges.length > 0);
  }

  // Row heights / column widths follow their lines.
  if (op.axis === "row") sheet.rowHeights = shiftNumMap(sheet.rowHeights, op);
  else sheet.colWidths = shiftNumMap(sheet.colWidths, op);

  recomputeExtent(sheet);

  if (wb.kind === "xlsx") {
    shiftXlsxXml(sheet, op);
    syncXlsxMerges(sheet);
  } else if (wb.kind === "ods") {
    shiftOdsRowMeta(sheet, op);
    shiftOdsColumns(sheet, op);
    sheet.odsDirty = true;
  } else if (sheet.csvRows) {
    if (op.axis === "row") {
      const at = Math.min(op.at - 1, sheet.csvRows.length);
      if (op.kind === "insert") {
        const term = sheet.csvRows.find((r) => r.terminator)?.terminator ?? "\n";
        sheet.csvRows.splice(at, 0, ...Array.from({ length: op.count }, () => ({ raw: "", terminator: term, dirty: true, width: 0 })));
      } else {
        sheet.csvRows.splice(at, op.count);
      }
    } else {
      // A column shift re-frames every line: the whole file re-serializes.
      for (const r of sheet.csvRows) {
        r.dirty = true;
        if (op.kind === "insert") {
          if (r.width >= op.at) r.width += op.count;
        } else {
          const overlap = Math.max(0, Math.min(r.width, op.at + op.count - 1) - op.at + 1);
          r.width -= overlap;
        }
      }
    }
  }
}
