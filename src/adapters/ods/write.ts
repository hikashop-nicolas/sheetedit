import type { Cell, Sheet, Workbook } from "../../core/model";
import { getCell, key, serializeXml } from "../../core/model";
import { ODS, a1ToOdf } from "./shared";
import { ensureOdsAutoStyles, internOdsStyle, odsColStyle } from "./styles";
// ---------------------------------------------------------------------------
// ods write: cell/row emission and the save pass
// ---------------------------------------------------------------------------

export function makeOdsCell(doc: Document, cell: Cell, edited: boolean): Element {
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
  sheet.odsDirty = true;
}

// Add or remove a merged range; writeOds emits the spans and covered cells.
export function setOdsMerge(sheet: Sheet, r1: number, c1: number, r2: number, c2: number, merge: boolean): void {
  sheet.odsDirty = true;
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

export function writeOds(wb: Workbook): void {
  const doc = wb.contentDoc!;
  for (const sheet of wb.sheets) {
    const table = sheet.tableEl;
    if (!table) continue;
    // An untouched sheet keeps its original XML verbatim (repeats, header groups,
    // row attributes, covered cells); only touched sheets are re-emitted.
    let cellsDirty = false;
    for (const cell of sheet.cells.values())
      if (cell.edited || cell.recomputed) {
        cellsDirty = true;
        break;
      }
    if (!cellsDirty && !sheet.odsDirty) continue;
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
    // Repeated content runs beyond the expansion cap, re-emitted verbatim.
    const runAt = new Map<number, { to: number; el: Element }>();
    let lastRow = maxRow;
    for (const rr of sheet.odsRowRuns ?? []) {
      runAt.set(rr.from, { to: rr.to, el: rr.el });
      lastRow = Math.max(lastRow, rr.to);
    }
    // Original header-rows group: its rows are re-wrapped in a fresh group element.
    const hdr = sheet.odsHeaderRows;
    const headerEl = hdr ? (hdr.el.cloneNode(false) as Element) : null;
    const appendRow = (rowEl: Element, r: number) => {
      if (hdr && headerEl && r >= hdr.from && r <= hdr.to) {
        if (!headerEl.parentNode) table.appendChild(headerEl);
        headerEl.appendChild(rowEl);
      } else table.appendChild(rowEl);
    };
    for (let r = 1; r <= lastRow; r++) {
      const run = runAt.get(r);
      if (run) {
        // The preserved tail of a repeated content run: one clone with its repeat count.
        const clone = run.el.cloneNode(true) as Element;
        const n = run.to - r + 1;
        if (n > 1) clone.setAttributeNS(ODS.table, "table:number-rows-repeated", String(n));
        else clone.removeAttribute("table:number-rows-repeated");
        appendRow(clone, r);
        r = run.to;
        continue;
      }
      if (r > maxRow) continue; // gaps past the content extent existed only as empty runs
      const orig = sheet.odsRowEls?.get(r);
      const rowEl = orig ? (orig.cloneNode(false) as Element) : doc.createElementNS(ODS.table, "table:table-row");
      rowEl.removeAttribute("table:number-rows-repeated");
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
          const origCov = sheet.odsCoveredEls?.get(key(r, c));
          const cov = origCov ? (origCov.cloneNode(true) as Element) : doc.createElementNS(ODS.table, "table:covered-table-cell");
          cov.removeAttribute("table:number-columns-repeated");
          rowEl.appendChild(cov);
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
      appendRow(rowEl, r);
    }
  }
  wb.files["content.xml"] = serializeXml(doc);
}
