import type { Cell, Sheet } from "../../core/model";
import { colToLetters, parseA1Ref } from "../../core/model";
// Shared xlsx (OOXML SpreadsheetML) constants and DOM helpers used by the
// read, write and style modules.

export const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export function ensureXlsxCellEl(sheet: Sheet, cell: Cell): Element {
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

export const xmlOf = (el: Element): string => new XMLSerializer().serializeToString(el);

