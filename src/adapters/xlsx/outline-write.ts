import { type Sheet, type Workbook } from "../../core/model";

// Write outline (grouping) levels back into the worksheet: @outlineLevel / @collapsed / @hidden on
// each <row>, and a <col> span per column level. Rows and columns outside any group have the
// attributes cleared, so ungrouping really removes them rather than leaving a stale level behind.
// The summary side (<sheetPr><outlinePr>) is written only when it differs from Excel's default.

const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function ensureRowEl(sheet: Sheet, row: number): Element | undefined {
  const doc = sheet.doc, sd = sheet.sheetData;
  if (!doc || !sd) return undefined;
  for (const re of Array.from(sd.children)) if (re.localName === "row" && Number(re.getAttribute("r") || "0") === row) return re;
  const rowEl = doc.createElementNS(doc.documentElement.namespaceURI || SS_MAIN, "row");
  rowEl.setAttribute("r", String(row));
  let next: Element | null = null;
  for (const re of Array.from(sd.children)) if (re.localName === "row" && Number(re.getAttribute("r") || "0") > row) { next = re; break; }
  sd.insertBefore(rowEl, next);
  return rowEl;
}

/** Rewrite one row's outline attributes. Level 0 clears them. */
export function setXlsxRowOutline(sheet: Sheet, row: number, level: number, collapsed: boolean, hidden: boolean): void {
  const rowEl = ensureRowEl(sheet, row);
  if (!rowEl) return;
  if (level > 0) rowEl.setAttribute("outlineLevel", String(Math.min(7, level)));
  else rowEl.removeAttribute("outlineLevel");
  if (collapsed) rowEl.setAttribute("collapsed", "1"); else rowEl.removeAttribute("collapsed");
  if (hidden) rowEl.setAttribute("hidden", "1"); else rowEl.removeAttribute("hidden");
  sheet.layoutDirty = true;
}

/** The <cols> container, created in schema position (after sheetFormatPr, before sheetData). */
function ensureColsEl(sheet: Sheet): Element | undefined {
  const doc = sheet.doc;
  const ws = doc?.documentElement;
  if (!doc || !ws) return undefined;
  const existing = Array.from(ws.children).find((e) => e.localName === "cols");
  if (existing) return existing;
  const cols = doc.createElementNS(ws.namespaceURI || SS_MAIN, "cols");
  ws.insertBefore(cols, sheet.sheetData ?? null);
  return cols;
}

/**
 * Rewrite the <cols> spans so each column's width / hidden / outline state matches the model.
 * Excel stores columns as ranges, so this emits one <col> per column and lets Excel re-coalesce;
 * a sheet with no column state at all keeps its original <cols> untouched.
 */
export function writeXlsxColOutline(sheet: Sheet): void {
  const doc = sheet.doc;
  if (!doc || !sheet.colOutline) return;
  const cols = ensureColsEl(sheet);
  if (!cols) return;
  // Start from what the file already says per column, so widths and other attributes survive.
  const attrs = new Map<number, Record<string, string>>();
  for (const col of Array.from(cols.children)) {
    if (col.localName !== "col") continue;
    const min = Number(col.getAttribute("min") || "0");
    const max = Number(col.getAttribute("max") || String(min));
    if (!min) continue;
    const base: Record<string, string> = {};
    for (const a of Array.from(col.attributes)) if (a.name !== "min" && a.name !== "max") base[a.name] = a.value;
    for (let c = min; c <= Math.min(max, min + 16383); c++) attrs.set(c, { ...base });
  }
  const touched = new Set<number>([...attrs.keys(), ...sheet.colOutline.keys(), ...(sheet.hiddenCols ?? []), ...(sheet.colCollapsed ?? [])]);
  for (const c of touched) {
    const a = attrs.get(c) ?? {};
    const level = sheet.colOutline.get(c) ?? 0;
    if (level > 0) a.outlineLevel = String(Math.min(7, level)); else delete a.outlineLevel;
    if (sheet.colCollapsed?.has(c)) a.collapsed = "1"; else delete a.collapsed;
    if (sheet.hiddenCols?.has(c)) a.hidden = "1"; else delete a.hidden;
    // <col> needs at least one property besides the range to be worth writing.
    if (Object.keys(a).length) attrs.set(c, a); else attrs.delete(c);
  }
  while (cols.firstChild) cols.removeChild(cols.firstChild);
  for (const c of [...attrs.keys()].sort((x, y) => x - y)) {
    const el = doc.createElementNS(cols.namespaceURI || SS_MAIN, "col");
    el.setAttribute("min", String(c));
    el.setAttribute("max", String(c));
    for (const [k, v] of Object.entries(attrs.get(c)!)) el.setAttribute(k, v);
    cols.appendChild(el);
  }
  if (!cols.childNodes.length) cols.parentNode?.removeChild(cols);
  sheet.layoutDirty = true;
}

/** <sheetPr><outlinePr summaryBelow summaryRight/>, written only when it differs from the default. */
export function writeXlsxOutlinePr(sheet: Sheet): void {
  const doc = sheet.doc;
  const ws = doc?.documentElement;
  if (!doc || !ws) return;
  const below = sheet.summaryBelow ?? true, right = sheet.summaryRight ?? true;
  let sheetPr = Array.from(ws.children).find((e) => e.localName === "sheetPr");
  const existing = sheetPr ? Array.from(sheetPr.children).find((e) => e.localName === "outlinePr") : undefined;
  if (below && right && !existing) return;
  if (!sheetPr) {
    sheetPr = doc.createElementNS(ws.namespaceURI || SS_MAIN, "sheetPr");
    ws.insertBefore(sheetPr, ws.firstChild);
  }
  const pr = existing ?? doc.createElementNS(ws.namespaceURI || SS_MAIN, "outlinePr");
  pr.setAttribute("summaryBelow", below ? "1" : "0");
  pr.setAttribute("summaryRight", right ? "1" : "0");
  if (!existing) sheetPr.insertBefore(pr, sheetPr.firstChild);
  sheet.layoutDirty = true;
}

/** Persist every sheet's outline state (called from the save path). */
export function writeXlsxOutlines(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    if (!sheet.outlineDirty) continue;
    const maxRow = Math.max(sheet.maxRow, ...(sheet.rowOutline ? [...sheet.rowOutline.keys()] : [0]));
    for (let r = 1; r <= maxRow; r++) {
      const level = sheet.rowOutline?.get(r) ?? 0;
      const collapsed = sheet.rowCollapsed?.has(r) ?? false;
      const hidden = sheet.hiddenRows?.has(r) ?? false;
      // Skip rows the file never mentioned and the model has nothing to say about.
      if (!level && !collapsed && !hidden && !rowExists(sheet, r)) continue;
      setXlsxRowOutline(sheet, r, level, collapsed, hidden);
    }
    writeXlsxColOutline(sheet);
    writeXlsxOutlinePr(sheet);
    sheet.outlineDirty = false;
  }
}

function rowExists(sheet: Sheet, row: number): boolean {
  const sd = sheet.sheetData;
  if (!sd) return false;
  for (const re of Array.from(sd.children)) if (re.localName === "row" && Number(re.getAttribute("r") || "0") === row) return true;
  return false;
}
