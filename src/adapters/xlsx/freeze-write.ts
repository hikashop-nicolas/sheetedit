import { colToLetters, type Sheet, type Workbook } from "../../core/model";

// Pane boundaries: <sheetView><pane xSplit ySplit topLeftCell state="frozen"/>. For a FROZEN pane
// xSplit / ySplit are the counts of leading columns / rows; for a draggable SPLIT they are an
// offset in twips (1/20 pt), so the same boundary is written differently per state. topLeftCell is
// the first cell of the trailing pane and activePane names the pane the selection lives in.
// Clearing removes the element; the rest of the sheetView (zoom, gridlines, selection) is kept.

const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** The first <sheetView>, created (with its <sheetViews> parent) when the sheet has none. */
function ensureSheetView(sheet: Sheet): Element | undefined {
  const doc = sheet.doc;
  const ws = doc?.documentElement;
  if (!doc || !ws) return undefined;
  let views = Array.from(ws.children).find((e) => e.localName === "sheetViews");
  if (!views) {
    views = doc.createElementNS(ws.namespaceURI || SS_MAIN, "sheetViews");
    // sheetViews sits after sheetPr / dimension and before sheetFormatPr.
    const after = Array.from(ws.children).find((e) => e.localName !== "sheetPr" && e.localName !== "dimension");
    ws.insertBefore(views, after ?? null);
  }
  let view = Array.from(views.children).find((e) => e.localName === "sheetView");
  if (!view) {
    view = doc.createElementNS(views.namespaceURI || SS_MAIN, "sheetView");
    view.setAttribute("workbookViewId", "0");
    views.appendChild(view);
  }
  return view;
}

const DEFAULT_COL_PX = 64;
const DEFAULT_ROW_PX = 20;
/** Total pixel size of the first `n` lines. */
const sizeSum = (n: number, sizePx: (line: number) => number): number => {
  let px = 0;
  for (let i = 1; i <= n; i++) px += sizePx(i);
  return px;
};
const pxToTwips = (px: number): number => Math.round(px * 15);

/** Which pane the selection belongs to once rows and/or columns are frozen. */
function activePaneOf(rows: number, cols: number): string {
  if (rows > 0 && cols > 0) return "bottomRight";
  return rows > 0 ? "bottomLeft" : "topRight";
}

/** Write (or clear) one sheet's frozen panes from sheet.freeze. */
export function writeXlsxFreeze(sheet: Sheet): void {
  const doc = sheet.doc;
  if (!doc) return;
  const view = ensureSheetView(sheet);
  if (!view) return;
  const existing = Array.from(view.children).find((e) => e.localName === "pane");
  const rows = sheet.freeze?.rows ?? 0, cols = sheet.freeze?.cols ?? 0;
  if (!rows && !cols) {
    if (existing) existing.parentNode?.removeChild(existing);
    // A selection tagged with a pane that no longer exists confuses Excel; drop the pane hint.
    for (const sel of Array.from(view.children).filter((e) => e.localName === "selection")) sel.removeAttribute("pane");
    sheet.layoutDirty = true;
    return;
  }
  const pane = existing ?? doc.createElementNS(view.namespaceURI || SS_MAIN, "pane");
  const split = !!sheet.paneSplit;
  // A split states the boundary as a pixel offset in twips; sum the line sizes up to it.
  const xVal = split ? pxToTwips(sizeSum(cols, (c) => sheet.colWidths?.get(c) ?? DEFAULT_COL_PX)) : cols;
  const yVal = split ? pxToTwips(sizeSum(rows, (r) => sheet.rowHeights?.get(r) ?? DEFAULT_ROW_PX)) : rows;
  if (cols > 0) pane.setAttribute("xSplit", String(xVal)); else pane.removeAttribute("xSplit");
  if (rows > 0) pane.setAttribute("ySplit", String(yVal)); else pane.removeAttribute("ySplit");
  pane.setAttribute("topLeftCell", `${colToLetters(cols + 1)}${rows + 1}`);
  pane.setAttribute("activePane", activePaneOf(rows, cols));
  pane.setAttribute("state", split ? "split" : "frozen");
  // <pane> is the first child of sheetView, before any <selection>.
  if (!existing) view.insertBefore(pane, view.firstChild);
  sheet.layoutDirty = true;
}

/** Persist every sheet whose freeze changed (called from the save path). */
export function writeXlsxFreezes(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    if (!sheet.freezeDirty) continue;
    writeXlsxFreeze(sheet);
    sheet.freezeDirty = false;
  }
}
