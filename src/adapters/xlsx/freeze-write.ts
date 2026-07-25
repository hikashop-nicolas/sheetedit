import { colToLetters, type Sheet, type Workbook } from "../../core/model";

// Freeze panes: <sheetView><pane xSplit ySplit topLeftCell state="frozen"/>. xSplit / ySplit are
// the counts of frozen leading columns / rows, topLeftCell is the first cell of the scrolling pane,
// and activePane names which pane the selection lives in. Clearing a freeze removes the element.
// The rest of the sheetView (zoom, gridlines, the selection) is left as it was.

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
  if (cols > 0) pane.setAttribute("xSplit", String(cols)); else pane.removeAttribute("xSplit");
  if (rows > 0) pane.setAttribute("ySplit", String(rows)); else pane.removeAttribute("ySplit");
  pane.setAttribute("topLeftCell", `${colToLetters(cols + 1)}${rows + 1}`);
  pane.setAttribute("activePane", activePaneOf(rows, cols));
  pane.setAttribute("state", "frozen");
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
