import type { Cell, Sheet, StyleChange, Workbook } from "../../core/model";
import { firstByLocal, mergeCellStyle } from "../../core/model";
import { SS_MAIN, ensureXlsxCellEl, xmlOf } from "./shared";
// ---------------------------------------------------------------------------
// xlsx styles: style-pool management for user style changes
// ---------------------------------------------------------------------------

export function poolIndex(parent: Element, candidate: Element): number {
  const want = xmlOf(candidate);
  const kids = Array.from(parent.children);
  for (let i = 0; i < kids.length; i++) if (xmlOf(kids[i]!) === want) return i;
  parent.appendChild(candidate);
  parent.setAttribute("count", String(parent.children.length));
  return parent.children.length - 1;
}

export const argbOf = (css: string): string => "FF" + css.replace("#", "").toUpperCase();

// Compute the resulting CellStyle from the current one plus a change. Shared by the

/**
 * Apply a style change to a cell, managing the xlsx style pools: derive a new font /
 * fill / border from the cell's current format plus the change, find-or-create each in
 * styles.xml, find-or-create the combined <xf>, and point the cell at it.
 */
export function setXlsxCellStyle(wb: Workbook, sheet: Sheet, cell: Cell, change: StyleChange): void {
  const doc = wb.stylesDoc;
  if (!doc || !sheet.doc || !sheet.sheetData) return; // chartsheets have no cells to style
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
