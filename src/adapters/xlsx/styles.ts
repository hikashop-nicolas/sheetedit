import { strToU8 } from "fflate";
import type { Cell, Sheet, StyleChange, Workbook } from "../../core/model";
import { firstByLocal, formatNumber, mergeCellStyle, parseXmlOpt, serializeXml } from "../../core/model";
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
  const underline = change.underline ?? cur.underline;
  const strike = change.strike ?? cur.strike;
  const fontSize = change.fontSize ?? cur.fontSize;
  const fontFamily = change.fontFamily ?? cur.fontFamily;
  const color = change.color ?? cur.color;
  const bg = change.bg ?? cur.bg;
  const align = change.align ?? cur.align;
  const valign = change.valign ?? cur.valign;
  const wrap = change.wrap ?? cur.wrap;
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
  flag("u", underline);
  flag("strike", strike);
  // Size/name: set when known, never remove (undefined = inherit the font's own).
  const val = (tag: string, v: string | undefined) => {
    if (v == null) return;
    const el = firstByLocal(font, tag) ?? (font.appendChild(ce(tag)) as Element);
    el.setAttribute("val", v);
  };
  val("sz", fontSize != null ? String(fontSize) : undefined);
  val("name", fontFamily);
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
  if (align || valign || wrap) {
    xf.setAttribute("applyAlignment", "1");
    const a = ce("alignment");
    if (align) a.setAttribute("horizontal", align);
    if (valign) a.setAttribute("vertical", valign === "middle" ? "center" : valign);
    if (wrap) a.setAttribute("wrapText", "1");
    xf.appendChild(a);
  }
  const sIdx = poolIndex(cellXfsEl, xf);

  cell.style = String(sIdx);
  ensureXlsxCellEl(sheet, cell).setAttribute("s", String(sIdx));
  cell.cellStyle = mergeCellStyle(cur, change);
  cell.edited = true;
  wb.stylesDirty = true;
}

// Minimal styles.xml (default font/fills/border/xf), created when a workbook has
// none (e.g. one produced by the CSV converter) and a number format must persist.
const MINIMAL_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<styleSheet xmlns="${SS_MAIN}"><fonts count="1"><font/></fonts>` +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>';

export function ensureXlsxStylesDoc(wb: Workbook): Document | undefined {
  if (wb.stylesDoc) return wb.stylesDoc;
  wb.stylesDoc = parseXmlOpt(strToU8(MINIMAL_STYLES));
  if (!wb.stylesDoc) return undefined;
  wb.files["xl/styles.xml"] = strToU8(MINIMAL_STYLES);
  // Wire the new part in: content type + workbook relationship.
  const ct = wb.files["[Content_Types].xml"];
  const ctDoc = ct ? parseXmlOpt(ct) : undefined;
  if (ctDoc) {
    const has = Array.from(ctDoc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === "/xl/styles.xml");
    if (!has) {
      const o = ctDoc.createElementNS(ctDoc.documentElement.namespaceURI, "Override");
      o.setAttribute("PartName", "/xl/styles.xml");
      o.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
      ctDoc.documentElement.appendChild(o);
      wb.files["[Content_Types].xml"] = serializeXml(ctDoc);
    }
  }
  const relsPath = "xl/_rels/workbook.xml.rels";
  const relsDoc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]!) : undefined;
  if (relsDoc) {
    const rels = Array.from(relsDoc.getElementsByTagName("Relationship"));
    const has = rels.some((r) => (r.getAttribute("Target") ?? "").endsWith("styles.xml"));
    if (!has) {
      let n = rels.length + 1;
      while (rels.some((r) => r.getAttribute("Id") === `rId${n}`)) n++;
      const r = relsDoc.createElementNS(relsDoc.documentElement.namespaceURI, "Relationship");
      r.setAttribute("Id", `rId${n}`);
      r.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
      r.setAttribute("Target", "styles.xml");
      relsDoc.documentElement.appendChild(r);
      wb.files[relsPath] = serializeXml(relsDoc);
    }
  }
  return wb.stylesDoc;
}

// Format code -> numFmtId: reuse a matching custom <numFmt>, else mint one (>= 164).
function numFmtIdFor(doc: Document, fmt: string): number {
  const root = doc.documentElement;
  let numFmts = firstByLocal(root, "numFmts");
  if (!numFmts) {
    numFmts = doc.createElementNS(root.namespaceURI || SS_MAIN, "numFmts");
    root.insertBefore(numFmts, firstByLocal(root, "fonts") ?? root.firstElementChild);
  }
  let maxId = 163;
  for (const nf of Array.from(numFmts.children)) {
    if (nf.localName !== "numFmt") continue;
    const id = Number(nf.getAttribute("numFmtId") || "0");
    if (nf.getAttribute("formatCode") === fmt) return id;
    if (id > maxId) maxId = id;
  }
  const nf = doc.createElementNS(root.namespaceURI || SS_MAIN, "numFmt");
  nf.setAttribute("numFmtId", String(maxId + 1));
  nf.setAttribute("formatCode", fmt);
  numFmts.appendChild(nf);
  numFmts.setAttribute("count", String(numFmts.children.length));
  return maxId + 1;
}

/**
 * Point a cell at a number format (code, built-in id, or undefined = General),
 * cloning its current <xf> so fonts/fills/borders stay, and interning the result.
 */
export function setXlsxCellNumFmt(wb: Workbook, sheet: Sheet, cell: Cell, fmt: string | number | undefined): void {
  const doc = ensureXlsxStylesDoc(wb);
  if (!doc || !sheet.doc || !sheet.sheetData) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const root = doc.documentElement;
  const cellXfsEl = firstByLocal(root, "cellXfs") ?? (root.appendChild(doc.createElementNS(ns, "cellXfs")) as Element);
  const id = fmt == null ? 0 : typeof fmt === "number" ? fmt : numFmtIdFor(doc, fmt);
  const cur = cellXfsEl.children[cell.style ? Number(cell.style) : 0];
  const xf = cur ? (cur.cloneNode(true) as Element) : doc.createElementNS(ns, "xf");
  if (!cur) {
    xf.setAttribute("fontId", "0");
    xf.setAttribute("fillId", "0");
    xf.setAttribute("borderId", "0");
    xf.setAttribute("xfId", "0");
  }
  xf.setAttribute("numFmtId", String(id));
  if (id) xf.setAttribute("applyNumberFormat", "1");
  else xf.removeAttribute("applyNumberFormat");
  const sIdx = poolIndex(cellXfsEl, xf);
  cell.style = String(sIdx);
  ensureXlsxCellEl(sheet, cell).setAttribute("s", String(sIdx));
  cell.numFmt = fmt;
  cell.numFmtDirty = false;
  cell.display = cell.kind === "n" && fmt != null ? formatNumber(fmt, cell.value) ?? undefined : undefined;
  cell.edited = true;
  wb.stylesDirty = true;
}
