import { parseXmlOpt, serializeXml, type Sheet, type SheetShape, type Workbook } from "../../core/model";
import { GEOM_PRESET } from "../../core/shape-geom";
import { pxToEmu } from "../../core/chart-model";
import { ensureSheetDrawing } from "./chart-write";

// Persist authored / moved / resized / restyled drawing shapes into the worksheet drawing part.
// A brand-new shape appends a <xdr:twoCellAnchor><xdr:sp> to the drawing XML (creating one via
// ensureSheetDrawing if needed); an existing shape has its anchor and fill / outline / text patched
// in place, leaving other spPr children (xfrm, effects, custom geometry) untouched.

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const kid = (p: Element, local: string): Element | undefined => Array.from(p.children).find((c) => c.localName === local);
const hex = (c: string): string => c.replace(/^#/, "").toUpperCase();

/** The preset geometry name to write (round-trips the file's original when we kept it). */
function prstOf(sh: SheetShape): string {
  return sh.preset ?? GEOM_PRESET[sh.geom] ?? "rect";
}

/**
 * Inner spPr / txBody markup shared by the new-shape and restyle paths.
 *
 * A restyle replaces the whole <a:ln> and the whole fill, so everything they carried has to be
 * written back out of the model: without that, changing the colour of a dashed arrow turned it
 * into a solid line with no arrowhead, and a translucent watermark came back opaque.
 */
function spPrXml(sh: SheetShape): string {
  // An opaque fill stays the plain self-closing element the format normally carries.
  const clr = sh.fillOpacity != null && sh.fillOpacity < 1
    ? `<a:srgbClr val="${hex(sh.fill ?? "")}"><a:alpha val="${Math.round(sh.fillOpacity * 100000)}"/></a:srgbClr>`
    : `<a:srgbClr val="${hex(sh.fill ?? "")}"/>`;
  const fill = sh.fill ? `<a:solidFill>${clr}</a:solidFill>` : `<a:noFill/>`;
  const dash = sh.dash ? `<a:prstDash val="${sh.dash}"/>` : "";
  const ends = `${sh.headEnd ? `<a:headEnd type="${sh.headEnd}"/>` : ""}${sh.tailEnd ? `<a:tailEnd type="${sh.tailEnd}"/>` : ""}`;
  const ln = sh.stroke
    ? `<a:ln w="${pxToEmu(sh.strokeWidth ?? 1)}"><a:solidFill><a:srgbClr val="${hex(sh.stroke)}"/></a:solidFill>${dash}${ends}</a:ln>`
    : `<a:ln><a:noFill/></a:ln>`;
  // <a:avLst> carries the geometry's adjustment (where an elbow bends, how wide a brace is).
  const adj = sh.adjust != null ? `<a:gd name="adj1" fmla="val ${Math.round(sh.adjust * 100000)}"/>` : "";
  return `<a:prstGeom prst="${prstOf(sh)}"><a:avLst>${adj}</a:avLst></a:prstGeom>${fill}${ln}`;
}

/** <a:xfrm> for a shape that carries a turn or a mirror, so an edit does not straighten it. */
function xfrmXml(sh: SheetShape): string {
  if (!sh.rotation && !sh.flipH && !sh.flipV) return "";
  const rot = sh.rotation ? ` rot="${Math.round(sh.rotation * 60000)}"` : "";
  const flip = `${sh.flipH ? ' flipH="1"' : ""}${sh.flipV ? ' flipV="1"' : ""}`;
  // No <a:off>/<a:ext> of our own: the restyle path carries the file's across, and a shape with
  // neither takes its frame from the anchor. Writing an offset of zero would move the shape.
  return `<a:xfrm${rot}${flip}/>`;
}
function txBodyXml(sh: SheetShape): string {
  if (!sh.text) return `<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p/></xdr:txBody>`;
  const t = sh.text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const clr = sh.textColor ? `<a:solidFill><a:srgbClr val="${hex(sh.textColor)}"/></a:solidFill>` : "";
  return `<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US">${clr}</a:rPr><a:t>${t}</a:t></a:r></a:p></xdr:txBody>`;
}

function anchorPointsXml(sh: SheetShape): string {
  const a = sh.anchor;
  const pt = (tag: string, col: number, colOff: number, row: number, rowOff: number): string =>
    `<xdr:${tag}><xdr:col>${col - 1}</xdr:col><xdr:colOff>${pxToEmu(colOff)}</xdr:colOff><xdr:row>${row - 1}</xdr:row><xdr:rowOff>${pxToEmu(rowOff)}</xdr:rowOff></xdr:${tag}>`;
  return pt("from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff) + pt("to", a.toCol, a.toColOff, a.toRow, a.toRowOff);
}

/** Append a freshly-authored shape to the sheet's drawing; records the drawing path + anchor index. */
function createShape(wb: Workbook, sheet: Sheet, sh: SheetShape, id: number): void {
  const drawingPath = ensureSheetDrawing(wb, sheet);
  const sp = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>${xfrmXml(sh)}${spPrXml(sh)}</xdr:spPr>${txBodyXml(sh)}</xdr:sp>`;
  const anchor = `<xdr:twoCellAnchor editAs="oneCell">${anchorPointsXml(sh)}${sp}<xdr:clientData/></xdr:twoCellAnchor>`;
  const xml = new TextDecoder().decode(wb.files[drawingPath]);
  wb.files[drawingPath] = new TextEncoder().encode(xml.replace(/<\/xdr:wsDr>\s*$/, `${anchor}</xdr:wsDr>`));
  const doc = parseXmlOpt(wb.files[drawingPath]);
  const anchors = doc ? Array.from(doc.documentElement.children).filter((e) => /Anchor$/.test(e.localName)) : [];
  sh.drawingPath = drawingPath;
  sh.anchorIndex = anchors.length - 1;
  sh.created = false;
}

/** Patch an existing shape's anchor + fill / outline / text in its drawing part. */
function patchShape(wb: Workbook, sh: SheetShape): void {
  if (!sh.drawingPath || sh.anchorIndex == null || !wb.files[sh.drawingPath]) return;
  const doc = parseXmlOpt(wb.files[sh.drawingPath]);
  if (!doc) return;
  const anchorEl = Array.from(doc.documentElement.children).filter((e) => /Anchor$/.test(e.localName))[sh.anchorIndex];
  if (!anchorEl) return;
  // anchor from/to (only for twoCellAnchor; a shape we authored always is one)
  const a = sh.anchor;
  const setPt = (tag: string, col: number, colOff: number, row: number, rowOff: number): void => {
    const p = kid(anchorEl, tag);
    if (!p) return;
    const set = (l: string, v: number) => { const e = kid(p, l); if (e) e.textContent = String(v); };
    set("col", col - 1); set("colOff", pxToEmu(colOff)); set("row", row - 1); set("rowOff", pxToEmu(rowOff));
  };
  setPt("from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff);
  setPt("to", a.toCol, a.toColOff, a.toRow, a.toRowOff);
  const sp = kid(anchorEl, "sp") ?? kid(anchorEl, "cxnSp");
  const spPr = sp && kid(sp, "spPr");
  // A move or a resize changes the anchor and NOTHING else. Rewriting the paint here would flatten
  // whatever the file states (a gradient, or a shape that takes its look from the theme's format
  // scheme through <xdr:style>) into the one colour the model carries, so dragging a shape would
  // silently repaint it. Only an actual restyle touches spPr.
  // Only turned or mirrored: touch <a:xfrm> and nothing else, so the paint stays the file's.
  if (spPr && sh.xfrmDirty && !sh.styleDirty && !sh.created) {
    let xfrm = kid(spPr, "xfrm");
    if (!xfrm) {
      xfrm = doc.createElementNS(A, "a:xfrm");
      spPr.insertBefore(xfrm, spPr.firstChild);
    }
    for (const [attr, on] of [["flipH", sh.flipH], ["flipV", sh.flipV]] as const) {
      if (on) xfrm.setAttribute(attr, "1");
      else xfrm.removeAttribute(attr);
    }
    if (sh.rotation) xfrm.setAttribute("rot", String(Math.round(sh.rotation * 60000)));
    else xfrm.removeAttribute("rot");
  }
  if (spPr && (sh.styleDirty || sh.created)) {
    // Replace the fill child (after prstGeom) and the ln, then rebuild txBody, from a parsed fragment.
    const frag = parseXmlOpt(new TextEncoder().encode(`<r xmlns:a="${A}" xmlns:xdr="${XDR}"><a:spPr>${xfrmXml(sh)}${spPrXml(sh)}</a:spPr>${txBodyXml(sh)}</r>`));
    const newSpPr = frag && kid(frag.documentElement, "spPr");
    const newTx = frag && kid(frag.documentElement, "txBody");
    if (newSpPr) {
      for (const c of Array.from(spPr.children)) if (["prstGeom", "noFill", "solidFill", "gradFill", "pattFill", "blipFill", "ln"].includes(c.localName)) spPr.removeChild(c);
      // The turn and the mirror live in <a:xfrm>, which the fragment only carries when the model
      // has one; an existing xfrm is otherwise left exactly as the file wrote it.
      const newXfrm = frag && kid(frag.documentElement, "spPr") && kid(kid(frag.documentElement, "spPr")!, "xfrm");
      let xfrm = kid(spPr, "xfrm");
      if (newXfrm) {
        const keepExt = xfrm ? kid(xfrm, "ext") : undefined;
        const keepOff = xfrm ? kid(xfrm, "off") : undefined;
        const fresh = doc.importNode(newXfrm, true) as Element;
        if (!kid(fresh, "ext") && keepExt) { if (keepOff) fresh.appendChild(keepOff.cloneNode(true)); fresh.appendChild(keepExt.cloneNode(true)); }
        if (xfrm) spPr.replaceChild(fresh, xfrm); else spPr.insertBefore(fresh, spPr.firstChild);
        xfrm = fresh;
      }
      const ref = xfrm ? xfrm.nextSibling : spPr.firstChild; // capture once so inserts keep order
      for (const c of Array.from(newSpPr.children)) spPr.insertBefore(doc.importNode(c, true), ref);
    }
    if (newTx && sp) { const oldTx = kid(sp, "txBody"); if (oldTx) sp.removeChild(oldTx); sp.appendChild(doc.importNode(newTx, true)); }
  }
  wb.files[sh.drawingPath] = serializeXml(doc);
}

/** Remove a shape's anchor from its drawing part, and re-index the remaining shapes on that drawing
    (their anchorIndex shifts down by one for anchors after the removed one). */
export function deleteXlsxShape(wb: Workbook, sheet: Sheet, sh: SheetShape): void {
  if (!sh.drawingPath || sh.anchorIndex == null || !wb.files[sh.drawingPath]) return;
  const doc = parseXmlOpt(wb.files[sh.drawingPath]);
  if (!doc) return;
  const anchors = Array.from(doc.documentElement.children).filter((e) => /Anchor$/.test(e.localName));
  const target = anchors[sh.anchorIndex];
  if (!target) return;
  target.parentNode?.removeChild(target);
  wb.files[sh.drawingPath] = serializeXml(doc);
  for (const other of sheet.shapes ?? []) {
    if (other.drawingPath === sh.drawingPath && other.anchorIndex != null && other.anchorIndex > sh.anchorIndex) other.anchorIndex -= 1;
  }
}

/** Persist all dirty shapes to the workbook's drawing parts. */
export function writeXlsxShapes(wb: Workbook): void {
  let nextId = 1000;
  for (const sheet of wb.sheets as Sheet[]) {
    for (const sh of sheet.shapes ?? []) {
      if (!sh.dirty && !sh.created && !sh.xfrmDirty) continue;
      if (sh.created || sh.drawingPath == null) createShape(wb, sheet, sh, nextId++);
      else patchShape(wb, sh);
      sh.dirty = false;
      sh.xfrmDirty = false;
      sh.styleDirty = false;
    }
  }
}
