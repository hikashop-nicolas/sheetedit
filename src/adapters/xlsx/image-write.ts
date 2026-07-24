import { parseXmlOpt, serializeXml, type Sheet, type SheetImage, type Workbook } from "../../core/model";
import { pxToEmu } from "../../core/chart-model";

// Persist a moved/resized picture back into its worksheet drawing part. Only images flagged dirty
// by the overlay are touched; every other drawing (and untouched image) stays verbatim. The anchor
// element is patched in place for a twoCellAnchor, or converted to one (from + to) for a
// oneCellAnchor / absoluteAnchor so the new position and size round-trip.

const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const kid = (parent: Element, local: string): Element | undefined => Array.from(parent.children).find((c) => c.localName === local);

/** Build an <xdr:from>/<xdr:to> point element for the given cell + pixel offset. */
function pointEl(doc: Document, tag: string, col: number, colOff: number, row: number, rowOff: number): Element {
  const p = doc.createElementNS(XDR, `xdr:${tag}`);
  const add = (local: string, v: number) => { const e = doc.createElementNS(XDR, `xdr:${local}`); e.textContent = String(v); p.appendChild(e); };
  add("col", col - 1); add("colOff", pxToEmu(colOff)); add("row", row - 1); add("rowOff", pxToEmu(rowOff));
  return p;
}

function patchAnchorEl(doc: Document, anchorEl: Element, im: SheetImage): void {
  const a = im.anchor;
  const from = kid(anchorEl, "from");
  const to = kid(anchorEl, "to");
  if (from && to) {
    // twoCellAnchor: rewrite the from/to children's col/colOff/row/rowOff.
    const setPt = (p: Element, col: number, colOff: number, row: number, rowOff: number): void => {
      const set = (local: string, v: number): void => { const e = kid(p, local); if (e) e.textContent = String(v); };
      set("col", col - 1); set("colOff", pxToEmu(colOff)); set("row", row - 1); set("rowOff", pxToEmu(rowOff));
    };
    setPt(from, a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff);
    setPt(to, a.toCol, a.toColOff, a.toRow, a.toRowOff);
    return;
  }
  // oneCellAnchor / absoluteAnchor: rebuild as a twoCellAnchor, keeping the object + clientData.
  const two = doc.createElementNS(XDR, "xdr:twoCellAnchor");
  const editAs = anchorEl.getAttribute("editAs"); if (editAs) two.setAttribute("editAs", editAs);
  two.appendChild(pointEl(doc, "from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff));
  two.appendChild(pointEl(doc, "to", a.toCol, a.toColOff, a.toRow, a.toRowOff));
  // Carry over the drawing object (pic/sp/graphicFrame) and clientData; drop from/pos/ext.
  for (const ch of Array.from(anchorEl.children)) {
    if (ch.localName === "from" || ch.localName === "to" || ch.localName === "pos" || ch.localName === "ext") continue;
    two.appendChild(ch);
  }
  anchorEl.parentNode?.replaceChild(two, anchorEl);
}

function rewriteImage(wb: Workbook, im: SheetImage): void {
  if (!im.drawingPath || im.anchorIndex == null || !wb.files[im.drawingPath]) return;
  const doc = parseXmlOpt(wb.files[im.drawingPath]);
  if (!doc) return;
  const anchors = Array.from(doc.documentElement.children).filter((e) => /Anchor$/.test(e.localName));
  const anchorEl = anchors[im.anchorIndex];
  if (!anchorEl) return;
  patchAnchorEl(doc, anchorEl, im);
  wb.files[im.drawingPath] = serializeXml(doc);
}

/** Ensure [Content_Types].xml has a Default for the given image extension. */
function ensureContentType(wb: Workbook, ext: string): void {
  const ct = wb.files["[Content_Types].xml"];
  if (!ct) return;
  const doc = parseXmlOpt(ct);
  if (!doc) return;
  const has = Array.from(doc.getElementsByTagName("Default")).some((d) => (d.getAttribute("Extension") || "").toLowerCase() === ext);
  if (has) return;
  const mime: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", tif: "image/tiff", tiff: "image/tiff" };
  const d = doc.createElementNS(doc.documentElement.namespaceURI, "Default");
  d.setAttribute("Extension", ext); d.setAttribute("ContentType", mime[ext] ?? "application/octet-stream");
  doc.documentElement.insertBefore(d, doc.documentElement.firstChild);
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

/** Point the drawing's relationship for this image at a new media target (used when a replacement
    has a different extension, so the part path changes). */
function retargetDrawingRel(wb: Workbook, im: SheetImage, newFile: string): void {
  if (!im.drawingPath) return;
  const relsPath = im.drawingPath.replace(/drawings\/(drawing[^/]+\.xml)$/i, "drawings/_rels/$1.rels");
  const relsDoc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (!relsDoc) return;
  for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
    const target = r.getAttribute("Target") || "";
    if (/image/i.test(r.getAttribute("Type") || "") && target.replace(/^.*\//, "") === im.mediaPath?.replace(/^.*\//, "")) {
      r.setAttribute("Target", target.replace(/[^/]+$/, newFile));
    }
  }
  wb.files[relsPath] = serializeXml(relsDoc);
}

/** Swap an image's media bytes. Same extension overwrites the part in place; a different extension
    writes a new part, retargets the drawing relationship and registers the content type. */
function replaceImageBytes(wb: Workbook, im: SheetImage): void {
  if (!im.replaceBytes || !im.mediaPath) return;
  const oldExt = (im.mediaPath.split(".").pop() || "png").toLowerCase();
  const ext = (im.replaceExt || oldExt).toLowerCase();
  if (ext === oldExt) { wb.files[im.mediaPath] = im.replaceBytes; }
  else {
    const newPath = im.mediaPath.replace(/[^/]+$/, `${im.mediaPath.replace(/^.*\//, "").replace(/\.[^.]+$/, "")}.${ext}`);
    wb.files[newPath] = im.replaceBytes;
    retargetDrawingRel(wb, im, newPath.replace(/^.*\//, ""));
    ensureContentType(wb, ext);
    im.mediaPath = newPath; // leave the old part as a harmless orphan (may be shared)
  }
  im.replaceBytes = undefined;
  im.replaceExt = undefined;
}

/** Persist all dirty images: media replacements and moved/resized anchors. */
export function writeXlsxImages(wb: Workbook): void {
  for (const sheet of wb.sheets as Sheet[]) {
    for (const im of sheet.images ?? []) {
      if (!im.dirty) continue;
      if (im.replaceBytes) replaceImageBytes(wb, im);
      if (im.odsFrame == null) rewriteImage(wb, im); // anchor rewrite (xlsx images have no odsFrame)
      im.dirty = false;
    }
  }
}
