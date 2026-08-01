import { ensureCell, parseXmlOpt, serializeXml, type Sheet, type SheetImage, type Workbook } from "./model";
import { pxToEmu, type ChartAnchor } from "./chart-model";

// Adding and removing pictures.
//
// Until now a picture could be moved, resized and replaced but never added or removed, so
// a workbook could only ever hold the pictures it arrived with. That is a real limit for
// anyone building a sheet rather than editing someone else's, and it also made the
// collaboration ids derivable from position, which stops being true the moment two people
// can each add one.
//
// Both formats keep the bytes in a part of their own and the placement in the document, so
// adding one is three steps in each: write the media part, register it, and put an anchored
// object in the sheet.

const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const XLINK = "http://www.w3.org/1999/xlink";
const TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

const MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
  svg: "image/svg+xml",
  tif: "image/tiff",
  tiff: "image/tiff",
};

/** Formats whose pictures this editor can add to and remove from. */
export const imagesInsertable = (wb: Workbook): boolean => wb.kind === "xlsx" || wb.kind === "ods";

const toDataUri = (bytes: Uint8Array, ext: string): string => {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return `data:${MIME[ext] ?? "application/octet-stream"};base64,${btoa(bin)}`;
};

/** A part path nothing is using yet. */
function freePath(wb: Workbook, dir: string, stem: string, ext: string): string {
  for (let n = 1; ; n++) {
    const path = `${dir}/${stem}${n}.${ext}`;
    if (!wb.files[path]) return path;
  }
}

function ensureContentType(wb: Workbook, ext: string): void {
  const raw = wb.files["[Content_Types].xml"];
  const doc = raw ? parseXmlOpt(raw) : undefined;
  if (!doc) return;
  const has = Array.from(doc.getElementsByTagName("Default")).some(
    (d) => (d.getAttribute("Extension") || "").toLowerCase() === ext,
  );
  if (has) return;
  const d = doc.createElementNS(CT_NS, "Default");
  d.setAttribute("Extension", ext);
  d.setAttribute("ContentType", MIME[ext] ?? "application/octet-stream");
  doc.documentElement.insertBefore(d, doc.documentElement.firstChild);
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

/** Add a relationship and return the id it was given. */
function addRel(wb: Workbook, relsPath: string, type: string, target: string): string | null {
  const doc = wb.files[relsPath]
    ? parseXmlOpt(wb.files[relsPath])
    : parseXmlOpt(new TextEncoder().encode(`<?xml version="1.0"?><Relationships xmlns="${PKG_REL}"></Relationships>`));
  if (!doc) return null;
  const used = new Set(Array.from(doc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id") ?? ""));
  let id = "";
  for (let n = 1; ; n++) {
    id = `rId${n}`;
    if (!used.has(id)) break;
  }
  const rel = doc.createElementNS(PKG_REL, "Relationship");
  rel.setAttribute("Id", id);
  rel.setAttribute("Type", type);
  rel.setAttribute("Target", target);
  doc.documentElement.appendChild(rel);
  wb.files[relsPath] = serializeXml(doc);
  return id;
}

const drawingRelsPath = (drawingPath: string): string =>
  drawingPath.replace(/drawings\/(drawing[^/]+\.xml)$/i, "drawings/_rels/$1.rels");

/**
 * The sheet's drawing part, created if it has none.
 *
 * A worksheet with no pictures or charts has no drawing at all, so the first picture added
 * to it has to bring the part, the relationship, the content type and the `<drawing>`
 * element in the sheet with it.
 */
function ensureDrawing(wb: Workbook, sheet: Sheet): string | null {
  const existing = sheet.images?.find((im) => im.drawingPath)?.drawingPath ?? sheet.charts?.find((c) => c.original?.drawingPath)?.original?.drawingPath;
  if (existing) return existing;
  if (!sheet.path || !sheet.doc) return null;

  const path = freePath(wb, "xl/drawings", "drawing", "xml");
  wb.files[path] = new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"></xdr:wsDr>`,
  );
  const ct = wb.files["[Content_Types].xml"];
  const ctDoc = ct ? parseXmlOpt(ct) : undefined;
  if (ctDoc) {
    const ov = ctDoc.createElementNS(CT_NS, "Override");
    ov.setAttribute("PartName", `/${path}`);
    ov.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
    ctDoc.documentElement.appendChild(ov);
    wb.files["[Content_Types].xml"] = serializeXml(ctDoc);
  }

  const sheetRels = sheet.path.replace(/worksheets\/([^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const rId = addRel(wb, sheetRels, `${R_NS}/drawing`, `../drawings/${path.replace(/^.*\//, "")}`);
  if (!rId) return null;

  // The <drawing> element goes last in the worksheet: the schema fixes the order and Excel
  // refuses the file outright if it is anywhere else.
  const el = sheet.doc.createElementNS(sheet.doc.documentElement.namespaceURI, "drawing");
  el.setAttributeNS(R_NS, "r:id", rId);
  sheet.doc.documentElement.appendChild(el);
  sheet.layoutDirty = true;
  return path;
}

function insertXlsxImage(wb: Workbook, sheet: Sheet, bytes: Uint8Array, ext: string, anchor: ChartAnchor): SheetImage | null {
  const drawingPath = ensureDrawing(wb, sheet);
  if (!drawingPath) return null;
  const doc = parseXmlOpt(wb.files[drawingPath]);
  if (!doc) return null;

  const mediaPath = freePath(wb, "xl/media", "image", ext);
  wb.files[mediaPath] = bytes;
  ensureContentType(wb, ext);
  const rId = addRel(wb, drawingRelsPath(drawingPath), `${R_NS}/image`, `../media/${mediaPath.replace(/^.*\//, "")}`);
  if (!rId) return null;

  const point = (tag: string, col: number, colOff: number, row: number, rowOff: number): Element => {
    const p = doc.createElementNS(XDR, `xdr:${tag}`);
    const add = (local: string, v: number): void => {
      const e = doc.createElementNS(XDR, `xdr:${local}`);
      e.textContent = String(v);
      p.appendChild(e);
    };
    add("col", col - 1);
    add("colOff", pxToEmu(colOff));
    add("row", row - 1);
    add("rowOff", pxToEmu(rowOff));
    return p;
  };

  const two = doc.createElementNS(XDR, "xdr:twoCellAnchor");
  two.setAttribute("editAs", "oneCell");
  two.appendChild(point("from", anchor.fromCol, anchor.fromColOff, anchor.fromRow, anchor.fromRowOff));
  two.appendChild(point("to", anchor.toCol, anchor.toColOff, anchor.toRow, anchor.toRowOff));

  const pic = doc.createElementNS(XDR, "xdr:pic");
  const nv = doc.createElementNS(XDR, "xdr:nvPicPr");
  const cNvPr = doc.createElementNS(XDR, "xdr:cNvPr");
  const anchors = Array.from(doc.documentElement.children).filter((e) => /Anchor$/.test(e.localName));
  cNvPr.setAttribute("id", String(anchors.length + 2));
  cNvPr.setAttribute("name", `Picture ${anchors.length + 1}`);
  nv.appendChild(cNvPr);
  nv.appendChild(doc.createElementNS(XDR, "xdr:cNvPicPr"));
  pic.appendChild(nv);

  const blipFill = doc.createElementNS(XDR, "xdr:blipFill");
  const blip = doc.createElementNS(A_NS, "a:blip");
  blip.setAttributeNS(R_NS, "r:embed", rId);
  blipFill.appendChild(blip);
  blipFill.appendChild(doc.createElementNS(A_NS, "a:stretch"));
  pic.appendChild(blipFill);

  const spPr = doc.createElementNS(XDR, "xdr:spPr");
  const geom = doc.createElementNS(A_NS, "a:prstGeom");
  geom.setAttribute("prst", "rect");
  geom.appendChild(doc.createElementNS(A_NS, "a:avLst"));
  spPr.appendChild(geom);
  pic.appendChild(spPr);

  two.appendChild(pic);
  two.appendChild(doc.createElementNS(XDR, "xdr:clientData"));
  doc.documentElement.appendChild(two);
  wb.files[drawingPath] = serializeXml(doc);

  const image: SheetImage = {
    anchor: { ...anchor },
    dataUri: toDataUri(bytes, ext),
    drawingPath,
    anchorIndex: anchors.length,
    mediaPath,
  };
  (sheet.images ??= []).push(image);
  return image;
}

function insertOdsImage(wb: Workbook, sheet: Sheet, bytes: Uint8Array, ext: string, anchor: ChartAnchor): SheetImage | null {
  const doc = wb.contentDoc;
  if (!doc || !sheet.tableEl) return null;
  const cell = ensureCell(sheet, anchor.fromRow, anchor.fromCol);
  if (!cell.el) return null;

  const href = freePath(wb, "Pictures", "image", ext);
  wb.files[href] = bytes;

  const frame = doc.createElementNS(DRAW, "draw:frame");
  frame.setAttributeNS(TABLE, "table:end-cell-address", `${colLetters(anchor.toCol)}${anchor.toRow}`);
  frame.setAttributeNS(SVG, "svg:x", `${anchor.fromColOff}px`);
  frame.setAttributeNS(SVG, "svg:y", `${anchor.fromRowOff}px`);
  frame.setAttributeNS(SVG, "svg:width", `${Math.max(1, anchor.toColOff - anchor.fromColOff)}px`);
  frame.setAttributeNS(SVG, "svg:height", `${Math.max(1, anchor.toRowOff - anchor.fromRowOff)}px`);
  const img = doc.createElementNS(DRAW, "draw:image");
  img.setAttributeNS(XLINK, "xlink:href", href);
  img.setAttributeNS(XLINK, "xlink:type", "simple");
  img.setAttributeNS(XLINK, "xlink:show", "embed");
  img.setAttributeNS(XLINK, "xlink:actuate", "onLoad");
  // A draw:image must carry a text:p, per the ODF schema; readers reject the frame without it.
  img.appendChild(doc.createElementNS(TEXT_NS, "text:p"));
  frame.appendChild(img);
  cell.el.appendChild(frame);

  addOdsManifestEntry(wb, href, ext);
  sheet.odsDirty = true;

  const image: SheetImage = {
    anchor: { ...anchor },
    dataUri: toDataUri(bytes, ext),
    odsFrameEl: frame,
    odsAnchorCol: anchor.fromCol,
    odsAnchorRow: anchor.fromRow,
    mediaPath: href,
  };
  (sheet.images ??= []).push(image);
  return image;
}

/** ODS lists every part in META-INF/manifest.xml; one that is missing is one that is ignored. */
function addOdsManifestEntry(wb: Workbook, href: string, ext: string): void {
  const raw = wb.files["META-INF/manifest.xml"];
  const doc = raw ? parseXmlOpt(raw) : undefined;
  if (!doc) return;
  const ns = doc.documentElement.namespaceURI ?? "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
  const already = Array.from(doc.getElementsByTagName("*")).some(
    (e) => e.localName === "file-entry" && (e.getAttribute("manifest:full-path") ?? "") === href,
  );
  if (already) return;
  const entry = doc.createElementNS(ns, "manifest:file-entry");
  entry.setAttributeNS(ns, "manifest:full-path", href);
  entry.setAttributeNS(ns, "manifest:media-type", MIME[ext] ?? "application/octet-stream");
  doc.documentElement.appendChild(entry);
  wb.files["META-INF/manifest.xml"] = serializeXml(doc);
}

const colLetters = (col: number): string => {
  let s = "";
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

/** Put a picture on a sheet. Returns it, or null when this file type has no drawings. */
export function insertImage(
  wb: Workbook,
  sheet: Sheet,
  bytes: Uint8Array,
  ext: string,
  anchor: ChartAnchor,
): SheetImage | null {
  const clean = (ext || "png").toLowerCase().replace(/[^a-z0-9]/g, "");
  if (wb.kind === "xlsx") return insertXlsxImage(wb, sheet, bytes, clean, anchor);
  if (wb.kind === "ods") return insertOdsImage(wb, sheet, bytes, clean, anchor);
  return null;
}

/**
 * Take a picture off a sheet.
 *
 * The media part is left where it is. It may be shared with another picture, and an unused
 * part is a few kilobytes that every reader ignores, where deleting one still referenced
 * would produce a file that does not open.
 */
export function deleteImage(wb: Workbook, sheet: Sheet, im: SheetImage): boolean {
  const at = sheet.images?.indexOf(im) ?? -1;
  if (at < 0) return false;

  if (im.odsFrameEl) {
    im.odsFrameEl.parentNode?.removeChild(im.odsFrameEl);
    sheet.odsDirty = true;
  } else if (im.drawingPath && im.anchorIndex != null && wb.files[im.drawingPath]) {
    const doc = parseXmlOpt(wb.files[im.drawingPath]);
    if (doc) {
      const anchors = Array.from(doc.documentElement.children).filter((e) => /Anchor$/.test(e.localName));
      const el = anchors[im.anchorIndex];
      if (el) {
        el.parentNode?.removeChild(el);
        wb.files[im.drawingPath] = serializeXml(doc);
        // Everything after it in the same drawing has shifted up by one.
        for (const other of sheet.images ?? []) {
          if (other !== im && other.drawingPath === im.drawingPath && (other.anchorIndex ?? 0) > im.anchorIndex) {
            other.anchorIndex = (other.anchorIndex ?? 0) - 1;
          }
        }
      }
    }
  }

  sheet.images!.splice(at, 1);
  return true;
}
