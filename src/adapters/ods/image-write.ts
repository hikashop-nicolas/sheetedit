import { parseXmlOpt, serializeXml, type SheetImage, type Workbook } from "../../core/model";
import { ODS } from "./shared";

// Persist a moved/resized picture back into its draw:frame. The frame stays anchored to its original
// cell; we rewrite svg:x / svg:y (the offset from that cell, computed against the grid geometry in
// the overlay) and svg:width / svg:height, and drop any two-cell end address so the size is explicit.
// The frame lives in the persistent contentDoc, so patching it in place flows through writeOds's
// serialization. Runs BEFORE writeOds so a touched sheet re-emits the patched frame.

const cm = (px: number): string => `${((px / 96) * 2.54).toFixed(3)}cm`; // 96px/in, 2.54cm/in

function patchFrame(im: SheetImage): void {
  const frame = im.odsFrameEl;
  const f = im.odsFrame;
  if (!frame || !f) return;
  frame.setAttributeNS(ODS.svg, "svg:x", cm(f.x));
  frame.setAttributeNS(ODS.svg, "svg:y", cm(f.y));
  frame.setAttributeNS(ODS.svg, "svg:width", cm(f.w));
  frame.setAttributeNS(ODS.svg, "svg:height", cm(f.h));
  // An explicit size supersedes a two-cell end anchor; remove it so the two do not conflict.
  for (const a of ["end-cell-address", "end-x", "end-y"]) frame.removeAttributeNS(ODS.table, a);
}

const IMG_MIME: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", tif: "image/tiff", tiff: "image/tiff" };

/** Register a part in META-INF/manifest.xml (adds a file-entry if absent). */
function addManifestEntry(wb: Workbook, path: string, mime: string): void {
  const m = wb.files["META-INF/manifest.xml"];
  if (!m) return;
  const doc = parseXmlOpt(m);
  if (!doc) return;
  const has = Array.from(doc.getElementsByTagName("*")).some((e) => e.localName === "file-entry" && e.getAttribute("manifest:full-path") === path);
  if (!has) {
    const MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
    const fe = doc.createElementNS(MANIFEST, "manifest:file-entry");
    fe.setAttribute("manifest:full-path", path);
    fe.setAttribute("manifest:media-type", mime);
    doc.documentElement.appendChild(fe);
    wb.files["META-INF/manifest.xml"] = serializeXml(doc);
  }
}

/** Swap an ods image's bytes. Same extension overwrites the Pictures/ part in place; a different
    extension writes a new part, repoints the frame's draw:image href and registers the part. */
function replaceOdsImage(wb: Workbook, im: SheetImage): void {
  if (!im.replaceBytes || !im.mediaPath) return;
  const oldExt = (im.mediaPath.split(".").pop() || "png").toLowerCase();
  const ext = (im.replaceExt || oldExt).toLowerCase();
  if (ext === oldExt) { wb.files[im.mediaPath] = im.replaceBytes; }
  else {
    const newPath = im.mediaPath.replace(/[^/]+$/, `${im.mediaPath.replace(/^.*\//, "").replace(/\.[^.]+$/, "")}.${ext}`);
    wb.files[newPath] = im.replaceBytes;
    addManifestEntry(wb, newPath, IMG_MIME[ext] ?? "application/octet-stream");
    const img = im.odsFrameEl && Array.from(im.odsFrameEl.children).find((c) => c.localName === "image");
    if (img) img.setAttributeNS(ODS.xlink, "xlink:href", newPath);
    im.mediaPath = newPath; // old part left as a harmless orphan (may be shared)
  }
  im.replaceBytes = undefined;
  im.replaceExt = undefined;
}

/** Persist all dirty ods images (frame moves/resizes + media replacements). Call before writeOds
    serializes content.xml. */
export function writeOdsImages(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    for (const im of sheet.images ?? []) {
      if (!im.dirty) continue;
      if (im.replaceBytes) replaceOdsImage(wb, im);
      if (im.odsFrameEl && im.odsFrame) patchFrame(im);
    }
  }
}
