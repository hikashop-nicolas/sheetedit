import { parseXmlOpt, type Sheet } from "../../core/model";
import { anchorOf, relMap, resolvePart } from "./chart-read";

// Read the pictures anchored on a worksheet: sheet rels -> drawingN.xml (the anchors) -> each
// <xdr:pic> whose blip embeds an image in xl/media. Produces a data-URI + cell anchor per image,
// rendered on an overlay. The drawing + media parts are preserved verbatim on save.

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const descend = (root: Element, local: string): Element[] => Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);

const MIME: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", tif: "image/tiff", tiff: "image/tiff", emf: "image/emf", wmf: "image/wmf" };

function toDataUri(bytes: Uint8Array, path: string): string {
  const ext = (path.split(".").pop() || "").toLowerCase();
  const mime = MIME[ext] ?? "application/octet-stream";
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return `data:${mime};base64,${btoa(bin)}`;
}

/** Populate sheet.images from the worksheet's drawing + media parts. */
export function readImages(sheet: Sheet, files: Record<string, Uint8Array>, path: string): void {
  const relsPath = path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const drawings = relMap(files, relsPath).byType.filter((r) => /drawing/i.test(r.type) && /drawings\//i.test(r.target)).map((r) => resolvePart("xl/worksheets", r.target));
  const out: import("../../core/model").SheetImage[] = [];
  for (const drawPath of drawings) {
    const drawDoc = files[drawPath] ? parseXmlOpt(files[drawPath]) : undefined;
    if (!drawDoc) continue;
    const drawRels = relMap(files, drawPath.replace(/drawings\/(drawing[^/]+\.xml)$/i, "drawings/_rels/$1.rels")).byId;
    const drawBase = drawPath.replace(/\/[^/]+$/, "");
    // Index against all anchor elements so a dirty image can find its exact anchor on write.
    const anchorEls = Array.from(drawDoc.documentElement.children).filter((e) => /Anchor$/.test(e.localName));
    anchorEls.forEach((anchorEl, anchorIndex) => {
      const pic = descend(anchorEl, "pic")[0];
      if (!pic) return;
      const blip = descend(pic, "blip")[0];
      const rid = blip?.getAttributeNS(R_NS, "embed") ?? blip?.getAttribute("r:embed");
      if (!rid || !drawRels.has(rid)) return;
      const mediaPath = resolvePart(drawBase, drawRels.get(rid)!);
      const bytes = files[mediaPath];
      const anchor = anchorOf(anchorEl);
      if (!bytes || !anchor) return;
      out.push({ anchor, dataUri: toDataUri(bytes, mediaPath), drawingPath: drawPath, anchorIndex, mediaPath });
    });
  }
  if (out.length) sheet.images = out;
}
