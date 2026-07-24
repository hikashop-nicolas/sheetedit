import type { Sheet, Workbook } from "../../core/model";
import type { ChartAnchor } from "../../core/chart-model";
import { anchorOf, attrByLocal as A, descend } from "./chart-read";

// Read images embedded in an ODS: a <draw:frame> (anchored to a cell) containing a <draw:image>
// whose xlink:href points to a part in Pictures/ (or a data: URI). Produces a data-URI + cell
// anchor per image, rendered on the same overlay layer as xlsx images. The frame + picture parts
// are preserved verbatim on save (untouched tables/pictures are re-emitted as-is).

const MIME: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", tif: "image/tiff", tiff: "image/tiff" };

function toDataUri(bytes: Uint8Array, path: string): string {
  const ext = (path.split(".").pop() || "").toLowerCase();
  const mime = MIME[ext] ?? "application/octet-stream";
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return `data:${mime};base64,${btoa(bin)}`;
}

/** Populate each sheet's images from the ODS content.xml draw:frame picture objects. */
export function readOdsImages(wb: Workbook, files: Record<string, Uint8Array>): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  for (const frame of descend(doc.documentElement, "frame")) {
    const img = Array.from(frame.children).find((c) => c.localName === "image");
    if (!img) continue;
    // Some producers inline the bytes as <office:binary-data> (base64) instead of an href part.
    const href = (A(img, "href") || "").replace(/^\.\//, "");
    let dataUri: string | undefined;
    if (href.startsWith("data:")) dataUri = href;
    else if (href && files[href]) dataUri = toDataUri(files[href], href);
    else {
      const bin = Array.from(img.children).find((c) => c.localName === "binary-data")?.textContent?.trim();
      if (bin) dataUri = `data:image/png;base64,${bin.replace(/\s+/g, "")}`;
    }
    if (!dataUri) continue;
    // The frame's sheet: nearest ancestor table:table.
    let t: Element | null = frame.parentElement;
    while (t && t.localName !== "table") t = t.parentElement;
    const sheet: Sheet | undefined = t ? wb.sheets.find((s) => s.tableEl === t) : undefined;
    if (!sheet || !t) continue;
    const anchor: ChartAnchor = anchorOf(frame, t);
    (sheet.images ??= []).push({ anchor, dataUri, odsFrameEl: frame, odsAnchorCol: anchor.fromCol, odsAnchorRow: anchor.fromRow, mediaPath: href && !href.startsWith("data:") && files[href] ? href : undefined });
  }
}
