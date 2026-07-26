// Windows metafiles (WMF / EMF), which a browser cannot decode on its own.
//
// They are not images in the sense a browser understands: a metafile is a recorded list of GDI
// drawing calls, so showing one means replaying those calls onto a canvas. Office writes them
// constantly - pasted charts, clipart, logos, an ActiveX control's Picture - and until now every
// one of them came through as a `data:image/emf` URI that rendered as nothing at all.
//
// The replay is emf-converter (Apache-2.0, no dependencies of its own), lazy-loaded so a workbook
// with no metafile in it never pays for the code. A conversion that fails leaves the image absent
// rather than broken, which is the same rule the rest of the reader follows.

const METAFILE_MIMES = new Set(["image/emf", "image/x-emf", "image/wmf", "image/x-wmf", "application/x-msmetafile"]);

/** Whether a MIME type names a metafile, which needs replaying rather than decoding. */
export const isMetafileMime = (mime: string): boolean => METAFILE_MIMES.has(mime.toLowerCase());

/** The metafile a `data:` URI holds, or null when it is an ordinary image. */
export function metafileFromDataUri(uri: string): { kind: "emf" | "wmf"; bytes: Uint8Array } | null {
  const m = /^data:([^;,]+);base64,(.*)$/s.exec(uri);
  if (!m || !isMetafileMime(m[1]!)) return null;
  try {
    const bin = atob(m[2]!);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return { kind: metafileKind(bytes) ?? (/wmf/i.test(m[1]!) ? "wmf" : "emf"), bytes };
  } catch {
    return null;
  }
}

/**
 * Which metafile a run of bytes is, by its own header rather than by the name it arrived under.
 * A placeable WMF starts with the 0x9AC6CDD7 key; a bare WMF with its METAHEADER's type and
 * header size; an EMF with record type 1 and the " EMF" signature 40 bytes in.
 */
export function metafileKind(b: Uint8Array): "emf" | "wmf" | null {
  if (b.length < 8) return null;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (dv.getUint32(0, true) === 0x9ac6cdd7) return "wmf";
  if (b.length >= 44 && dv.getUint32(0, true) === 1 && dv.getUint32(40, true) === 0x464d4520) return "emf";
  const type = dv.getUint16(0, true);
  const headerWords = dv.getUint16(2, true);
  if ((type === 1 || type === 2) && headerWords === 9) return "wmf";
  return null;
}

/**
 * The size a metafile asks to be drawn at, in CSS pixels.
 *
 * This has to be worked out here rather than left to the converter: a placeable WMF states its
 * frame in the header's own units, and a converter that misses it renders into a square canvas
 * with the drawing squashed into a corner - and, at the default cap, an 8192x8192 PNG of mostly
 * white for a picture two inches across.
 */
export function metafileSize(b: Uint8Array): { width: number; height: number } | undefined {
  if (b.length < 24) return undefined;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let inches: [number, number] | undefined;
  if (dv.getUint32(0, true) === 0x9ac6cdd7) {
    // Placeable WMF: a bounding box in logical units, and how many of those go to the inch.
    const perInch = dv.getUint16(14, true) || 1440;
    const w = dv.getInt16(10, true) - dv.getInt16(6, true);
    const h = dv.getInt16(12, true) - dv.getInt16(8, true);
    inches = [Math.abs(w) / perInch, Math.abs(h) / perInch];
  } else if (b.length >= 44 && dv.getUint32(0, true) === 1 && dv.getUint32(40, true) === 0x464d4520) {
    // EMF: rclFrame is the picture's real size, in hundredths of a millimetre.
    const w = dv.getInt32(32, true) - dv.getInt32(24, true);
    const h = dv.getInt32(36, true) - dv.getInt32(28, true);
    if (w > 0 && h > 0) inches = [w / 100 / 25.4, h / 100 / 25.4];
  }
  if (!inches || !inches[0] || !inches[1]) return undefined;
  // Bound the long edge: a metafile may declare a frame far larger than anything worth rasterising.
  const [wi, hi] = inches;
  const scale = Math.min(1, MAX_EDGE / (96 * Math.max(wi, hi)));
  return { width: Math.max(1, Math.round(wi * 96 * scale)), height: Math.max(1, Math.round(hi * 96 * scale)) };
}

/** The longest edge worth rasterising, in CSS pixels, before the dpi scale is applied. */
const MAX_EDGE = 1600;

/** Conversions are keyed by the bytes themselves, since the same picture is redrawn on every render. */
const cache = new Map<string, string | null>();
const keyOf = (b: Uint8Array): string => `${b.length}:${b[0]}:${b[b.length >> 1]}:${b[b.length - 1]}`;

/**
 * Replay a metafile onto a canvas and hand back a PNG data URI, or undefined when it cannot be
 * replayed here. Needs a Canvas, so this only does anything in a browser; in a test or a worker
 * without one it returns undefined and the caller carries on without the picture.
 */
export async function metafileToPng(bytes: Uint8Array, kind: "emf" | "wmf"): Promise<string | undefined> {
  const key = `${kind}:${keyOf(bytes)}`;
  const hit = cache.get(key);
  if (hit !== undefined) return hit ?? undefined;
  let out: string | null = null;
  try {
    const mod = await import("emf-converter");
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const size = metafileSize(bytes);
    const opts = size ? { maxWidth: size.width, maxHeight: size.height, dpiScale: 2 } : { maxCanvasDimension: MAX_EDGE };
    out = (await (kind === "wmf" ? mod.convertWmfToDataUrl(buffer, opts) : mod.convertEmfToDataUrl(buffer, opts))) ?? null;
  } catch {
    out = null; // no canvas, or a record set the converter will not replay: no picture, no crash
  }
  cache.set(key, out);
  return out ?? undefined;
}

/**
 * Turn a metafile `data:` URI into one a browser can show. Anything else comes back unchanged, so
 * a caller can pipe every image through this without asking what it is first.
 */
export async function renderableDataUri(uri: string): Promise<string | undefined> {
  const meta = metafileFromDataUri(uri);
  if (!meta) return uri;
  return metafileToPng(meta.bytes, meta.kind);
}
