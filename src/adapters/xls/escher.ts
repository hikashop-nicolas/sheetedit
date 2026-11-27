// Office Drawing (Escher, [MS-ODRAW]) as .xls embeds it: the workbook's BLIP store (pictures) in
// MSODRAWINGGROUP and each sheet's shape tree in MSODRAWING, tied to OBJ / TXO records in order.
import { unzlibSync } from "fflate";

export interface EscherRec { ver: number; inst: number; type: number; start: number; len: number; data: Uint8Array }

/** Walk the records at one level of an Escher buffer. */
export function escherChildren(buf: Uint8Array, from = 0, to = buf.length): EscherRec[] {
  const out: EscherRec[] = [];
  let p = from;
  while (p + 8 <= to) {
    const vi = buf[p]! | (buf[p + 1]! << 8);
    const type = buf[p + 2]! | (buf[p + 3]! << 8);
    const len = (buf[p + 4]! | (buf[p + 5]! << 8) | (buf[p + 6]! << 16) | (buf[p + 7]! << 24)) >>> 0;
    if (type < 0xf000 || type > 0xf2ff) break;
    const end = Math.min(to, p + 8 + len);
    out.push({ ver: vi & 0xf, inst: vi >> 4, type, start: p + 8, len: end - p - 8, data: buf.subarray(p + 8, end) });
    p = end;
  }
  return out;
}

export interface Blip { ext: string; bytes: Uint8Array }

const BLIP_EXT: Record<number, string> = { 0xf01a: "emf", 0xf01b: "wmf", 0xf01c: "pict", 0xf01d: "jpeg", 0xf01e: "png", 0xf01f: "bmp", 0xf029: "tiff", 0xf02a: "jpeg" };
function decodeBlip(r: EscherRec): Blip | undefined {
  const ext = BLIP_EXT[r.type];
  if (!ext) return undefined;
  const d = r.data;
  let off = 16 + (r.inst & 1 ? 16 : 0); // an odd instance carries a second 16-byte UID
  if (ext === "emf" || ext === "wmf" || ext === "pict") {
    const dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
    if (off + 34 > d.length) return undefined;
    const cb = dv.getUint32(off, true);
    const bounds = [dv.getInt32(off + 4, true), dv.getInt32(off + 8, true), dv.getInt32(off + 12, true), dv.getInt32(off + 16, true)];
    const compression = d[off + 32];
    off += 34;
    let bytes = d.subarray(off);
    if (compression === 0) {
      try { bytes = unzlibSync(bytes); } catch { return undefined; }
    }
    if (cb && bytes.length > cb) bytes = bytes.subarray(0, cb);
    if (ext === "wmf" && !(bytes[0] === 0xd7 && bytes[1] === 0xcd && bytes[2] === 0xc6 && bytes[3] === 0x9a)) bytes = placeable(bytes, bounds);
    return { ext, bytes };
  }
  off += 1; // tag
  let bytes = d.subarray(off);
  if (ext === "bmp") bytes = bmpFromDib(bytes);
  return { ext, bytes };
}

/** A WMF placeable header in front of a raw metafile, so image tools know its extent. */
function placeable(wmf: Uint8Array, b: number[]): Uint8Array {
  const out = new Uint8Array(22 + wmf.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, 0x9ac6cdd7, true);
  dv.setInt16(6, b[0]!, true); dv.setInt16(8, b[1]!, true); dv.setInt16(10, b[2]!, true); dv.setInt16(12, b[3]!, true);
  dv.setUint16(14, 1440, true);
  let sum = 0;
  for (let i = 0; i < 20; i += 2) sum ^= dv.getUint16(i, true);
  dv.setUint16(20, sum, true);
  out.set(wmf, 22);
  return out;
}

/** A BMP file from a packed DIB (BITMAPINFOHEADER + palette + bits). */
function bmpFromDib(dib: Uint8Array): Uint8Array {
  const dv = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const hdr = dib.length >= 4 ? dv.getUint32(0, true) : 40;
  const bpp = dib.length >= 16 ? dv.getUint16(14, true) : 24;
  const used = dib.length >= 36 ? dv.getUint32(32, true) : 0;
  const colors = bpp <= 8 ? used || 1 << bpp : 0;
  const out = new Uint8Array(14 + dib.length);
  const ov = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d;
  ov.setUint32(2, out.length, true);
  ov.setUint32(10, 14 + hdr + colors * 4, true);
  out.set(dib, 14);
  return out;
}

/** The pictures in the workbook's BLIP store, by 1-based index (index 0 unused). */
export function readBlipStore(dgg: Uint8Array): (Blip | undefined)[] {
  const out: (Blip | undefined)[] = [undefined];
  for (const top of escherChildren(dgg)) {
    if (top.type !== 0xf000) continue;
    for (const r of escherChildren(dgg, top.start, top.start + top.len)) {
      if (r.type !== 0xf001) continue;
      for (const bse of escherChildren(dgg, r.start, r.start + r.len)) {
        if (bse.type !== 0xf007) { out.push(undefined); continue; }
        const cbName = bse.data[33] ?? 0;
        const inner = escherChildren(dgg, bse.start + 36 + cbName, bse.start + bse.len);
        out.push(inner[0] ? decodeBlip(inner[0]) : undefined);
      }
    }
  }
  return out;
}

export interface Anchor { col1: number; dx1: number; row1: number; dy1: number; col2: number; dx2: number; row2: number; dy2: number }

export interface DrawShape {
  spid: number;
  shapeType: number;
  anchor?: Anchor;
  /** 1-based BLIP index for a picture. */
  pib?: number;
  /** Index of the OBJ (ClientData) and TXO (ClientTextbox) records this shape owns. */
  obj?: number;
  txo?: number;
  name?: string;
}

/** The top-level shapes of a sheet drawing (group children are skipped). */
export function readSheetShapes(dg: Uint8Array): DrawShape[] {
  const out: DrawShape[] = [];
  let objN = 0, txoN = 0;
  const visit = (from: number, to: number, depth: number): void => {
    for (const r of escherChildren(dg, from, to)) {
      if (r.type === 0xf002 || r.type === 0xf003) { visit(r.start, r.start + r.len, r.type === 0xf003 ? depth + 1 : depth); continue; }
      if (r.type !== 0xf004) continue;
      const sh: DrawShape = { spid: 0, shapeType: 0 };
      let group = false;
      for (const k of escherChildren(dg, r.start, r.start + r.len)) {
        const dv = new DataView(k.data.buffer, k.data.byteOffset, k.data.byteLength);
        if (k.type === 0xf00a && k.len >= 8) {
          sh.spid = dv.getUint32(0, true);
          sh.shapeType = k.inst;
          if (dv.getUint32(4, true) & 0x5) group = true; // fGroup / fPatriarch
        } else if (k.type === 0xf00b) {
          const n = k.inst;
          let complex = 6 * n;
          for (let i = 0; i < n && i * 6 + 6 <= k.len; i++) {
            const opid = dv.getUint16(i * 6, true), op = dv.getUint32(i * 6 + 2, true);
            const pid = opid & 0x3fff;
            if (pid === 0x104 || (pid === 0x186 && sh.pib == null)) sh.pib = op;
            if (opid & 0x8000) {
              if (pid === 0x380 && complex + op <= k.len) sh.name = utf16z(k.data.subarray(complex, complex + op));
              complex += op;
            }
          }
        } else if (k.type === 0xf010 && k.len >= 18) {
          sh.anchor = { col1: dv.getUint16(2, true), dx1: dv.getUint16(4, true), row1: dv.getUint16(6, true), dy1: dv.getUint16(8, true), col2: dv.getUint16(10, true), dx2: dv.getUint16(12, true), row2: dv.getUint16(14, true), dy2: dv.getUint16(16, true) };
        } else if (k.type === 0xf011) sh.obj = objN++;
        else if (k.type === 0xf00d) sh.txo = txoN++;
      }
      if (!group && depth <= 1) out.push(sh);
    }
  };
  visit(0, dg.length, 0);
  return out;
}

/** One shape of a sheet drawing with everything its container says, groups holding their members. */
export interface DrawNode {
  spid: number;
  spt: number;
  /** FSP flags: 0x1 group, 0x2 child, 0x40 flipH, 0x80 flipV. */
  flags: number;
  props: Map<number, { value: number; complex?: Uint8Array }>;
  anchor?: Anchor;
  /** A group member's box in its group's coordinates. */
  child?: { l: number; t: number; r: number; b: number };
  /** A group's own coordinate space. */
  space?: { l: number; t: number; r: number; b: number };
  obj?: number;
  txo?: number;
  kids?: DrawNode[];
  /** The shape container's own bytes, header included, for an exact write-back. */
  raw: Uint8Array;
}

/** The sheet drawing as a tree: top-level shapes, with groups holding their members. */
export function readDrawingTree(dg: Uint8Array): DrawNode[] {
  let objN = 0, txoN = 0;
  const shape = (r: EscherRec): DrawNode => {
    const n: DrawNode = { spid: 0, spt: 0, flags: 0, props: new Map(), raw: dg.subarray(r.start - 8, r.start + r.len) };
    for (const k of escherChildren(dg, r.start, r.start + r.len)) {
      const dv = new DataView(k.data.buffer, k.data.byteOffset, k.data.byteLength);
      if (k.type === 0xf00a && k.len >= 8) { n.spid = dv.getUint32(0, true); n.spt = k.inst; n.flags = dv.getUint32(4, true); }
      else if ((k.type === 0xf00b || k.type === 0xf122) && k.len >= 6) {
        let complex = 6 * k.inst;
        for (let i = 0; i < k.inst && i * 6 + 6 <= k.len; i++) {
          const opid = dv.getUint16(i * 6, true), value = dv.getUint32(i * 6 + 2, true);
          const p: { value: number; complex?: Uint8Array } = { value };
          if (opid & 0x8000) { p.complex = k.data.subarray(complex, Math.min(k.len, complex + value)); complex += value; }
          if (!n.props.has(opid & 0x3fff)) n.props.set(opid & 0x3fff, p);
        }
      } else if (k.type === 0xf010 && k.len >= 18) {
        n.anchor = { col1: dv.getUint16(2, true), dx1: dv.getUint16(4, true), row1: dv.getUint16(6, true), dy1: dv.getUint16(8, true), col2: dv.getUint16(10, true), dx2: dv.getUint16(12, true), row2: dv.getUint16(14, true), dy2: dv.getUint16(16, true) };
      } else if (k.type === 0xf00f && k.len >= 16) n.child = { l: dv.getInt32(0, true), t: dv.getInt32(4, true), r: dv.getInt32(8, true), b: dv.getInt32(12, true) };
      else if (k.type === 0xf009 && k.len >= 16) n.space = { l: dv.getInt32(0, true), t: dv.getInt32(4, true), r: dv.getInt32(8, true), b: dv.getInt32(12, true) };
      else if (k.type === 0xf011) n.obj = objN++;
      else if (k.type === 0xf00d) n.txo = txoN++;
    }
    return n;
  };
  const group = (from: number, to: number, top: boolean): DrawNode[] => {
    const out: DrawNode[] = [];
    let head: DrawNode | undefined;
    for (const r of escherChildren(dg, from, to)) {
      if (r.type === 0xf004) {
        const n = shape(r);
        if (!head) { head = n; if (top || n.flags & 0x4) continue; n.kids = []; out.push(n); continue; }
        (out[0]?.kids && !top ? out[0].kids : out).push(n);
      } else if (r.type === 0xf003) {
        const g = group(r.start, r.start + r.len, false);
        (out[0]?.kids && !top ? out[0].kids : out).push(...g);
      }
    }
    return out;
  };
  for (const top of escherChildren(dg)) {
    if (top.type !== 0xf002) continue;
    for (const r of escherChildren(dg, top.start, top.start + top.len)) if (r.type === 0xf003) return group(r.start, r.start + r.len, true);
  }
  return [];
}

const utf16z = (b: Uint8Array): string => {
  let s = "";
  for (let i = 0; i + 1 < b.length; i += 2) {
    const c = b[i]! | (b[i + 1]! << 8);
    if (!c) break;
    s += String.fromCharCode(c);
  }
  return s;
};
