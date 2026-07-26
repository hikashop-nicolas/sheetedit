import { isCfb, readCfb } from "./cfb";

// ---------------------------------------------------------------------------
// VBA project reading (Stage 0 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// A workbook's macros live in vbaProject.bin, a compound file holding a /VBA storage: a `dir`
// stream of metadata and one stream per module, both compressed with the MS-OVBA algorithm.
//
// This extracts the source so it can be READ. Nothing here executes anything. Built from the
// published [MS-OVBA] specification.

/** A module's source and what kind of module it is. */
export interface VbaModule {
  name: string;
  /** The stream under /VBA that held it. */
  stream: string;
  /** MODULETYPE says procedural or class, and no more: a sheet's or the workbook's own module is a
      class module bound to a document, which the dir stream gives no way to tell apart. */
  kind: "standard" | "class";
  source: string;
}

export interface VbaProject {
  modules: VbaModule[];
  /** The code page the source was decoded with. */
  codePage: number;
  /** Whether the project is locked for viewing; the source is still readable, the flag is advisory. */
  locked: boolean;
}

/**
 * MS-OVBA decompression. The container is a 0x01 signature byte followed by chunks; each chunk has
 * a 16-bit header giving its size (bits 0-11, plus 3), a 0b011 signature (bits 12-14) and a
 * compressed flag (bit 15). A compressed chunk is a flag byte then eight tokens: a literal byte, or
 * a two-byte copy whose offset/length bit split widens as the chunk's output grows.
 */
export function decompressOvba(data: Uint8Array): Uint8Array {
  if (data.length === 0 || data[0] !== 0x01) return new Uint8Array(0);
  const out: number[] = [];
  let pos = 1;
  while (pos + 1 < data.length) {
    const header = data[pos]! | (data[pos + 1]! << 8);
    pos += 2;
    const size = (header & 0x0fff) + 3;
    const compressed = (header & 0x8000) !== 0;
    const end = Math.min(pos + size - 2, data.length);
    if (!compressed) {
      // A raw chunk is exactly its bytes, no tokens.
      for (let i = pos; i < end; i++) out.push(data[i]!);
      pos = end;
      continue;
    }
    const chunkStart = out.length;
    while (pos < end) {
      const flags = data[pos++]!;
      for (let bit = 0; bit < 8 && pos < end; bit++) {
        if ((flags & (1 << bit)) === 0) {
          out.push(data[pos++]!);
          continue;
        }
        if (pos + 1 >= end) break;
        const token = data[pos]! | (data[pos + 1]! << 8);
        pos += 2;
        // The split depends on how much of THIS chunk has been produced so far.
        const difference = out.length - chunkStart;
        let bitCount = 4;
        while (bitCount < 12 && 1 << bitCount < difference) bitCount++;
        const lengthMask = 0xffff >> bitCount;
        const length = (token & lengthMask) + 3;
        const offset = ((token & ~lengthMask & 0xffff) >> (16 - bitCount)) + 1;
        const from = out.length - offset;
        if (from < 0) return new Uint8Array(out); // corrupt: stop rather than invent bytes
        // Overlapping copies are legal and are how runs are encoded, so copy one byte at a time.
        for (let i = 0; i < length; i++) out.push(out[from + i]!);
      }
    }
    pos = end;
  }
  return new Uint8Array(out);
}

/** Decode bytes with the project's code page, falling back to latin1 for the ones we cannot name. */
function decodeText(bytes: Uint8Array, codePage: number): string {
  const label = codePage === 65001 ? "utf-8" : codePage === 1200 ? "utf-16le" : `windows-${codePage}`;
  for (const enc of [label, "windows-1252"]) {
    try {
      return new TextDecoder(enc).decode(bytes);
    } catch {
      /* try the next one */
    }
  }
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return s;
}

/**
 * Parse the decompressed `dir` stream. It is a sequence of records, each a 16-bit id, a 32-bit
 * size and that many bytes. Only the few needed to find the source are read; the rest are skipped
 * by their own length, which is what makes this robust to the parts we do not model.
 */
function parseDir(dir: Uint8Array): { codePage: number; modules: { name: string; stream: string; offset: number; kind: VbaModule["kind"] }[] } {
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  let codePage = 1252;
  const modules: { name: string; stream: string; offset: number; kind: VbaModule["kind"] }[] = [];
  let cur: { name: string; stream: string; offset: number; kind: VbaModule["kind"] } | null = null;
  let pos = 0;
  const ascii = (at: number, len: number): string => decodeText(dir.subarray(at, at + len), codePage);

  while (pos + 6 <= dir.length) {
    const id = dv.getUint16(pos, true);
    let size = dv.getUint32(pos + 2, true);
    const body = pos + 6;
    // PROJECTVERSION's 4-byte field is Reserved, not a size: its body is VersionMajor (4) plus
    // VersionMinor (2). Treating it as a size desynchronises the walk by two bytes and every
    // record after it is then read as garbage.
    if (id === 0x0009) size = 6;
    if (size > dir.length) break; // a size that cannot be right means we have lost the framing
    switch (id) {
      case 0x0003: // PROJECTCODEPAGE
        if (size >= 2) codePage = dv.getUint16(body, true);
        break;
      case 0x0019: // MODULENAME: starts a module record
        if (cur) modules.push(cur);
        cur = { name: ascii(body, size), stream: "", offset: 0, kind: "standard" };
        break;
      case 0x001a: // MODULESTREAMNAME
        if (cur) cur.stream = ascii(body, size);
        break;
      case 0x0031: // MODULEOFFSET: where the compressed source starts in the module stream
        if (cur && size >= 4) cur.offset = dv.getUint32(body, true);
        break;
      case 0x0021: // MODULETYPE procedural
        if (cur) cur.kind = "standard";
        break;
      case 0x0022: // MODULETYPE document / class
        if (cur) cur.kind = "class";
        break;
    }
    pos = body + size;
  }
  if (cur) modules.push(cur);
  return { codePage, modules };
}

/**
 * Read the VBA project out of a vbaProject.bin. Returns undefined when the bytes are not a
 * compound file or carry no VBA storage.
 */
export function readVbaProject(bin: Uint8Array): VbaProject | undefined {
  if (!isCfb(bin)) return undefined;
  let cfb;
  try {
    cfb = readCfb(bin);
  } catch {
    return undefined;
  }
  const streams = cfb.paths();
  const dirEntry = streams.find((s) => /\/vba\/dir$/i.test(s.path));
  if (!dirEntry) return undefined;
  const dir = decompressOvba(cfb.read(dirEntry.index));
  if (!dir.length) return undefined;
  const { codePage, modules } = parseDir(dir);

  const out: VbaModule[] = [];
  for (const m of modules) {
    const streamName = m.stream || m.name;
    const hit = streams.find((s) => new RegExp(`/vba/${streamName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`, "i").test(s.path));
    if (!hit) continue;
    const raw = cfb.read(hit.index);
    if (m.offset >= raw.length) continue;
    const source = decodeText(decompressOvba(raw.subarray(m.offset)), codePage);
    out.push({ name: m.name, stream: streamName, kind: m.kind, source });
  }
  // A locked project still yields its source here; the flag only reflects the author's intent.
  const locked = streams.some((s) => /\/vba\/_VBA_PROJECT$/i.test(s.path)) && /CMG=|DPB=|GC=/.test(decodeText(cfb.readPath("/PROJECT") ?? new Uint8Array(0), codePage));
  return { modules: out, codePage, locked };
}

/** The macro-bearing part of a workbook, if it has one. */
export const vbaPartOf = (files: Record<string, Uint8Array>): Uint8Array | undefined =>
  files["xl/vbaProject.bin"] ?? files["xl/vbaproject.bin"];

/** A rough list of the Sub names a module declares, for showing what a macro file offers. */
export function subNames(source: string): string[] {
  const out: string[] = [];
  for (const line of source.split(/\r?\n/)) {
    const m = /^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*(\(\s*\))?/i.exec(line);
    if (m) out.push(m[1]!);
  }
  return out;
}
