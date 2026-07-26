// ---------------------------------------------------------------------------
// Compound File Binary (CFB) reader
// ---------------------------------------------------------------------------
// The container behind vbaProject.bin, legacy .xls, and encrypted OOXML. A CFB is a FAT filesystem
// in a file: a header, a sector-allocation table, and a directory of streams. Only reading is
// implemented, and only what is needed to pull a named stream out.
//
// Built from [MS-CFB]. Sizes are little-endian throughout.

const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const DIFSECT = 0xfffffffc;
const FATSECT = 0xfffffffd;
const MAXREGSECT = 0xfffffffa;

/** Whether these bytes start with the compound-file signature. */
export const isCfb = (b: Uint8Array): boolean => b.length >= 8 && SIG.every((v, i) => b[i] === v);

export interface CfbEntry {
  name: string;
  /** 0 = unallocated, 1 = storage (directory), 2 = stream, 5 = root. */
  type: number;
  size: number;
  start: number;
  /** Indexes into the directory, or -1. */
  left: number;
  right: number;
  child: number;
}

export interface CfbFile {
  entries: CfbEntry[];
  /** The directory's raw sectors, so a writer can keep the tree exactly as it was found. */
  dirBytes: Uint8Array;
  /** Sector size in bytes, and the size below which a stream lives in the mini stream. */
  sectorSize: number;
  miniCutoff: number;
  /** Read one stream's bytes by directory index. */
  read(index: number): Uint8Array;
  /** Find an entry by name (case-insensitive), optionally under a parent storage. */
  find(name: string, parent?: number): number;
  /** The full path of every stream, as "/VBA/Module1". */
  paths(): { path: string; index: number }[];
  /** Read a stream by path, or undefined when there is no such stream. */
  readPath(path: string): Uint8Array | undefined;
}

/** Parse a compound file. Throws when the bytes are not one. */
export function readCfb(bytes: Uint8Array): CfbFile {
  if (!isCfb(bytes)) throw new Error("not a compound file");
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (o: number): number => dv.getUint16(o, true);
  const u32 = (o: number): number => dv.getUint32(o, true);

  const sectorShift = u16(30);
  const miniSectorShift = u16(32);
  const sectorSize = 1 << sectorShift;
  const miniSectorSize = 1 << miniSectorShift;
  const numFatSectors = u32(44);
  const dirStart = u32(48);
  const miniCutoff = u32(56);
  const miniFatStart = u32(60);
  const numMiniFat = u32(64);
  const difatStart = u32(68);
  const numDifat = u32(72);

  // A sector's byte offset: sector 0 begins right after the 512-byte header.
  const sectorOffset = (s: number): number => 512 + s * sectorSize;

  // --- FAT -------------------------------------------------------------------
  // The DIFAT lists the FAT sectors: 109 entries in the header, the rest chained through sectors.
  const fatSectors: number[] = [];
  for (let i = 0; i < Math.min(109, numFatSectors); i++) {
    const s = u32(76 + i * 4);
    if (s <= MAXREGSECT) fatSectors.push(s);
  }
  let next = difatStart;
  for (let n = 0; n < numDifat && next <= MAXREGSECT; n++) {
    const base = sectorOffset(next);
    const perSector = sectorSize / 4 - 1;
    for (let i = 0; i < perSector; i++) {
      const s = u32(base + i * 4);
      if (s <= MAXREGSECT) fatSectors.push(s);
    }
    next = u32(base + perSector * 4);
  }
  const fat: number[] = [];
  for (const s of fatSectors) {
    const base = sectorOffset(s);
    for (let i = 0; i < sectorSize / 4; i++) {
      if (base + i * 4 + 4 > bytes.length) break;
      fat.push(u32(base + i * 4));
    }
  }

  /** Walk a sector chain from `start`, guarding against a loop in a malformed file. */
  const chain = (start: number, table: number[]): number[] => {
    const out: number[] = [];
    const seen = new Set<number>();
    let s = start;
    while (s <= MAXREGSECT && !seen.has(s)) {
      seen.add(s);
      out.push(s);
      const nxt = table[s];
      if (nxt === undefined || nxt === ENDOFCHAIN || nxt === FREESECT || nxt === DIFSECT || nxt === FATSECT) break;
      s = nxt;
    }
    return out;
  };

  const readChain = (start: number, size: number): Uint8Array => {
    const out = new Uint8Array(size);
    let written = 0;
    for (const s of chain(start, fat)) {
      const from = sectorOffset(s);
      const n = Math.min(sectorSize, size - written, Math.max(0, bytes.length - from));
      if (n <= 0) break;
      out.set(bytes.subarray(from, from + n), written);
      written += n;
      if (written >= size) break;
    }
    return out;
  };

  // --- directory -------------------------------------------------------------
  const dirBytes = readChain(dirStart, chain(dirStart, fat).length * sectorSize);
  const dirView = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
  const entries: CfbEntry[] = [];
  for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
    const nameLen = dirView.getUint16(off + 64, true);
    let name = "";
    // The name is UTF-16LE and the length counts bytes, including the terminator.
    for (let i = 0; i + 1 < Math.max(0, nameLen - 2); i += 2) name += String.fromCharCode(dirView.getUint16(off + i, true));
    const type = dirBytes[off + 66]!;
    entries.push({
      name,
      type,
      left: dirView.getInt32(off + 68, true),
      right: dirView.getInt32(off + 72, true),
      child: dirView.getInt32(off + 76, true),
      start: dirView.getUint32(off + 116, true),
      // A stream's size is 64-bit; the high word is only meaningful in 4K-sector files.
      size: dirView.getUint32(off + 120, true) + dirView.getUint32(off + 124, true) * 2 ** 32,
    });
  }

  // --- mini stream -----------------------------------------------------------
  // Streams below the cutoff live inside the root entry's stream, indexed by the mini FAT.
  const miniFat: number[] = [];
  if (numMiniFat > 0 && miniFatStart <= MAXREGSECT) {
    const mf = readChain(miniFatStart, numMiniFat * sectorSize);
    const mv = new DataView(mf.buffer, mf.byteOffset, mf.byteLength);
    for (let i = 0; i + 4 <= mf.length; i += 4) miniFat.push(mv.getUint32(i, true));
  }
  const root = entries[0];
  let miniStream: Uint8Array | undefined;
  const getMiniStream = (): Uint8Array => {
    if (!miniStream) miniStream = root && root.size > 0 ? readChain(root.start, root.size) : new Uint8Array(0);
    return miniStream;
  };

  const read = (index: number): Uint8Array => {
    const e = entries[index];
    if (!e || e.size === 0) return new Uint8Array(0);
    if (e.size >= miniCutoff) return readChain(e.start, e.size);
    const mini = getMiniStream();
    const out = new Uint8Array(e.size);
    let written = 0;
    for (const s of chain(e.start, miniFat)) {
      const from = s * miniSectorSize;
      const n = Math.min(miniSectorSize, e.size - written, Math.max(0, mini.length - from));
      if (n <= 0) break;
      out.set(mini.subarray(from, from + n), written);
      written += n;
      if (written >= e.size) break;
    }
    return out;
  };

  const find = (name: string, parent = 0): number => {
    const want = name.toLowerCase();
    // The directory is a red-black tree per storage; a plain walk is simpler and fast enough here.
    const start = entries[parent]?.child ?? -1;
    const seen = new Set<number>();
    const stack = [start];
    while (stack.length) {
      const i = stack.pop()!;
      if (i < 0 || i >= entries.length || seen.has(i)) continue;
      seen.add(i);
      const e = entries[i]!;
      if (e.name.toLowerCase() === want) return i;
      stack.push(e.left, e.right, e.child);
    }
    return -1;
  };

  const paths = (): { path: string; index: number }[] => {
    const out: { path: string; index: number }[] = [];
    const seen = new Set<number>();
    const walk = (index: number, prefix: string): void => {
      if (index < 0 || index >= entries.length || seen.has(index)) return;
      seen.add(index);
      const e = entries[index]!;
      // Siblings share the parent's prefix; only a storage extends the path.
      walk(e.left, prefix);
      walk(e.right, prefix);
      const here = `${prefix}/${e.name}`;
      if (e.type === 2) out.push({ path: here, index });
      if (e.type === 1 || e.type === 5) walk(e.child, e.type === 5 ? "" : here);
    };
    walk(entries[0]?.child ?? -1, "");
    return out;
  };

  const readPath = (path: string): Uint8Array | undefined => {
    const want = path.toLowerCase();
    const hit = paths().find((p) => p.path.toLowerCase() === want);
    return hit ? read(hit.index) : undefined;
  };

  return { entries, dirBytes, sectorSize, miniCutoff, read, find, paths, readPath };
}
