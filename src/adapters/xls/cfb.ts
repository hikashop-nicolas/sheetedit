// Compound File Binary ([MS-CFB]) container: the OLE2 filesystem a legacy .xls lives in.
// Reading reuses vbalang's parser (already a dependency, for vbaProject.bin); writing builds a
// fresh version 3 file (512-byte sectors) with a mini stream, a FAT and DIFAT chaining.
import { isCfb, readCfb } from "vbalang";

export { isCfb };

/** The top-level streams of a compound file, by name. */
export interface CfbStreams {
  streams: Map<string, Uint8Array>;
  /** Names of top-level storages (e.g. _VBA_PROJECT_CUR, MBD... for embedded objects). */
  storages: string[];
}

/** Read the root-level streams of a compound file. Throws a clear error on a damaged container. */
export function readCfbStreams(bytes: Uint8Array): CfbStreams {
  let cfb: ReturnType<typeof readCfb>;
  try {
    cfb = readCfb(bytes);
  } catch (e) {
    throw new Error(`damaged compound file: ${(e as Error).message}`);
  }
  const streams = new Map<string, Uint8Array>();
  const storages: string[] = [];
  const seen = new Set<number>();
  const walk = (i: number): void => {
    if (i < 0 || i >= cfb.entries.length || seen.has(i)) return;
    seen.add(i);
    const e = cfb.entries[i]!;
    walk(e.left);
    walk(e.right);
    if (e.type === 2 && !streams.has(e.name)) {
      // A size beyond the file is damage: read what the chain actually holds.
      if (e.size > bytes.length) e.size = bytes.length;
      try {
        streams.set(e.name, cfb.read(i));
      } catch (err) {
        throw new Error(`damaged compound file: ${(err as Error).message}`);
      }
    } else if (e.type === 1) storages.push(e.name);
  };
  walk(cfb.entries[0]?.child ?? -1);
  return { streams, storages };
}

const SECTOR = 512;
const MINI = 64;
const MINI_CUTOFF = 4096;
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;

/** CFB name order: shorter names first, then an uppercase code-unit comparison. */
function cmpNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const ua = a.toUpperCase(), ub = b.toUpperCase();
  return ua < ub ? -1 : ua > ub ? 1 : 0;
}

/** Write a compound file holding these root-level streams. */
export function writeCfbStreams(entries: { name: string; data: Uint8Array }[]): Uint8Array {
  const sorted = [...entries].sort((a, b) => cmpNames(a.name, b.name));
  // Mini stream: every stream below the cutoff, in 64-byte mini sectors.
  const miniFat: number[] = [];
  const miniChunks: Uint8Array[] = [];
  const start = new Map<string, number>();
  let miniLen = 0;
  for (const e of sorted) {
    if (e.data.length >= MINI_CUTOFF || e.data.length === 0) continue;
    const n = Math.ceil(e.data.length / MINI);
    const first = miniLen / MINI;
    start.set(e.name, first);
    for (let i = 0; i < n; i++) miniFat.push(i === n - 1 ? ENDOFCHAIN : first + i + 1);
    const padded = new Uint8Array(n * MINI);
    padded.set(e.data);
    miniChunks.push(padded);
    miniLen += n * MINI;
  }
  const miniStream = concat(miniChunks);
  const secs = (len: number): number => Math.ceil(len / SECTOR);
  const dirCount = sorted.length + 1;
  const dirSecs = secs(dirCount * 128);
  const miniFatSecs = secs(miniFat.length * 4);
  const miniStreamSecs = secs(miniStream.length);
  let dataSecs = miniFatSecs + miniStreamSecs + dirSecs;
  for (const e of sorted) if (e.data.length >= MINI_CUTOFF) dataSecs += secs(e.data.length);
  // FAT sectors (each maps 128 sectors) and DIFAT sectors (each lists 127 FAT sectors beyond 109).
  let fatSecs = 1, difatSecs = 0;
  for (;;) {
    const total = dataSecs + fatSecs + difatSecs;
    const needFat = Math.ceil(total / 128);
    const needDifat = needFat > 109 ? Math.ceil((needFat - 109) / 127) : 0;
    if (needFat === fatSecs && needDifat === difatSecs) break;
    fatSecs = needFat;
    difatSecs = needDifat;
  }
  const totalSecs = dataSecs + fatSecs + difatSecs;
  const fat = new Uint32Array(fatSecs * 128).fill(FREESECT);
  const out = new Uint8Array(SECTOR + totalSecs * SECTOR);
  const dv = new DataView(out.buffer);
  let next = 0;
  // Lay out: FAT, DIFAT, then the chains.
  const fatAt: number[] = [];
  for (let i = 0; i < fatSecs; i++) { fat[next] = FATSECT; fatAt.push(next++); }
  const difatAt: number[] = [];
  for (let i = 0; i < difatSecs; i++) { fat[next] = DIFSECT; difatAt.push(next++); }
  const chain = (data: Uint8Array, minSecs = 0): number => {
    const n = Math.max(secs(data.length), minSecs);
    if (!n) return ENDOFCHAIN;
    const first = next;
    for (let i = 0; i < n; i++) fat[next + i] = i === n - 1 ? ENDOFCHAIN : next + i + 1;
    out.set(data, SECTOR + first * SECTOR);
    next += n;
    return first;
  };
  const mfBytes = new Uint8Array(miniFatSecs * SECTOR).fill(0xff);
  const mfv = new DataView(mfBytes.buffer);
  miniFat.forEach((v, i) => mfv.setUint32(i * 4, v, true));
  const miniFatStart = miniFat.length ? chain(mfBytes) : ENDOFCHAIN;
  const miniStreamStart = miniStream.length ? chain(miniStream) : ENDOFCHAIN;
  for (const e of sorted) if (e.data.length >= MINI_CUTOFF) start.set(e.name, chain(e.data));
  const dirStart = next;
  for (let i = 0; i < dirSecs; i++) fat[next + i] = i === dirSecs - 1 ? ENDOFCHAIN : next + i + 1;
  next += dirSecs;
  // Directory: root, then the streams as a balanced red-black tree under it.
  const dirBase = SECTOR + dirStart * SECTOR;
  const entry = (idx: number, name: string, type: number, color: number, left: number, right: number, child: number, startSec: number, size: number): void => {
    const o = dirBase + idx * 128;
    for (let i = 0; i < name.length && i < 31; i++) dv.setUint16(o + i * 2, name.charCodeAt(i), true);
    dv.setUint16(o + 64, (Math.min(name.length, 31) + 1) * 2, true);
    out[o + 66] = type;
    out[o + 67] = color;
    dv.setInt32(o + 68, left, true);
    dv.setInt32(o + 72, right, true);
    dv.setInt32(o + 76, child, true);
    dv.setUint32(o + 116, startSec, true);
    dv.setUint32(o + 120, size, true);
  };
  for (let i = dirCount; i < dirSecs * 4; i++) {
    const o = dirBase + i * 128;
    dv.setInt32(o + 68, -1, true); dv.setInt32(o + 72, -1, true); dv.setInt32(o + 76, -1, true);
  }
  const depthMax = Math.floor(Math.log2(sorted.length + 1));
  const perfect = (1 << depthMax) - 1 === sorted.length;
  const build = (lo: number, hi: number, depth: number): number => {
    if (lo > hi) return -1;
    const mid = (lo + hi) >> 1;
    const e = sorted[mid]!;
    const left = build(lo, mid - 1, depth + 1);
    const right = build(mid + 1, hi, depth + 1);
    const red = !perfect && depth === depthMax;
    const sz = e.data.length;
    entry(mid + 1, e.name, 2, red ? 0 : 1, left, right, -1, sz ? start.get(e.name)! : ENDOFCHAIN, sz);
    return mid + 1;
  };
  const rootChild = build(0, sorted.length - 1, 0);
  entry(0, "Root Entry", 5, 1, -1, -1, rootChild, miniStreamStart, miniStream.length);
  // FAT and DIFAT sectors.
  fatAt.forEach((s, i) => { for (let j = 0; j < 128; j++) dv.setUint32(SECTOR + s * SECTOR + j * 4, fat[i * 128 + j]!, true); });
  difatAt.forEach((s, i) => {
    const base = SECTOR + s * SECTOR;
    for (let j = 0; j < 127; j++) dv.setUint32(base + j * 4, fatAt[109 + i * 127 + j] ?? FREESECT, true);
    dv.setUint32(base + 508, difatAt[i + 1] ?? ENDOFCHAIN, true);
  });
  // Header.
  [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].forEach((b, i) => (out[i] = b));
  dv.setUint16(24, 0x3e, true);
  dv.setUint16(26, 3, true);
  dv.setUint16(28, 0xfffe, true);
  dv.setUint16(30, 9, true);
  dv.setUint16(32, 6, true);
  dv.setUint32(44, fatSecs, true);
  dv.setUint32(48, dirStart, true);
  dv.setUint32(56, MINI_CUTOFF, true);
  dv.setUint32(60, miniFatStart, true);
  dv.setUint32(64, miniFatSecs, true);
  dv.setUint32(68, difatSecs ? difatAt[0]! : ENDOFCHAIN, true);
  dv.setUint32(72, difatSecs, true);
  for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, fatAt[i] ?? FREESECT, true);
  return out;
}

export function concat(parts: Uint8Array[]): Uint8Array {
  let n = 0;
  for (const p of parts) n += p.length;
  const out = new Uint8Array(n);
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
