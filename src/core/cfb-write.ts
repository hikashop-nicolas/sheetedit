import type { CfbFile } from "./cfb";

// ---------------------------------------------------------------------------
// Compound File Binary (CFB) writer (Stage 5 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// Rebuilds a compound file from one that was read, with some streams replaced. Every stream that is
// not replaced is carried across byte-for-byte, and the directory keeps the tree the original had:
// same entries, same order, same red-black links, same CLSIDs and timestamps. Only each entry's
// start sector and size are rewritten, because only the layout moves.
//
// That is deliberate. A vbaProject.bin holds parts sheetedit does not model (_VBA_PROJECT, PROJECT,
// PROJECTwm, the srp streams), and rebuilding a directory from scratch would mean inventing values
// for fields whose meaning is not fully public. The reader is the round-trip oracle.

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;
const SECTOR = 512;
const MINI = 64;
const PER_SECTOR = SECTOR / 4; // FAT entries per sector
const MAX_HEADER_DIFAT = 109;

const ceilDiv = (a: number, b: number): number => Math.ceil(a / b);

/**
 * Write a compound file with `overrides` (directory index -> new content) applied.
 * Throws when the result would need a DIFAT beyond the header's 109 slots, which for 512-byte
 * sectors means a file over about 7 MB: a vbaProject.bin is never that.
 */
export function writeCfb(src: CfbFile, overrides: Map<number, Uint8Array> = new Map()): Uint8Array {
  const cutoff = src.miniCutoff || 4096;

  // --- what each stream now holds -------------------------------------------
  const content = new Map<number, Uint8Array>();
  src.entries.forEach((e, i) => {
    if (e.type !== 2) return;
    content.set(i, overrides.get(i) ?? src.read(i));
  });

  const miniIdx: number[] = [];
  const bigIdx: number[] = [];
  for (const [i, data] of content) {
    if (!data.length) continue;
    (data.length < cutoff ? miniIdx : bigIdx).push(i);
  }
  // Keep a stable order so the same input always produces the same bytes.
  miniIdx.sort((a, b) => a - b);
  bigIdx.sort((a, b) => a - b);

  // --- the mini stream -------------------------------------------------------
  const miniStart = new Map<number, number>();
  let miniSectors = 0;
  for (const i of miniIdx) {
    miniStart.set(i, miniSectors);
    miniSectors += ceilDiv(content.get(i)!.length, MINI);
  }
  const miniStream = new Uint8Array(miniSectors * MINI);
  for (const i of miniIdx) miniStream.set(content.get(i)!, miniStart.get(i)! * MINI);

  const miniFatEntries = miniSectors;
  const miniFatSecs = miniFatEntries ? ceilDiv(miniFatEntries, PER_SECTOR) : 0;

  // --- sector budget ---------------------------------------------------------
  const dirBytes = padTo(src.dirBytes, SECTOR);
  const dirSecs = dirBytes.length / SECTOR;
  const miniStreamSecs = ceilDiv(miniStream.length, SECTOR);
  const bigSecs = bigIdx.map((i) => ceilDiv(content.get(i)!.length, SECTOR));
  const dataSecs = dirSecs + miniFatSecs + miniStreamSecs + bigSecs.reduce((a, b) => a + b, 0);

  // The FAT has to describe itself, so its own size feeds back into the total.
  let fatSecs = 0;
  for (;;) {
    const need = ceilDiv(dataSecs + fatSecs, PER_SECTOR);
    if (need === fatSecs) break;
    fatSecs = need;
  }
  if (fatSecs > MAX_HEADER_DIFAT) throw new Error("this compound file is too large for sheetedit to rewrite");

  // --- allocation ------------------------------------------------------------
  // FAT sectors first, so the header's own DIFAT is 0..fatSecs-1 and no DIFAT sector is needed.
  let next = fatSecs;
  const take = (n: number): number => { const at = next; next += n; return n ? at : ENDOFCHAIN; };
  const dirAt = take(dirSecs);
  const miniFatAt = take(miniFatSecs);
  const miniStreamAt = take(miniStreamSecs);
  const bigAt = bigIdx.map((_i, k) => take(bigSecs[k]!));

  const fat = new Uint32Array(fatSecs * PER_SECTOR).fill(FREESECT);
  for (let i = 0; i < fatSecs; i++) fat[i] = FATSECT;
  const runChain = (at: number, n: number): void => {
    for (let i = 0; i < n; i++) fat[at + i] = i === n - 1 ? ENDOFCHAIN : at + i + 1;
  };
  runChain(dirAt, dirSecs);
  if (miniFatSecs) runChain(miniFatAt, miniFatSecs);
  if (miniStreamSecs) runChain(miniStreamAt, miniStreamSecs);
  bigIdx.forEach((_i, k) => runChain(bigAt[k]!, bigSecs[k]!));

  // The mini FAT chains mini sectors the same way, one chain per stream.
  const miniFat = new Uint32Array(miniFatSecs * PER_SECTOR).fill(FREESECT);
  for (const i of miniIdx) {
    const at = miniStart.get(i)!;
    const n = ceilDiv(content.get(i)!.length, MINI);
    for (let k = 0; k < n; k++) miniFat[at + k] = k === n - 1 ? ENDOFCHAIN : at + k + 1;
  }

  // --- directory -------------------------------------------------------------
  const dir = new Uint8Array(dirBytes); // a copy: only start and size are rewritten
  const dv = new DataView(dir.buffer);
  const setEntry = (index: number, start: number, size: number): void => {
    const off = index * 128;
    if (off + 128 > dir.length) return;
    dv.setUint32(off + 116, start >>> 0, true);
    dv.setUint32(off + 120, size >>> 0, true);
    dv.setUint32(off + 124, 0, true); // 512-byte sectors: the high word of the size is always zero
  };
  src.entries.forEach((e, i) => {
    if (e.type === 5) { setEntry(i, miniStreamSecs ? miniStreamAt : ENDOFCHAIN, miniStream.length); return; }
    if (e.type !== 2) return;
    const data = content.get(i)!;
    if (!data.length) { setEntry(i, ENDOFCHAIN, 0); return; }
    const k = bigIdx.indexOf(i);
    setEntry(i, k >= 0 ? bigAt[k]! : miniStart.get(i)!, data.length);
  });

  // --- assemble --------------------------------------------------------------
  const total = fatSecs + dataSecs;
  const out = new Uint8Array(512 + total * SECTOR);
  const ov = new DataView(out.buffer);
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
  ov.setUint16(24, 0x003e, true); // minor version, as Office writes it
  ov.setUint16(26, 3, true);      // major version 3 = 512-byte sectors
  ov.setUint16(28, 0xfffe, true); // little-endian marker
  ov.setUint16(30, 9, true);      // sector shift
  ov.setUint16(32, 6, true);      // mini sector shift
  ov.setUint32(44, fatSecs, true);
  ov.setUint32(48, dirAt, true);
  ov.setUint32(56, cutoff, true);
  ov.setUint32(60, miniFatSecs ? miniFatAt : ENDOFCHAIN, true);
  ov.setUint32(64, miniFatSecs, true);
  ov.setUint32(68, ENDOFCHAIN, true); // no DIFAT sector: every FAT sector fits in the header
  ov.setUint32(72, 0, true);
  for (let i = 0; i < MAX_HEADER_DIFAT; i++) ov.setUint32(76 + i * 4, i < fatSecs ? i : FREESECT, true);

  const at = (sector: number): number => 512 + sector * SECTOR;
  for (let i = 0; i < fatSecs; i++) {
    for (let k = 0; k < PER_SECTOR; k++) ov.setUint32(at(i) + k * 4, fat[i * PER_SECTOR + k]!, true);
  }
  for (let i = 0; i < miniFatSecs; i++) {
    for (let k = 0; k < PER_SECTOR; k++) ov.setUint32(at(miniFatAt + i) + k * 4, miniFat[i * PER_SECTOR + k]!, true);
  }
  if (miniStreamSecs) out.set(miniStream, at(miniStreamAt));
  bigIdx.forEach((i, k) => out.set(content.get(i)!, at(bigAt[k]!)));
  out.set(dir, at(dirAt));
  return out;
}

function padTo(data: Uint8Array, multiple: number): Uint8Array {
  const n = ceilDiv(Math.max(data.length, multiple), multiple) * multiple;
  if (n === data.length) return data;
  const out = new Uint8Array(n);
  out.set(data);
  return out;
}
