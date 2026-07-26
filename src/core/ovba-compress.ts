// ---------------------------------------------------------------------------
// MS-OVBA compression (Stage 5 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// The inverse of decompressOvba, built from the same published [MS-OVBA] description. The
// decompressor is the oracle: decompress(compress(x)) must equal x for arbitrary input, which is
// what the tests assert.
//
// Shape of the container: a 0x01 signature byte, then chunks. Each chunk holds at most 4096
// DECOMPRESSED bytes and carries a 16-bit header: bits 0-11 are (chunk length - 3) counting the
// header itself, bits 12-14 are a 0b011 signature, and bit 15 says whether tokens were used.
// Copy tokens reach back only within the current chunk, so every match search is bounded by it.

const CHUNK = 4096;

/** The offset/length bit split for a copy token, which widens as the chunk's output grows. */
function bitCountFor(difference: number): number {
  let bits = 4;
  while (bits < 12 && 1 << bits < difference) bits++;
  return bits;
}

/** Compress bytes into an MS-OVBA container. */
export function compressOvba(data: Uint8Array): Uint8Array {
  const out: number[] = [0x01];
  for (let start = 0; start < data.length || start === 0; start += CHUNK) {
    const end = Math.min(start + CHUNK, data.length);
    const body = encodeChunk(data, start, end);
    // A chunk that grew is written raw instead: legal, and the only way to stay inside the
    // 12-bit length field when the input does not compress.
    const raw = body === null || body.length > end - start;
    const bytes = raw ? Array.from(data.subarray(start, end)) : body;
    if (!bytes.length) break;
    const header = 0x3000 | (raw ? 0 : 0x8000) | (bytes.length - 1);
    out.push(header & 0xff, (header >> 8) & 0xff);
    for (const b of bytes) out.push(b);
    if (end >= data.length) break;
  }
  return new Uint8Array(out);
}

/**
 * Encode one chunk's worth of input as flag bytes and tokens. Returns null when the result would
 * not fit the 12-bit length field, which sends the caller to a raw chunk.
 */
function encodeChunk(data: Uint8Array, start: number, end: number): number[] | null {
  const body: number[] = [];
  // Three-byte prefixes seen in this chunk, so a match search does not walk the whole window.
  const buckets = new Map<number, number[]>();
  const hash = (i: number): number => (data[i]! << 16) | (data[i + 1]! << 8) | data[i + 2]!;

  let pos = start;
  while (pos < end) {
    const flagAt = body.length;
    body.push(0);
    let flags = 0;
    for (let bit = 0; bit < 8 && pos < end; bit++) {
      const difference = pos - start;
      const bits = bitCountFor(difference);
      const lengthMask = 0xffff >> bits;
      const maxLength = Math.min(lengthMask + 3, end - pos);
      const maxOffset = Math.min(1 << bits, difference);

      let bestLen = 0;
      let bestOff = 0;
      if (maxLength >= 3 && pos + 2 < end) {
        const candidates = buckets.get(hash(pos));
        // Newest candidates first: a nearer match is never worse and is often longer.
        if (candidates) {
          for (let i = candidates.length - 1; i >= 0; i--) {
            const at = candidates[i]!;
            const offset = pos - at;
            if (offset > maxOffset) break; // the list is ordered, so everything earlier is too far
            let len = 0;
            while (len < maxLength && data[at + len] === data[pos + len]) len++;
            if (len > bestLen) { bestLen = len; bestOff = offset; }
            if (bestLen >= maxLength) break;
          }
        }
      }

      if (bestLen >= 3) {
        const token = (((bestOff - 1) << (16 - bits)) & 0xffff) | (bestLen - 3);
        body.push(token & 0xff, (token >> 8) & 0xff);
        flags |= 1 << bit;
        for (let i = 0; i < bestLen; i++) index(buckets, hash, pos + i, end);
        pos += bestLen;
      } else {
        body.push(data[pos]!);
        index(buckets, hash, pos, end);
        pos++;
      }
      if (body.length > CHUNK) return null;
    }
    body[flagAt] = flags;
  }
  return body.length > CHUNK ? null : body;
}

function index(buckets: Map<number, number[]>, hash: (i: number) => number, at: number, end: number): void {
  if (at + 2 >= end) return;
  const key = hash(at);
  const list = buckets.get(key);
  if (list) list.push(at);
  else buckets.set(key, [at]);
}
