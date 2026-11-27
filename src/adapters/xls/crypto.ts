// Legacy .xls encryption ([MS-OFFCRYPTO] 2.3.6 / 2.3.7 and [MS-XLS] 2.2.10): RC4, RC4 CryptoAPI
// and XOR obfuscation. Synchronous MD5 / SHA-1 / RC4, since the BIFF reader is synchronous.

export function md5(data: Uint8Array): Uint8Array {
  const K = new Uint32Array(64), S = [7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21];
  for (let i = 0; i < 64; i++) K[i] = Math.floor(Math.abs(Math.sin(i + 1)) * 2 ** 32) >>> 0;
  const len = data.length, padLen = ((len + 8) >> 6) * 64 + 64;
  const buf = new Uint8Array(padLen);
  buf.set(data);
  buf[len] = 0x80;
  const dv = new DataView(buf.buffer);
  dv.setUint32(padLen - 8, (len * 8) >>> 0, true);
  dv.setUint32(padLen - 4, Math.floor(len / 0x20000000), true);
  let a0 = 0x67452301, b0 = 0xefcdab89, c0 = 0x98badcfe, d0 = 0x10325476;
  const M = new Uint32Array(16);
  for (let off = 0; off < padLen; off += 64) {
    for (let i = 0; i < 16; i++) M[i] = dv.getUint32(off + i * 4, true);
    let A = a0, B = b0, C = c0, D = d0;
    for (let i = 0; i < 64; i++) {
      let F: number, g: number;
      if (i < 16) { F = (B & C) | (~B & D); g = i; }
      else if (i < 32) { F = (D & B) | (~D & C); g = (5 * i + 1) & 15; }
      else if (i < 48) { F = B ^ C ^ D; g = (3 * i + 5) & 15; }
      else { F = C ^ (B | ~D); g = (7 * i) & 15; }
      const s = S[(i >> 4) * 4 + (i & 3)]!;
      const t = (A + F + K[i]! + M[g]!) | 0;
      A = D; D = C; C = B;
      B = (B + ((t << s) | (t >>> (32 - s)))) | 0;
    }
    a0 = (a0 + A) | 0; b0 = (b0 + B) | 0; c0 = (c0 + C) | 0; d0 = (d0 + D) | 0;
  }
  const out = new Uint8Array(16);
  const ov = new DataView(out.buffer);
  [a0, b0, c0, d0].forEach((v, i) => ov.setUint32(i * 4, v >>> 0, true));
  return out;
}

export function sha1(data: Uint8Array): Uint8Array {
  const len = data.length, padLen = ((len + 8) >> 6) * 64 + 64;
  const buf = new Uint8Array(padLen);
  buf.set(data);
  buf[len] = 0x80;
  const dv = new DataView(buf.buffer);
  dv.setUint32(padLen - 4, (len * 8) >>> 0, false);
  dv.setUint32(padLen - 8, Math.floor(len / 0x20000000), false);
  const H = [0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0];
  const W = new Uint32Array(80);
  for (let off = 0; off < padLen; off += 64) {
    for (let i = 0; i < 16; i++) W[i] = dv.getUint32(off + i * 4, false);
    for (let i = 16; i < 80; i++) { const x = W[i - 3]! ^ W[i - 8]! ^ W[i - 14]! ^ W[i - 16]!; W[i] = (x << 1) | (x >>> 31); }
    let [a, b, c, d, e] = H as [number, number, number, number, number];
    for (let i = 0; i < 80; i++) {
      const f = i < 20 ? (b & c) | (~b & d) : i < 40 ? b ^ c ^ d : i < 60 ? (b & c) | (b & d) | (c & d) : b ^ c ^ d;
      const k = i < 20 ? 0x5a827999 : i < 40 ? 0x6ed9eba1 : i < 60 ? 0x8f1bbcdc : 0xca62c1d6;
      const t = (((a << 5) | (a >>> 27)) + f + e + k + W[i]!) | 0;
      e = d; d = c; c = (b << 30) | (b >>> 2); b = a; a = t;
    }
    H[0] = (H[0]! + a) | 0; H[1] = (H[1]! + b) | 0; H[2] = (H[2]! + c) | 0; H[3] = (H[3]! + d) | 0; H[4] = (H[4]! + e) | 0;
  }
  const out = new Uint8Array(20);
  const ov = new DataView(out.buffer);
  H.forEach((v, i) => ov.setUint32(i * 4, v >>> 0, false));
  return out;
}

/** An RC4 keystream. */
export class Rc4 {
  private s = new Uint8Array(256);
  private i = 0;
  private j = 0;
  constructor(key: Uint8Array) {
    const s = this.s;
    for (let i = 0; i < 256; i++) s[i] = i;
    for (let i = 0, j = 0; i < 256; i++) {
      j = (j + s[i]! + key[i % key.length]!) & 255;
      const t = s[i]!; s[i] = s[j]!; s[j] = t;
    }
  }
  /** XOR the keystream into `data` in place (or just advance it when `skip`). */
  apply(data: Uint8Array, from = 0, to = data.length, skip = false): void {
    const s = this.s;
    let { i, j } = this;
    for (let n = from; n < to; n++) {
      i = (i + 1) & 255;
      j = (j + s[i]!) & 255;
      const t = s[i]!; s[i] = s[j]!; s[j] = t;
      if (!skip) data[n]! ^= s[(s[i]! + s[j]!) & 255]!;
    }
    this.i = i; this.j = j;
  }
}

const cat = (...parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
};
const utf16 = (s: string): Uint8Array => {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) { out[i * 2] = s.charCodeAt(i) & 255; out[i * 2 + 1] = s.charCodeAt(i) >> 8; }
  return out;
};
const u32le = (n: number): Uint8Array => new Uint8Array([n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255]);
const eq = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((v, i) => v === b[i]);

/** Something that turns a block number into that block's RC4 key. */
export interface BlockCipher { blockKey(block: number): Uint8Array }

/** RC4 (FILEPASS version 1.1): verify the password and return the key maker, or null. */
export function rc4Standard(password: string, salt: Uint8Array, verifier: Uint8Array, verifierHash: Uint8Array): BlockCipher | null {
  const h0 = md5(utf16(password.slice(0, 255))).subarray(0, 5);
  const parts: Uint8Array[] = [];
  for (let i = 0; i < 16; i++) parts.push(h0, salt);
  const h1 = md5(cat(...parts)).subarray(0, 5);
  const blockKey = (block: number): Uint8Array => md5(cat(h1, u32le(block)));
  const rc = new Rc4(blockKey(0));
  const v = verifier.slice(), vh = verifierHash.slice();
  rc.apply(v);
  rc.apply(vh);
  return eq(md5(v), vh) ? { blockKey } : null;
}

/** RC4 CryptoAPI (FILEPASS version 2/3/4 . 2): verify and return the key maker, or null. */
export function rc4CryptoApi(password: string, salt: Uint8Array, verifier: Uint8Array, verifierHash: Uint8Array, keyBits: number): BlockCipher | null {
  const h0 = sha1(cat(salt, utf16(password.slice(0, 255))));
  const bytes = keyBits && keyBits !== 40 ? keyBits / 8 : 5;
  const blockKey = (block: number): Uint8Array => {
    const k = sha1(cat(h0, u32le(block))).subarray(0, bytes);
    if (bytes !== 5) return k;
    const padded = new Uint8Array(16); // a 40-bit key is used as 128 bits with the rest zeroed
    padded.set(k);
    return padded;
  };
  const rc = new Rc4(blockKey(0));
  const v = verifier.slice(), vh = verifierHash.slice();
  rc.apply(v);
  rc.apply(vh);
  return eq(sha1(v), vh.subarray(0, 20)) ? { blockKey } : null;
}

/** The 16-bit password hash Excel stores for sheet / workbook protection and XOR verification. */
export function passwordHash16(password: string): number {
  const chars = Array.from(password.slice(0, 15), (c) => c.charCodeAt(0) & 0xff);
  let h = 0;
  for (let i = chars.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= chars[i]!;
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  return (h ^ chars.length ^ 0xce4b) & 0xffff;
}

const PAD = [0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00];
const INITIAL = [0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3];
const MATRIX = [
  0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09, 0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf, 0x4563,
  0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0, 0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40, 0xd849, 0xa0b3,
  0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5, 0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496, 0x390d, 0x721a, 0xeb23, 0xc667, 0x9cef,
  0x29ff, 0x53fe, 0xa7fc, 0x5fd9, 0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0, 0xb861, 0x60e3, 0xc1c6, 0x93ad,
  0x377b, 0x6ef6, 0xddec, 0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10, 0xaa51, 0x4483, 0x8906, 0x022d, 0x045a,
  0x08b4, 0x1168, 0x76b4, 0xed68, 0xcaf1, 0x85c3, 0x1ba7, 0x374e, 0x6e9c, 0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6,
  0x1dad, 0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48c4,
];

/** XOR obfuscation (method 1): the 16-bit verifier key and the 16-byte XOR array. */
export function xorKeyAndArray(password: string): { key: number; verifier: number; array: Uint8Array } {
  const pw = Array.from(password.slice(0, 15), (c) => {
    const code = c.charCodeAt(0);
    return code & 0xff ? code & 0xff : code >> 8;
  });
  let key = INITIAL[pw.length - 1] ?? 0;
  let el = 0x68;
  for (let i = pw.length - 1; i >= 0; i--) {
    let ch = pw[i]!;
    for (let j = 0; j < 7; j++) {
      if (ch & 0x40) key ^= MATRIX[el]!;
      ch <<= 1;
      el--;
    }
  }
  const ror = (b: number): number => ((b >> 1) | (b << 7)) & 0xff;
  const xr = (a: number, b: number): number => ror(a ^ b);
  const arr = new Uint8Array(16);
  let idx = pw.length;
  if (idx & 1) {
    arr[idx] = xr(PAD[0]!, key >> 8);
    idx--;
    arr[idx] = xr(pw[pw.length - 1]!, key & 0xff);
  }
  while (idx > 0) {
    idx--;
    arr[idx] = xr(pw[idx]!, key >> 8);
    idx--;
    arr[idx] = xr(pw[idx]!, key & 0xff);
  }
  idx = 15;
  let pad = 15 - pw.length;
  while (pad > 0) {
    arr[idx] = xr(PAD[pad]!, key >> 8);
    idx--;
    pad--;
    arr[idx] = xr(PAD[pad]!, key & 0xff);
    idx--;
    pad--;
  }
  return { key, verifier: passwordHash16(password), array: arr };
}
