// BIFF record plumbing shared by the reader: record iteration, CONTINUE-aware reading, and the
// string shapes ([MS-XLS] 2.5.293-296) that BIFF5 (8-bit, code page) and BIFF8 (UTF-16) use.

export interface BiffRecord {
  type: number;
  /** The record's payload (decrypted when the stream was encrypted). */
  data: Uint8Array;
  /** Offset of the record header in the stream. */
  pos: number;
  /** Payloads of the CONTINUE records that immediately follow, in order. */
  conts: Uint8Array[];
}

export const CONTINUE = 0x3c;

/** Split a BIFF stream into records from `start`, attaching each record's CONTINUEs. */
export function* records(stream: Uint8Array, start = 0): Generator<BiffRecord> {
  let p = start;
  let pending: BiffRecord | null = null;
  while (p + 4 <= stream.length) {
    const type = stream[p]! | (stream[p + 1]! << 8);
    const len = stream[p + 2]! | (stream[p + 3]! << 8);
    if (type === 0 && len === 0) {
      // Padding to the end of the stream (some writers fill the last sector with zeros).
      let rest = p;
      while (rest < stream.length && stream[rest] === 0) rest++;
      if (rest >= stream.length) break;
    }
    const data = stream.subarray(p + 4, Math.min(stream.length, p + 4 + len));
    if (type === CONTINUE && pending) pending.conts.push(data);
    else {
      if (pending) yield pending;
      pending = { type, data, pos: p, conts: [] };
    }
    p += 4 + len;
  }
  if (pending) yield pending;
}

const decoders = new Map<string, TextDecoder>();
/** A TextDecoder for a Windows code page, falling back to windows-1252. */
export function codepageDecoder(cp: number): TextDecoder {
  const label =
    cp === 1200 ? "utf-16le" : cp === 65001 ? "utf-8" : cp === 932 ? "shift_jis" : cp === 936 ? "gbk" : cp === 949 ? "euc-kr" : cp === 950 ? "big5"
      : cp === 10000 || cp === 32768 ? "macintosh" : cp === 367 || cp === 20127 ? "us-ascii" : cp >= 874 && cp <= 1258 ? `windows-${cp}` : cp === 437 || cp === 850 ? "ibm866" : "windows-1252";
  let d = decoders.get(label);
  if (!d) {
    try { d = new TextDecoder(label); } catch { d = new TextDecoder("windows-1252"); }
    decoders.set(label, d);
  }
  return d;
}

const latin1 = (b: Uint8Array): string => {
  let s = "";
  for (let i = 0; i < b.length; i += 4096) s += String.fromCharCode(...b.subarray(i, i + 4096));
  return s;
};
export const utf16le = (b: Uint8Array): string => {
  let s = "";
  for (let i = 0; i + 1 < b.length; i += 2) s += String.fromCharCode(b[i]! | (b[i + 1]! << 8));
  return s;
};

/** A cursor over a record and its CONTINUEs, aware of how BIFF8 strings resume across them. */
export class Cursor {
  private segs: Uint8Array[];
  private seg = 0;
  pos = 0;
  constructor(rec: { data: Uint8Array; conts?: Uint8Array[] } | Uint8Array, public cp = 1252) {
    this.segs = rec instanceof Uint8Array ? [rec] : [rec.data, ...(rec.conts ?? [])];
  }
  private cur(): Uint8Array {
    while (this.seg < this.segs.length - 1 && this.pos >= this.segs[this.seg]!.length) { this.seg++; this.pos = 0; }
    return this.segs[this.seg]!;
  }
  get left(): number {
    let n = (this.segs[this.seg]?.length ?? 0) - this.pos;
    for (let i = this.seg + 1; i < this.segs.length; i++) n += this.segs[i]!.length;
    return n;
  }
  /** Bytes left in the current segment only. */
  get segLeft(): number { return this.cur().length - this.pos; }
  get atSegStart(): boolean { this.cur(); return this.pos === 0 && this.seg > 0; }
  nextSegment(): boolean {
    if (this.seg >= this.segs.length - 1) return false;
    this.seg++;
    this.pos = 0;
    return true;
  }
  u8(): number {
    const s = this.cur();
    if (this.pos >= s.length) throw new RangeError("record too short");
    return s[this.pos++]!;
  }
  u16(): number { return this.u8() | (this.u8() << 8); }
  i16(): number { const v = this.u16(); return v & 0x8000 ? v - 0x10000 : v; }
  u32(): number { return (this.u16() | (this.u16() << 16)) >>> 0; }
  i32(): number { return this.u16() | (this.u16() << 16); }
  f64(): number {
    const b = this.bytes(8);
    return new DataView(b.buffer, b.byteOffset, 8).getFloat64(0, true);
  }
  bytes(n: number): Uint8Array {
    const s = this.cur();
    if (this.pos + n <= s.length) { const out = s.subarray(this.pos, this.pos + n); this.pos += n; return out; }
    const out = new Uint8Array(n);
    for (let i = 0; i < n; i++) out[i] = this.u8();
    return out;
  }
  skip(n: number): void { for (let k = n; k > 0;) { const s = this.cur(); const take = Math.min(k, s.length - this.pos); if (take <= 0) throw new RangeError("record too short"); this.pos += take; k -= take; } }
  /** `cch` characters in BIFF8 form; a CONTINUE boundary restarts with a fresh high-byte flag. */
  chars8(cch: number, high: boolean): string {
    let out = "";
    let left = cch;
    while (left > 0) {
      const s = this.segs[this.seg]!;
      const avail = Math.floor((s.length - this.pos) / (high ? 2 : 1));
      const take = Math.min(left, avail);
      if (take > 0) {
        const b = s.subarray(this.pos, this.pos + take * (high ? 2 : 1));
        out += high ? utf16le(b) : latin1(b);
        this.pos += b.length;
        left -= take;
      }
      if (left > 0) {
        // A string never splits inside a character: the rest resumes after a fresh flag byte.
        if (!this.nextSegment()) break;
        high = (this.u8() & 1) === 1;
      }
    }
    return out;
  }
  /** An 8-bit string in the workbook's code page (BIFF2-7). */
  chars7(cch: number): string {
    return codepageDecoder(this.cp).decode(this.bytes(Math.min(cch, this.left)));
  }
}

export interface RichRun { ich: number; font: number }
export interface XlsString { text: string; runs?: RichRun[] }

/** XLUnicodeRichExtendedString ([MS-XLS] 2.5.293), the SST / LABEL shape. */
export function readRichString(c: Cursor, cchBytes: 1 | 2 = 2): XlsString {
  const cch = cchBytes === 2 ? c.u16() : c.u8();
  const flags = c.u8();
  const cRun = flags & 0x08 ? c.u16() : 0;
  const cbExt = flags & 0x04 ? c.u32() : 0;
  const text = c.chars8(cch, (flags & 1) === 1);
  let runs: RichRun[] | undefined;
  if (cRun) {
    runs = [];
    for (let i = 0; i < cRun && c.left >= 4; i++) runs.push({ ich: c.u16(), font: c.u16() });
  }
  if (cbExt) c.skip(Math.min(cbExt, c.left));
  return runs?.length ? { text, runs } : { text };
}

/** A BIFF8 XLUnicodeString (u16 count) or ShortXLUnicodeString (u8 count), or the BIFF2-7 8-bit form. */
export function readString(c: Cursor, biff: number, cchBytes: 1 | 2 = 2): string {
  if (biff >= 8) {
    const cch = cchBytes === 2 ? c.u16() : c.u8();
    const flags = c.u8();
    return c.chars8(cch, (flags & 1) === 1);
  }
  const cch = cchBytes === 2 ? c.u16() : c.u8();
  return c.chars7(cch);
}

/** BIFF8 string with the count given separately (XLUnicodeStringNoCch). */
export function readStringNoCch(c: Cursor, cch: number, biff: number): string {
  if (biff < 8) return c.chars7(cch);
  const flags = c.u8();
  return c.chars8(cch, (flags & 1) === 1);
}
