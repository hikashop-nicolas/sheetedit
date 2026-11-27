// A growable little-endian byte buffer and BIFF8 record framing for the writer.

export class Bytes {
  buf = new Uint8Array(256);
  len = 0;
  private grow(n: number): void {
    if (this.len + n <= this.buf.length) return;
    let cap = this.buf.length * 2;
    while (cap < this.len + n) cap *= 2;
    const next = new Uint8Array(cap);
    next.set(this.buf.subarray(0, this.len));
    this.buf = next;
  }
  u8(v: number): this { this.grow(1); this.buf[this.len++] = v & 255; return this; }
  u16(v: number): this { this.grow(2); this.buf[this.len++] = v & 255; this.buf[this.len++] = (v >> 8) & 255; return this; }
  u32(v: number): this { this.grow(4); for (let i = 0; i < 4; i++) this.buf[this.len++] = (v >>> (8 * i)) & 255; return this; }
  f64(v: number): this {
    this.grow(8);
    new DataView(this.buf.buffer).setFloat64(this.len, v, true);
    this.len += 8;
    return this;
  }
  bytes(b: ArrayLike<number>): this { this.grow(b.length); this.buf.set(b, this.len); this.len += b.length; return this; }
  zeros(n: number): this { this.grow(n); this.buf.fill(0, this.len, this.len + n); this.len += n; return this; }
  /** BIFF8 XLUnicodeString body: flags + characters (compressed when they all fit in 8 bits). */
  chars(s: string): this {
    const wide = /[^\u0000-ÿ]/.test(s);
    this.u8(wide ? 1 : 0);
    for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); if (wide) this.u16(c); else this.u8(c); }
    return this;
  }
  /** XLUnicodeString with a 16-bit count. */
  str16(s: string): this { return this.u16(s.length).chars(s); }
  /** ShortXLUnicodeString with an 8-bit count. */
  str8(s: string): this { s = s.slice(0, 255); return this.u8(s.length).chars(s); }
  /** UTF-16LE characters followed by a null, as hyperlinks spell them. */
  utf16z(s: string): this { for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i)); return this.u16(0); }
  set16(at: number, v: number): void { this.buf[at] = v & 255; this.buf[at + 1] = (v >> 8) & 255; }
  set32(at: number, v: number): void { for (let i = 0; i < 4; i++) this.buf[at + i] = (v >>> (8 * i)) & 255; }
  out(): Uint8Array { return this.buf.subarray(0, this.len); }
}

export const MAX_REC = 8224;

/** A record stream: writes BIFF records, splitting long payloads into CONTINUE records. */
export class RecStream extends Bytes {
  rec(type: number, body?: Bytes | ArrayLike<number>, cont = 0x3c): this {
    const data = body instanceof Bytes ? body.out() : body ? Uint8Array.from(body) : new Uint8Array(0);
    let off = 0;
    let t = type;
    do {
      const n = Math.min(MAX_REC, data.length - off);
      this.u16(t).u16(n).bytes(data.subarray(off, off + n));
      off += n;
      t = cont;
    } while (off < data.length);
    return this;
  }
}

export const body = (): Bytes => new Bytes();
