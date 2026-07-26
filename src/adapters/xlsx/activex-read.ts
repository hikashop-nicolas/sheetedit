// ---------------------------------------------------------------------------
// ActiveX controls: the persisted binary, per [MS-OFORMS]
// ---------------------------------------------------------------------------
// An .xlsm keeps each ActiveX control in two parts: `xl/activeX/activeXN.xml`, which names the
// control's class, and `activeXN.bin`, which holds everything else. The XML is nearly empty when
// persistence is `persistStreamInit` (a class id and a relationship, nothing more), so the caption,
// the value and the size are only in the binary.
//
// That binary is NOT opaque, which is the thing worth knowing here. [MS-OFORMS] specifies it, and
// section 2 of that specification is normative, so the Forms 2.0 controls Excel's ActiveX toolbox
// inserts are readable on the same published-spec basis as the VBA container. Only a THIRD-PARTY
// COM control is genuinely beyond reach: OOXML says its content "shall be solely determined by the
// corresponding object", which means the format belongs to whoever wrote the control.
//
// Layout of a persisted stream: a 16-byte class id (the persistStreamInit convention), then the
// control structure itself, which is MinorVersion (1), MajorVersion (1), a byte count (2), a
// 4-byte PropMask, then a DataBlock of the properties 4 bytes or smaller, then an ExtraDataBlock of
// the larger ones. A bit set in the mask means the property is present; a bit clear means it is at
// its file-format default and was not written at all.

/** What a control is, from its class id. */
export type ActiveXKind =
  | "commandButton" | "checkbox" | "radio" | "textbox" | "dropdown" | "list"
  | "toggle" | "label" | "scroll" | "spin" | "image" | "unknown";

/** The Forms 2.0 class ids. Anything else is a control whose binary is its author's business. */
const KIND_BY_CLSID: Record<string, ActiveXKind> = {
  "D7053240-CE69-11CD-A777-00DD01143C57": "commandButton",
  "8BD21D40-EC42-11CE-9E0D-00AA006002F3": "checkbox",
  "8BD21D50-EC42-11CE-9E0D-00AA006002F3": "radio",
  "8BD21D10-EC42-11CE-9E0D-00AA006002F3": "textbox",
  "8BD21D30-EC42-11CE-9E0D-00AA006002F3": "dropdown",
  "8BD21D20-EC42-11CE-9E0D-00AA006002F3": "list",
  "8BD21D60-EC42-11CE-9E0D-00AA006002F3": "toggle",
  "978C9E23-D4B0-11CE-BF2D-00AA003F40D0": "label",
  "DFD181E0-5E2F-11CE-A449-00AA004A803D": "scroll",
  "79176FB0-B7F2-11CE-97EF-00AA006D2776": "spin",
  "4C599241-6926-101B-9992-00000B65C6F9": "image",
};

/**
 * The families whose whole control is a flat field list: a 4-byte mask, a DataBlock in bit order,
 * and an ExtraDataBlock. Each table below is [MS-OFORMS] section 2.2 read straight across, and the
 * `cb` check at the end of the walk is what proves the widths right on any given file.
 *
 * Two subtleties the tables encode. ScrollBar and SpinButton put fPrevEnabled and fNextEnabled in
 * the mask with NO field behind them, since they only mirror VariousPropertyBits.Enabled; and the
 * ExtraDataBlock order differs by family, Caption before Size for a label, Size alone otherwise.
 */
interface SimpleField { bit: number; width: 1 | 2 | 4; prop?: "foreColor" | "backColor" | "min" | "max" | "position" | "caption" }
interface SimpleLayout { fields: SimpleField[]; sizeBit: number; extraOrder: ("size" | "caption")[] }

const SIMPLE_LAYOUTS: Partial<Record<ActiveXKind, SimpleLayout>> = {
  // 0 fForeColor, 1 fBackColor, 2 fVariousPropertyBits, 3 fSize, 4 fMousePointer, 5 fMin, 6 fMax,
  // 7 fPosition, 8 unused, 9 fPrevEnabled, 10 fNextEnabled, 11 fSmallChange, 12 fLargeChange,
  // 13 fOrientation, 14 fProportionalThumb, 15 fDelay, 16 fMouseIcon.
  scroll: {
    sizeBit: 3,
    extraOrder: ["size"],
    fields: [
      { bit: 0, width: 4, prop: "foreColor" }, { bit: 1, width: 4, prop: "backColor" },
      { bit: 2, width: 4 }, { bit: 4, width: 1 },
      { bit: 5, width: 4, prop: "min" }, { bit: 6, width: 4, prop: "max" },
      { bit: 7, width: 4, prop: "position" },
      { bit: 11, width: 4 }, { bit: 12, width: 4 }, { bit: 13, width: 4 },
      { bit: 14, width: 2 }, { bit: 15, width: 4 }, { bit: 16, width: 2 },
    ],
  },
  // The same family but not the same mask: no LargeChange or ProportionalThumb, and fMousePointer
  // moves to the end. 0 fForeColor, 1 fBackColor, 2 fVariousPropertyBits, 3 fSize, 4 unused,
  // 5 fMin, 6 fMax, 7 fPosition, 8 fPrevEnabled, 9 fNextEnabled, 10 fSmallChange, 11 fOrientation,
  // 12 fDelay, 13 fMouseIcon, 14 fMousePointer.
  spin: {
    sizeBit: 3,
    extraOrder: ["size"],
    fields: [
      { bit: 0, width: 4, prop: "foreColor" }, { bit: 1, width: 4, prop: "backColor" },
      { bit: 2, width: 4 },
      { bit: 5, width: 4, prop: "min" }, { bit: 6, width: 4, prop: "max" },
      { bit: 7, width: 4, prop: "position" },
      { bit: 10, width: 4 }, { bit: 11, width: 4 }, { bit: 12, width: 4 },
      { bit: 13, width: 2 }, { bit: 14, width: 1 },
    ],
  },
  // A Label is its own control, NOT a MorphData one, which is easy to assume and wrong.
  // 0 fForeColor, 1 fBackColor, 2 fVariousPropertyBits, 3 fCaption, 4 fPicturePosition, 5 fSize,
  // 6 fMousePointer, 7 fBorderColor, 8 fBorderStyle, 9 fSpecialEffect, 10 fPicture,
  // 11 fAccelerator, 12 fMouseIcon.
  label: {
    sizeBit: 5,
    extraOrder: ["caption", "size"],
    fields: [
      { bit: 0, width: 4, prop: "foreColor" }, { bit: 1, width: 4, prop: "backColor" },
      { bit: 2, width: 4 }, { bit: 3, width: 4, prop: "caption" }, { bit: 4, width: 4 },
      { bit: 6, width: 1 }, { bit: 7, width: 4 }, { bit: 8, width: 2 }, { bit: 9, width: 2 },
      { bit: 10, width: 2 }, { bit: 11, width: 2 }, { bit: 12, width: 2 },
    ],
  },
};

/** The kinds that persist through MorphDataControl, which share one structure and one mask. */
const MORPH_KINDS = new Set<ActiveXKind>(["checkbox", "radio", "textbox", "dropdown", "list", "toggle"]);

/** The class id a class string names, normalised to bare uppercase hex with dashes. */
export const kindOfClsid = (clsid: string): ActiveXKind =>
  KIND_BY_CLSID[clsid.replace(/[{}]/g, "").toUpperCase()] ?? "unknown";

export interface ActiveXControl {
  kind: ActiveXKind;
  /** Range controls (scroll bar, spin button) report their bounds and where they sit. */
  min?: number;
  max?: number;
  position?: number;
  /** The text on the control, when it carries one. */
  caption?: string;
  /** The control's current value, as the file spells it: "0"/"1" for a checkbox, the chosen text
      for a combo box. Left as the stored string rather than coerced, since its meaning is the
      control's own. */
  value?: string;
  /** Which option-button group this belongs to, for the controls that carry one. */
  groupName?: string;
  /** Size in HIMETRIC units (1/100 mm), which is what the format stores. */
  size?: { cx: number; cy: number };
  /** Colours as the OLE_COLOR the file holds, left unconverted; 0x80xxxxxx values are system
      colours whose meaning depends on the theme, so they are not turned into CSS here. */
  backColor?: number;
  foreColor?: number;
}

/** A 16-byte little-endian GUID, in the usual 8-4-4-4-12 text form. */
function readClsid(b: Uint8Array): string {
  const hex = (n: number): string => n.toString(16).toUpperCase().padStart(2, "0");
  const le = (from: number, len: number): string =>
    Array.from({ length: len }, (_v, i) => hex(b[from + len - 1 - i]!)).join("");
  const be = (from: number, len: number): string =>
    Array.from({ length: len }, (_v, i) => hex(b[from + i]!)).join("");
  return `${le(0, 4)}-${le(4, 2)}-${le(6, 2)}-${be(8, 2)}-${be(10, 6)}`;
}

/** Walks a DataBlock / ExtraDataBlock, keeping each read aligned as the format requires. */
class Cursor {
  constructor(private readonly b: Uint8Array, private at: number, private readonly base: number) {}
  private get dv(): DataView { return new DataView(this.b.buffer, this.b.byteOffset, this.b.byteLength); }
  /** Every property is aligned to its own size, relative to the start of its block. */
  private align(n: number): void {
    const off = (this.at - this.base) % n;
    if (off) this.at += n - off;
  }
  get position(): number { return this.at; }
  u8(): number { const v = this.b[this.at]!; this.at += 1; return v; }
  /** Where the next read of this width will land, once alignment is applied. */
  aligned(n: number): number {
    const off = (this.at - this.base) % n;
    return off ? this.at + (n - off) : this.at;
  }
  u16(): number { this.align(2); const v = this.dv.getUint16(this.at, true); this.at += 2; return v; }
  u32(): number { this.align(4); const v = this.dv.getUint32(this.at, true); this.at += 4; return v; }
  i32(): number { this.align(4); const v = this.dv.getInt32(this.at, true); this.at += 4; return v; }
  skip(n: number): void { this.at += n; }
  /**
   * A string whose length and compression flag were recorded in the DataBlock. The high bit means
   * one byte per character rather than UTF-16, and the whole run is padded to a 4-byte boundary.
   */
  text(lengthAndFlag: number): string {
    const compressed = (lengthAndFlag & 0x80000000) !== 0;
    const len = lengthAndFlag & 0x7fffffff;
    if (len === 0 || this.at + len > this.b.length) return "";
    const bytes = this.b.subarray(this.at, this.at + len);
    this.at += len;
    const pad = this.at % 4;
    if (pad) this.at += 4 - pad;
    if (compressed) return new TextDecoder("windows-1252").decode(bytes);
    let s = "";
    for (let i = 0; i + 1 < bytes.length; i += 2) s += String.fromCharCode(bytes[i]! | (bytes[i + 1]! << 8));
    return s;
  }
}

/** A string property in the ExtraDataBlock: where its bytes are, and where its length word is. */
interface StringSlot {
  prop: "value" | "caption" | "groupName";
  /** Offset of the 4-byte length-and-flag word in the DataBlock. */
  lenAt: number;
  /** Offset and byte length of the characters themselves, padding excluded. */
  at: number;
  byteLen: number;
}

/** Where the pieces sit, so a write can splice rather than repeat the walk and risk diverging. */
interface Layout {
  /** Every string the control carries, in the order the ExtraDataBlock holds them. */
  strings: StringSlot[];
  /** The ExtraDataBlock's bounds. */
  extraStart: number;
  extraEnd: number;
  /** Where the PropMask starts and how long it is, so cb can be recomputed. */
  maskAt: number;
  maskBytes: number;
}

/**
 * Read a persisted ActiveX control stream. Returns the kind for every Forms 2.0 control, and the
 * properties for the ones whose masks are modelled; a control whose layout is not modelled comes
 * back with its kind alone rather than with guessed values.
 */
export function readActiveXStream(bytes: Uint8Array): ActiveXControl | undefined {
  return walk(bytes)?.control;
}

function walk(bytes: Uint8Array): { control: ActiveXControl; layout?: Layout } | undefined {
  if (bytes.length < 24) return undefined;
  const kind = kindOfClsid(readClsid(bytes));
  if (kind === "unknown") return undefined;
  const out: ActiveXControl = { kind };

  // The control structure follows the class id.
  const at = 16;
  if (bytes[at] !== 0x00 || bytes[at + 1] !== 0x02) return { control: out }; // versions the spec pins
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const mask = dv.getUint32(at + 4, true);
  const dataStart = at + 8;
  const data = new Cursor(bytes, dataStart, dataStart);

  if (kind === "commandButton") {
    // CommandButtonPropMask: 0 fForeColor, 1 fBackColor, 2 fVariousPropertyBits, 3 fCaption,
    // 4 fPicturePosition, 5 fSize, 6 fMousePointer, 7 fPicture, 8 fAccelerator,
    // 9 fTakeFocusOnClick, 10 fMouseIcon.
    const bit = (n: number): boolean => (mask & (1 << n)) !== 0;
    if (bit(0)) out.foreColor = data.u32();
    if (bit(1)) out.backColor = data.u32();
    if (bit(2)) data.u32();          // VariousPropertyBits
    let captionLenAt = -1, caption = 0;
    if (bit(3)) { captionLenAt = data.aligned(4); caption = data.u32(); }
    if (bit(4)) data.u32();          // PicturePosition
    if (bit(6)) data.u16();          // MousePointer
    if (bit(7)) data.u16();          // Picture, a placeholder here and the image in StreamData
    if (bit(8)) data.u16();          // Accelerator
    if (bit(10)) data.u16();         // MouseIcon, likewise
    // The ExtraDataBlock starts on a 4-byte boundary after the DataBlock.
    const extraStart = data.position + ((4 - (data.position % 4)) % 4);
    const extra = new Cursor(bytes, extraStart, extraStart);
    const strings: StringSlot[] = [];
    if (caption) {
      strings.push({ prop: "caption", lenAt: captionLenAt, at: extra.position, byteLen: caption & 0x7fffffff });
      out.caption = extra.text(caption);
    }
    if (bit(5)) out.size = { cx: extra.i32(), cy: extra.i32() };
    if (extra.position - (at + 4) !== dv.getUint16(at + 2, true)) return { control: out };
    return { control: out, layout: { strings, extraStart, extraEnd: extra.position, maskAt: at + 4, maskBytes: 4 } };
  }

  const morph = MORPH_KINDS.has(kind);
  if (morph) {
    // MorphDataPropMask is EIGHT bytes, unlike CommandButton's four, and its bits run:
    // 0 fVariousPropertyBits, 1 fBackColor, 2 fForeColor, 3 fMaxLength, 4 fBorderStyle,
    // 5 fScrollBars, 6 fDisplayStyle, 7 fMousePointer, 8 fSize, 9 fPasswordChar, 10 fListWidth,
    // 11 fBoundColumn, 12 fTextColumn, 13 fColumnCount, 14 fListRows, 15 fcColumnInfo,
    // 16 fMatchEntry, 17 fListStyle, 18 fShowDropButtonWhen, 19 unused, 20 fDropButtonStyle,
    // 21 fMultiSelect, 22 fValue, 23 fCaption, 24 fPicturePosition, 25 fBorderColor,
    // 26 fSpecialEffect, 27 fMouseIcon, 28 fPicture, 29 fAccelerator, 30 unused, 31 reserved,
    // and 32 fGroupName in the second word.
    const lo = dv.getUint32(at + 4, true);
    const hi = dv.getUint32(at + 8, true);
    const bit = (n: number): boolean => (n < 32 ? (lo & (1 << n)) !== 0 : (hi & (1 << (n - 32))) !== 0);
    const start = at + 12;
    const block = new Cursor(bytes, start, start);
    // Every DataBlock field in bit order, at the width the format gives it.
    const WIDTHS: [number, number][] = [
      [0, 4], [1, 4], [2, 4], [3, 4], [4, 1], [5, 1], [6, 1], [7, 1], [9, 2], [10, 4],
      [11, 2], [12, 2], [13, 2], [14, 2], [15, 2], [16, 1], [17, 1], [18, 1], [20, 1], [21, 1],
    ];
    for (const [n, width] of WIDTHS) {
      if (!bit(n)) continue;
      if (width === 1) block.u8();
      else if (width === 2) block.u16();
      else if (n === 1) out.backColor = block.u32();
      else if (n === 2) out.foreColor = block.u32();
      else block.u32();
    }
    // The offsets of the length words are kept so a write can patch them without walking again.
    let valueLenAt = -1, captionLenAt = -1, groupLenAt = -1;
    let value = 0, caption = 0;
    if (bit(22)) { valueLenAt = block.aligned(4); value = block.u32(); }
    if (bit(23)) { captionLenAt = block.aligned(4); caption = block.u32(); }
    for (const [n, width] of [[24, 4], [25, 4], [26, 4], [27, 2], [28, 2], [29, 2]] as [number, number][]) {
      if (bit(n)) { if (width === 2) block.u16(); else block.u32(); }
    }
    let group = 0;
    if (bit(32)) { groupLenAt = block.aligned(4); group = block.u32(); }
    // The ExtraDataBlock follows on a 4-byte boundary, Size first and then the strings.
    const extraStart = block.position + ((4 - ((block.position - start) % 4)) % 4);
    const extra = new Cursor(bytes, extraStart, extraStart);
    const strings: StringSlot[] = [];
    if (bit(8)) out.size = { cx: extra.i32(), cy: extra.i32() };
    const take = (prop: StringSlot["prop"], lenAndFlag: number, lenAt: number): void => {
      if (!lenAndFlag) return;
      strings.push({ prop, lenAt, at: extra.position, byteLen: lenAndFlag & 0x7fffffff });
      const text = extra.text(lenAndFlag);
      if (prop === "value") out.value = text;
      else if (prop === "caption") out.caption = text;
      else out.groupName = text;
    };
    take("value", value, valueLenAt);
    take("caption", caption, captionLenAt);
    take("groupName", group, groupLenAt);
    // cb states PropMask + DataBlock + ExtraDataBlock. If the walk did not land exactly there,
    // some field's width is wrong and every value read is suspect, so none of them are kept.
    const cb = dv.getUint16(at + 2, true);
    if (extra.position - (at + 4) !== cb) return { control: { kind } };
    return { control: out, layout: { strings, extraStart, extraEnd: extra.position, maskAt: at + 4, maskBytes: 8 } };
  }

  const simple = SIMPLE_LAYOUTS[kind];
  if (simple) {
    const mask = dv.getUint32(at + 4, true);
    const bit = (n: number): boolean => (mask & (1 << n)) !== 0;
    const start = at + 8;
    const block = new Cursor(bytes, start, start);
    let captionLenAt = -1, caption = 0;
    for (const f of simple.fields) {
      if (!bit(f.bit)) continue;
      if (f.prop === "caption") { captionLenAt = block.aligned(4); caption = block.u32(); continue; }
      const v = f.width === 1 ? block.u8() : f.width === 2 ? block.u16() : block.i32();
      if (f.prop === "foreColor") out.foreColor = v >>> 0;
      else if (f.prop === "backColor") out.backColor = v >>> 0;
      else if (f.prop === "min") out.min = v;
      else if (f.prop === "max") out.max = v;
      else if (f.prop === "position") out.position = v;
    }
    const extraStart = block.position + ((4 - ((block.position - start) % 4)) % 4);
    const extra = new Cursor(bytes, extraStart, extraStart);
    // The order inside the ExtraDataBlock is not the same across families, so each states its own.
    const strings: StringSlot[] = [];
    for (const what of simple.extraOrder) {
      if (what === "size" && bit(simple.sizeBit)) out.size = { cx: extra.i32(), cy: extra.i32() };
      else if (what === "caption" && caption) {
        strings.push({ prop: "caption", lenAt: captionLenAt, at: extra.position, byteLen: caption & 0x7fffffff });
        out.caption = extra.text(caption);
      }
    }
    // The same guard the MorphData path uses: land on cb or report nothing but the kind.
    if (extra.position - (at + 4) !== dv.getUint16(at + 2, true)) return { control: { kind } };
    return { control: out, layout: { strings, extraStart, extraEnd: extra.position, maskAt: at + 4, maskBytes: 4 } };
  }

  // Anything left is a Forms 2.0 kind with no layout here: its kind is still trustworthy, since
  // that comes from the class id rather than from the binary.
  return { control: out };
}

/** How a string is stored: one byte per character when it fits, UTF-16 otherwise. */
function encodeText(text: string): { bytes: number[]; lengthAndFlag: number } {
  const compressible = [...text].every((c) => c.charCodeAt(0) < 256);
  if (compressible) {
    const bytes = [...text].map((c) => c.charCodeAt(0));
    return { bytes, lengthAndFlag: 0x80000000 | bytes.length };
  }
  const bytes: number[] = [];
  for (const c of text) { const n = c.charCodeAt(0); bytes.push(n & 0xff, (n >> 8) & 0xff); }
  return { bytes, lengthAndFlag: bytes.length };
}

const padTo4 = (bytes: number[]): number[] => [...bytes, ...Array((4 - (bytes.length % 4)) % 4).fill(0)];

/**
 * Change one of a control's persisted strings, returning the new stream.
 *
 * Returns undefined rather than guessing when the control does not carry that property, or when
 * its layout was not understood: the reader refuses a stream whose walk does not land on `cb`, and
 * a write must not proceed where a read would not.
 *
 * Where the new text encodes to the same length as the old, the bytes are patched IN PLACE, so
 * everything else stays byte-identical, padding included. Only a change of length rebuilds the
 * ExtraDataBlock, and then the strings around it and the blocks this does not model (the picture
 * streams, the font properties, a combo's column widths) are carried over from where they sat.
 */
export function setActiveXText(
  bytes: Uint8Array,
  prop: "value" | "caption" | "groupName",
  text: string,
): Uint8Array | undefined {
  const found = walk(bytes);
  const slot = found?.layout?.strings.find((x) => x.prop === prop);
  if (!found?.layout || !slot) return undefined;
  const L = found.layout;
  const next = encodeText(text);

  if (next.bytes.length === slot.byteLen) {
    // Nothing moves, so only the characters themselves change.
    const out = new Uint8Array(bytes);
    out.set(next.bytes, slot.at);
    return out;
  }

  // Rebuild the ExtraDataBlock from its parts, keeping every other run exactly as it was.
  const pieces: number[] = [];
  let cursor = L.extraStart;
  for (const st of L.strings) {
    pieces.push(...Array.from(bytes.subarray(cursor, st.at)));   // whatever preceded it (the Size)
    const run = st === slot ? next.bytes : Array.from(bytes.subarray(st.at, st.at + st.byteLen));
    pieces.push(...padTo4(run));
    cursor = st.at + st.byteLen + ((4 - (st.byteLen % 4)) % 4);
  }
  pieces.push(...Array.from(bytes.subarray(cursor, L.extraEnd)));  // anything after the last string

  const out = new Uint8Array([
    ...Array.from(bytes.subarray(0, L.extraStart)),
    ...pieces,
    ...Array.from(bytes.subarray(L.extraEnd)),   // StreamData, TextProps, rgColumnInfo
  ]);
  const ov = new DataView(out.buffer);
  ov.setUint32(slot.lenAt, next.lengthAndFlag >>> 0, true);
  // cb states PropMask + DataBlock + ExtraDataBlock, so it moves with the block that changed.
  ov.setUint16(L.maskAt - 2, ((L.extraStart - L.maskAt) + pieces.length) & 0xffff, true);
  return out;
}

/** Change a control's persisted Value. A thin name over setActiveXText, which is the general one. */
export const setActiveXValue = (bytes: Uint8Array, value: string): Uint8Array | undefined =>
  setActiveXText(bytes, "value", value);
