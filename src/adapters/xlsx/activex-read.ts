import { metafileKind } from "../../core/metafile";

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
 * [MS-OFORMS] VariousPropertyBits: a dozen boolean properties packed into one 32-bit field. The
 * bit positions are pinned by the spec's own file-format defaults (0x2C80081B for the MorphData
 * family, 0x0080001B for a label), which is a useful check on having read the table right.
 */
function readVariousProperties(out: ActiveXControl, bits: number): void {
  const on = (n: number): boolean => (bits & (1 << n)) !== 0;
  out.enabled = on(1);
  out.locked = on(2);
  out.transparent = !on(3);        // BackStyle: 1 opaque, 0 transparent
  out.columnHeads = on(10);
  out.matchRequired = on(12);
  out.captionLeft = on(13);
  out.editable = on(14);
  out.wordWrap = on(23);
  out.autoSize = on(28);
  out.multiLine = on(31);
}

/**
 * The families whose whole control is a flat field list: a mask, a DataBlock in bit order, and an
 * ExtraDataBlock. Each table below is [MS-OFORMS] section 2.2 read straight across, and the `cb`
 * check at the end of the walk is what proves the widths right on any given file.
 *
 * Two subtleties the tables encode. ScrollBar and SpinButton put fPrevEnabled and fNextEnabled in
 * the mask with NO field behind them, since they only mirror VariousPropertyBits.Enabled; and the
 * ExtraDataBlock order differs by family, Caption before Size for a label, Size alone otherwise.
 */
type FieldSetter = (out: ActiveXControl, v: number) => void;
interface SimpleField { bit: number; width: 1 | 2 | 4; caption?: true; set?: FieldSetter }
interface SimpleLayout {
  fields: SimpleField[];
  sizeBit: number;
  extraOrder: ("size" | "caption")[];
  /** The mask bits whose picture goes in StreamData, in the order it is written. */
  pictureBits: number[];
}

/** Shorthand for "store this number under that name". */
const num = (prop: keyof ActiveXControl): FieldSetter =>
  (out, v) => { (out as unknown as Record<string, unknown>)[prop] = v; };
/** The same, for a colour, which is unsigned however the field was read. */
const color = (prop: "foreColor" | "backColor" | "borderColor"): FieldSetter =>
  (out, v) => { out[prop] = v >>> 0; };
const bool = (prop: keyof ActiveXControl): FieldSetter =>
  (out, v) => { (out as unknown as Record<string, unknown>)[prop] = v !== 0; };
const various: FieldSetter = (out, v) => readVariousProperties(out, v >>> 0);

const SIMPLE_LAYOUTS: Partial<Record<ActiveXKind, SimpleLayout>> = {
  // 0 fForeColor, 1 fBackColor, 2 fVariousPropertyBits, 3 fSize, 4 fMousePointer, 5 fMin, 6 fMax,
  // 7 fPosition, 8 unused, 9 fPrevEnabled, 10 fNextEnabled, 11 fSmallChange, 12 fLargeChange,
  // 13 fOrientation, 14 fProportionalThumb, 15 fDelay, 16 fMouseIcon.
  scroll: {
    sizeBit: 3,
    extraOrder: ["size"],
    pictureBits: [16], // fMouseIcon
    fields: [
      { bit: 0, width: 4, set: color("foreColor") }, { bit: 1, width: 4, set: color("backColor") },
      { bit: 2, width: 4, set: various }, { bit: 4, width: 1, set: num("mousePointer") },
      { bit: 5, width: 4, set: num("min") }, { bit: 6, width: 4, set: num("max") },
      { bit: 7, width: 4, set: num("position") },
      { bit: 11, width: 4, set: num("smallChange") }, { bit: 12, width: 4, set: num("largeChange") },
      { bit: 13, width: 4, set: num("orientation") },
      { bit: 14, width: 2, set: bool("proportionalThumb") }, { bit: 15, width: 4, set: num("delay") },
      { bit: 16, width: 2 },
    ],
  },
  // The same family but not the same mask: no LargeChange or ProportionalThumb, and fMousePointer
  // moves to the end. 0 fForeColor, 1 fBackColor, 2 fVariousPropertyBits, 3 fSize, 4 unused,
  // 5 fMin, 6 fMax, 7 fPosition, 8 fPrevEnabled, 9 fNextEnabled, 10 fSmallChange, 11 fOrientation,
  // 12 fDelay, 13 fMouseIcon, 14 fMousePointer.
  spin: {
    sizeBit: 3,
    extraOrder: ["size"],
    pictureBits: [13], // fMouseIcon
    fields: [
      { bit: 0, width: 4, set: color("foreColor") }, { bit: 1, width: 4, set: color("backColor") },
      { bit: 2, width: 4, set: various },
      { bit: 5, width: 4, set: num("min") }, { bit: 6, width: 4, set: num("max") },
      { bit: 7, width: 4, set: num("position") },
      { bit: 10, width: 4, set: num("smallChange") }, { bit: 11, width: 4, set: num("orientation") },
      { bit: 12, width: 4, set: num("delay") },
      { bit: 13, width: 2 }, { bit: 14, width: 1, set: num("mousePointer") },
    ],
  },
  // A Label is its own control, NOT a MorphData one, which is easy to assume and wrong.
  // 0 fForeColor, 1 fBackColor, 2 fVariousPropertyBits, 3 fCaption, 4 fPicturePosition, 5 fSize,
  // 6 fMousePointer, 7 fBorderColor, 8 fBorderStyle, 9 fSpecialEffect, 10 fPicture,
  // 11 fAccelerator, 12 fMouseIcon.
  label: {
    sizeBit: 5,
    extraOrder: ["caption", "size"],
    pictureBits: [10, 12], // fPicture, fMouseIcon
    fields: [
      { bit: 0, width: 4, set: color("foreColor") }, { bit: 1, width: 4, set: color("backColor") },
      { bit: 2, width: 4, set: various }, { bit: 3, width: 4, caption: true },
      { bit: 4, width: 4, set: num("picturePosition") },
      { bit: 6, width: 1, set: num("mousePointer") }, { bit: 7, width: 4, set: color("borderColor") },
      { bit: 8, width: 2, set: num("borderStyle") }, { bit: 9, width: 2, set: num("specialEffect") },
      { bit: 10, width: 2 }, { bit: 11, width: 2, set: num("accelerator") }, { bit: 12, width: 2 },
    ],
  },
  // ImagePropMask: 0-1 unused, 2 fAutoSize, 3 fBorderColor, 4 fBackColor, 5 fBorderStyle,
  // 6 fMousePointer, 7 fPictureSizeMode, 8 fSpecialEffect, 9 fSize, 10 fPicture,
  // 11 fPictureAlignment, 12 fPictureTiling, 13 fVariousPropertyBits, 14 fMouseIcon.
  // fAutoSize and fPictureTiling are mask-only, like the scroll bar's fPrevEnabled: the bit IS the
  // value, and consuming bytes for them would push every later read out of place.
  image: {
    sizeBit: 9,
    extraOrder: ["size"],
    pictureBits: [10, 14], // fPicture, fMouseIcon
    fields: [
      { bit: 3, width: 4, set: color("borderColor") }, { bit: 4, width: 4, set: color("backColor") },
      { bit: 5, width: 1, set: num("borderStyle") }, { bit: 6, width: 1, set: num("mousePointer") },
      { bit: 7, width: 1, set: num("pictureSizeMode") }, { bit: 8, width: 1, set: num("specialEffect") },
      { bit: 10, width: 2 }, { bit: 11, width: 1, set: num("pictureAlignment") },
      { bit: 13, width: 4, set: various }, { bit: 14, width: 2 },
    ],
  },
};

/** The mask bits that carry their value in the bit itself, with no DataBlock field behind them. */
const MASK_ONLY: Partial<Record<ActiveXKind, { bit: number; set: FieldSetter }[]>> = {
  image: [
    { bit: 2, set: bool("autoSize") },
    { bit: 12, set: bool("pictureTiling") },
  ],
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
  borderColor?: number;

  /** MorphData's DisplayStyle: which of the six controls this actually is. An editable combo (3)
      and a drop-list combo (7) share a class id, so this is the only thing that tells them apart. */
  displayStyle?: number;

  // --- VariousPropertyBits, the bitfield a dozen boolean properties share -----------------------
  enabled?: boolean;
  locked?: boolean;
  /** BackStyle 0 = transparent. */
  transparent?: boolean;
  columnHeads?: boolean;
  matchRequired?: boolean;
  /** Alignment: the caption sits to the LEFT of a checkbox / option button. */
  captionLeft?: boolean;
  editable?: boolean;
  wordWrap?: boolean;
  autoSize?: boolean;
  multiLine?: boolean;

  // --- the rest of the DataBlock, kept as the file spells it ------------------------------------
  maxLength?: number;
  passwordChar?: number;
  borderStyle?: number;
  specialEffect?: number;
  scrollBars?: number;
  listRows?: number;
  listWidth?: number;
  listStyle?: number;
  columnCount?: number;
  boundColumn?: number;
  textColumn?: number;
  multiSelect?: number;
  matchEntry?: number;
  showDropButtonWhen?: number;
  dropButtonStyle?: number;
  mousePointer?: number;
  accelerator?: number;
  picturePosition?: number;
  /** Range controls: how far a click moves the thumb, and which way round they sit. */
  smallChange?: number;
  largeChange?: number;
  orientation?: number;
  delay?: number;
  proportionalThumb?: boolean;
  /** Image control. */
  pictureSizeMode?: number;
  pictureAlignment?: number;
  pictureTiling?: boolean;

  /** The font, from the TextProps that follows the control structure. */
  font?: ActiveXFont;
  /** The embedded picture, when the control carries one and a browser can show it. */
  picture?: ActiveXPicture;
}

/** A control's text properties, from [MS-OFORMS] TextProps. */
export interface ActiveXFont {
  name?: string;
  /** Point size. The file holds twips, which is what FontHeight means. */
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** PARAFORMAT_Alignment: 1 left, 2 right, 3 centre. */
  align?: number;
  /** FontWeight, when stated separately from the Bold effect. */
  weight?: number;
}

/** An embedded picture, already sniffed to a MIME type a browser can render. */
export interface ActiveXPicture {
  mime: string;
  bytes: Uint8Array;
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
  /** For the MorphData family, enough to REBUILD the blocks rather than only patch them, which is
      what adding a property the control does not yet carry needs. */
  morph?: MorphRaw;
}

/** A MorphData control's DataBlock as read: every present field's raw value, by mask bit. */
interface MorphRaw {
  lo: number;
  hi: number;
  /** bit -> the value read and the width it was read at. */
  fields: Map<number, { width: 1 | 2 | 4; value: number }>;
  /** The Size property, when the control carries one. */
  size?: { cx: number; cy: number };
  /** Where the DataBlock starts (just after the 8-byte mask). */
  dataStart: number;
}

/** The DataBlock fields of a MorphData control, in bit order, at the width the format gives each. */
const MORPH_WIDTHS: [number, 1 | 2 | 4][] = [
  [0, 4], [1, 4], [2, 4], [3, 4], [4, 1], [5, 1], [6, 1], [7, 1], [9, 2], [10, 4],
  [11, 2], [12, 2], [13, 2], [14, 2], [15, 2], [16, 1], [17, 1], [18, 1], [20, 1], [21, 1],
  [22, 4], [23, 4],
  [24, 4], [25, 4], [26, 4], [27, 2], [28, 2], [29, 2],
  [32, 4],
];

/** Emit a DataBlock from raw field values, applying the format's own alignment rules. */
function emitMorphData(fields: Map<number, { width: 1 | 2 | 4; value: number }>): number[] {
  const out: number[] = [];
  const align = (n: number): void => { while (out.length % n) out.push(0); };
  for (const [bit, width] of MORPH_WIDTHS) {
    const f = fields.get(bit);
    if (!f) continue;
    align(width);
    const v = f.value >>> 0;
    if (width === 1) out.push(v & 0xff);
    else if (width === 2) out.push(v & 0xff, (v >> 8) & 0xff);
    else out.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >> 24) & 0xff);
  }
  align(4); // the block is padded out so the ExtraDataBlock starts aligned
  return out;
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
    // fPicture (7) and fMouseIcon (10) each put a GuidAndPicture in StreamData, in that order.
    readTrailing(bytes, out, extra.position, (bit(7) ? 1 : 0) + (bit(10) ? 1 : 0));
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
    // Every DataBlock field in bit order, at the width the format gives it, and what it means.
    const FIELDS: SimpleField[] = [
      { bit: 0, width: 4, set: various }, { bit: 1, width: 4, set: color("backColor") },
      { bit: 2, width: 4, set: color("foreColor") }, { bit: 3, width: 4, set: num("maxLength") },
      { bit: 4, width: 1, set: num("borderStyle") }, { bit: 5, width: 1, set: num("scrollBars") },
      { bit: 6, width: 1, set: num("displayStyle") }, { bit: 7, width: 1, set: num("mousePointer") },
      { bit: 9, width: 2, set: num("passwordChar") }, { bit: 10, width: 4, set: num("listWidth") },
      { bit: 11, width: 2, set: num("boundColumn") }, { bit: 12, width: 2, set: num("textColumn") },
      { bit: 13, width: 2, set: num("columnCount") }, { bit: 14, width: 2, set: num("listRows") },
      { bit: 15, width: 2 }, { bit: 16, width: 1, set: num("matchEntry") },
      { bit: 17, width: 1, set: num("listStyle") }, { bit: 18, width: 1, set: num("showDropButtonWhen") },
      { bit: 20, width: 1, set: num("dropButtonStyle") }, { bit: 21, width: 1, set: num("multiSelect") },
    ];
    const raw = new Map<number, { width: 1 | 2 | 4; value: number }>();
    for (const f of FIELDS) {
      if (!bit(f.bit)) continue;
      const v = f.width === 1 ? block.u8() : f.width === 2 ? block.u16() : block.i32();
      raw.set(f.bit, { width: f.width, value: v });
      f.set?.(out, v);
    }
    // The offsets of the length words are kept so a write can patch them without walking again.
    let valueLenAt = -1, captionLenAt = -1, groupLenAt = -1;
    let value = 0, caption = 0;
    if (bit(22)) { valueLenAt = block.aligned(4); value = block.u32(); raw.set(22, { width: 4, value }); }
    if (bit(23)) { captionLenAt = block.aligned(4); caption = block.u32(); raw.set(23, { width: 4, value: caption }); }
    const TAIL: SimpleField[] = [
      { bit: 24, width: 4, set: num("picturePosition") }, { bit: 25, width: 4, set: color("borderColor") },
      { bit: 26, width: 4, set: num("specialEffect") }, { bit: 27, width: 2 },
      { bit: 28, width: 2 }, { bit: 29, width: 2, set: num("accelerator") },
    ];
    for (const f of TAIL) {
      if (!bit(f.bit)) continue;
      const v = f.width === 2 ? block.u16() : block.i32();
      raw.set(f.bit, { width: f.width, value: v });
      f.set?.(out, v);
    }
    let group = 0;
    if (bit(32)) { groupLenAt = block.aligned(4); group = block.u32(); raw.set(32, { width: 4, value: group }); }
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
    // fPicture (28) and fMouseIcon (27) put a GuidAndPicture each in StreamData.
    readTrailing(bytes, out, extra.position, (bit(28) ? 1 : 0) + (bit(27) ? 1 : 0));
    const morph: MorphRaw = { lo, hi, fields: raw, dataStart: start, ...(out.size ? { size: out.size } : {}) };
    return { control: out, layout: { strings, extraStart, extraEnd: extra.position, maskAt: at + 4, maskBytes: 8, morph } };
  }

  const simple = SIMPLE_LAYOUTS[kind];
  if (simple) {
    const mask = dv.getUint32(at + 4, true);
    const bit = (n: number): boolean => (mask & (1 << n)) !== 0;
    const start = at + 8;
    const block = new Cursor(bytes, start, start);
    let captionLenAt = -1, caption = 0;
    for (const m of MASK_ONLY[kind] ?? []) m.set(out, bit(m.bit) ? 1 : 0);
    for (const f of simple.fields) {
      if (!bit(f.bit)) continue;
      if (f.caption) { captionLenAt = block.aligned(4); caption = block.u32(); continue; }
      const v = f.width === 1 ? block.u8() : f.width === 2 ? block.u16() : block.i32();
      f.set?.(out, v);
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
    readTrailing(bytes, out, extra.position, simple.pictureBits.filter(bit).length);
    return { control: out, layout: { strings, extraStart, extraEnd: extra.position, maskAt: at + 4, maskBytes: 4 } };
  }

  // Anything left is a Forms 2.0 kind with no layout here: its kind is still trustworthy, since
  // that comes from the class id rather than from the binary.
  return { control: out };
}

// --- what follows the control structure -------------------------------------------------------
// After `cb` bytes of PropMask + DataBlock + ExtraDataBlock come, in order: StreamData (the
// pictures, each a GuidAndPicture), then TextProps (the font), then rgColumnInfo for a multi-column
// list. Each is self-describing, so a reader that knows where cb ended can walk them.

/** The image formats a browser can render straight from bytes, by their magic numbers. */
function pictureMime(b: Uint8Array): string | undefined {
  if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return "image/png";
  if (b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return "image/jpeg";
  if (b.length > 3 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return "image/gif";
  if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4d) return "image/bmp";
  if (b.length > 4 && b[0] === 0x00 && b[1] === 0x00 && b[2] === 0x01 && b[3] === 0x00) return "image/x-icon";
  // A Windows or Enhanced Metafile is a recorded list of drawing calls rather than an image. It is
  // named as one here and replayed onto a canvas at render time; see core/metafile.ts.
  const meta = metafileKind(b);
  if (meta) return meta === "wmf" ? "image/wmf" : "image/emf";
  return undefined;
}

/**
 * A GuidAndPicture: a 16-byte GUID then a StdPicture (preamble 0x0000746C, a size, and the bytes).
 * Returns the picture and where it ended, or just the end when the format is one a page cannot show.
 */
function readGuidAndPicture(b: Uint8Array, at: number): { picture?: ActiveXPicture; end: number } | undefined {
  if (at + 24 > b.length) return undefined;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (dv.getUint32(at + 16, true) !== 0x0000746c) return undefined; // not a StdPicture after all
  const size = dv.getUint32(at + 20, true);
  const start = at + 24;
  if (size === 0 || start + size > b.length) return undefined;
  const data = b.subarray(start, start + size);
  const mime = pictureMime(data);
  return { ...(mime ? { picture: { mime, bytes: data } } : {}), end: start + size };
}

/** [MS-OFORMS] TextProps: the control's font, as its own little versioned structure. */
function readTextProps(b: Uint8Array, at: number): { font: ActiveXFont; end: number } | undefined {
  if (at + 8 > b.length || b[at] !== 0x00 || b[at + 1] !== 0x02) return undefined;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const cb = dv.getUint16(at + 2, true);
  const maskAt = at + 4;
  if (maskAt + cb > b.length) return undefined;
  const mask = dv.getUint32(maskAt, true);
  const bit = (n: number): boolean => (mask & (1 << n)) !== 0;
  const start = maskAt + 4;
  const block = new Cursor(b, start, start);
  const font: ActiveXFont = {};
  let nameLen = 0;
  // 0 fFontName, 1 fFontEffects, 2 fFontHeight, 3 unused, 4 fFontCharSet, 5 fFontPitchAndFamily,
  // 6 fParagraphAlign, 7 fFontWeight.
  if (bit(0)) nameLen = block.u32();
  if (bit(1)) {
    // fmFontEffects: 0 bold, 1 italic, 2 underline, 3 strikeout.
    const fx = block.u32() >>> 0;
    font.bold = (fx & 1) !== 0;
    font.italic = (fx & 2) !== 0;
    font.underline = (fx & 4) !== 0;
    font.strike = (fx & 8) !== 0;
  }
  if (bit(2)) font.sizePt = block.u32() / 20; // FontHeight is in twips
  if (bit(4)) block.u8();                     // FontCharSet
  if (bit(5)) block.u8();                     // FontPitchAndFamily
  if (bit(6)) font.align = block.u8();        // PARAFORMAT_Alignment: 1 left, 2 right, 3 centre
  if (bit(7)) font.weight = block.u16();
  const extraStart = block.position + ((4 - ((block.position - start) % 4)) % 4);
  const extra = new Cursor(b, extraStart, extraStart);
  if (nameLen) font.name = extra.text(nameLen);
  // The same guard the control itself uses: land on cb or report nothing.
  if (extra.position - maskAt !== cb) return undefined;
  return { font, end: extra.position };
}

/** Walk StreamData then TextProps, filling in whatever they carry. */
function readTrailing(b: Uint8Array, out: ActiveXControl, from: number, pictures: number): void {
  let at = from;
  for (let i = 0; i < pictures; i++) {
    const got = readGuidAndPicture(b, at);
    if (!got) return; // not what we expected, so stop rather than read past into nothing
    if (got.picture && !out.picture) out.picture = got.picture;
    at = got.end;
  }
  const text = readTextProps(b, at);
  if (text && (text.font.name || text.font.sizePt || text.font.bold || text.font.italic || text.font.underline || text.font.strike || text.font.align)) {
    out.font = text.font;
  }
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
  if (!found?.layout) return undefined;
  // The control does not carry that property yet, which is what an empty text box looks like: its
  // mask bit is clear and there is nothing to patch. Adding one means rebuilding both blocks, so
  // that path is taken only where the whole DataBlock was recorded field by field.
  if (!slot) return found.layout.morph ? addMorphString(bytes, found.layout, prop, text) : undefined;
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

/** Which mask bit each MorphData string lives behind. */
const MORPH_STRING_BIT: Record<"value" | "caption" | "groupName", number> = { value: 22, caption: 23, groupName: 32 };

/**
 * Add a string property a MorphData control does not yet carry, by rebuilding its DataBlock and
 * ExtraDataBlock from the fields the read recorded. Splicing would not do: a length word has to
 * land 4-aligned inside the DataBlock, and inserting it shifts every field after it, so the block
 * is re-emitted rather than patched.
 *
 * Everything the read did not model (the picture streams, the font, a combo's column widths) is
 * carried across untouched from where it sat after `cb`.
 */
function addMorphString(
  bytes: Uint8Array,
  L: Layout,
  prop: "value" | "caption" | "groupName",
  text: string,
): Uint8Array | undefined {
  const M = L.morph;
  if (!M || !text) return undefined; // adding an empty string is what "not present" already means
  const bit = MORPH_STRING_BIT[prop];
  const next = encodeText(text);
  const fields = new Map(M.fields);
  fields.set(bit, { width: 4, value: next.lengthAndFlag >>> 0 });

  // The strings in the ExtraDataBlock, in the order the format writes them.
  const order: ("value" | "caption" | "groupName")[] = ["value", "caption", "groupName"];
  const runs: number[] = [];
  if (M.size) {
    const push4 = (n: number): void => { runs.push(n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff); };
    push4(M.size.cx); push4(M.size.cy);
  }
  for (const which of order) {
    if (which === prop) { runs.push(...padTo4(next.bytes)); continue; }
    const had = L.strings.find((x) => x.prop === which);
    if (had) runs.push(...padTo4(Array.from(bytes.subarray(had.at, had.at + had.byteLen))));
  }

  const data = emitMorphData(fields);
  const lo = (M.lo | (bit < 32 ? 1 << bit : 0)) >>> 0;
  const hi = (M.hi | (bit >= 32 ? 1 << (bit - 32) : 0)) >>> 0;
  const mask: number[] = [];
  for (const word of [lo, hi]) mask.push(word & 0xff, (word >> 8) & 0xff, (word >> 16) & 0xff, (word >> 24) & 0xff);

  const cb = mask.length + data.length + runs.length;
  const out = new Uint8Array([
    ...Array.from(bytes.subarray(0, L.maskAt - 2)),
    cb & 0xff, (cb >> 8) & 0xff,
    ...mask, ...data, ...runs,
    ...Array.from(bytes.subarray(L.extraEnd)),  // StreamData, TextProps, rgColumnInfo
  ]);
  // Read it back before handing it over: a rebuild that cannot be parsed is worse than a refusal.
  const check = walk(out);
  const got = prop === "value" ? check?.control.value : prop === "caption" ? check?.control.caption : check?.control.groupName;
  return got === text ? out : undefined;
}
