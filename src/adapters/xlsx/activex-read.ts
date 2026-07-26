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

/** The kinds that persist through MorphDataControl, which share one structure and one mask. */
const MORPH_KINDS = new Set<ActiveXKind>(["checkbox", "radio", "textbox", "dropdown", "list", "toggle", "label"]);

/** The class id a class string names, normalised to bare uppercase hex with dashes. */
export const kindOfClsid = (clsid: string): ActiveXKind =>
  KIND_BY_CLSID[clsid.replace(/[{}]/g, "").toUpperCase()] ?? "unknown";

export interface ActiveXControl {
  kind: ActiveXKind;
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

/**
 * Read a persisted ActiveX control stream. Returns the kind for every Forms 2.0 control, and the
 * properties for the ones whose masks are modelled; a control whose layout is not modelled comes
 * back with its kind alone rather than with guessed values.
 */
export function readActiveXStream(bytes: Uint8Array): ActiveXControl | undefined {
  if (bytes.length < 24) return undefined;
  const kind = kindOfClsid(readClsid(bytes));
  if (kind === "unknown") return undefined;
  const out: ActiveXControl = { kind };

  // The control structure follows the class id.
  const at = 16;
  if (bytes[at] !== 0x00 || bytes[at + 1] !== 0x02) return out; // versions the spec pins
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
    const caption = bit(3) ? data.u32() : 0;
    if (bit(4)) data.u32();          // PicturePosition
    if (bit(6)) data.u16();          // MousePointer
    if (bit(7)) data.u16();          // Picture, a placeholder here and the image in StreamData
    if (bit(8)) data.u16();          // Accelerator
    if (bit(10)) data.u16();         // MouseIcon, likewise
    // The ExtraDataBlock starts on a 4-byte boundary after the DataBlock.
    const extraStart = data.position + ((4 - (data.position % 4)) % 4);
    const extra = new Cursor(bytes, extraStart, extraStart);
    if (caption) out.caption = extra.text(caption);
    if (bit(5)) out.size = { cx: extra.i32(), cy: extra.i32() };
    return out;
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
    const value = bit(22) ? block.u32() : 0;
    const caption = bit(23) ? block.u32() : 0;
    for (const [n, width] of [[24, 4], [25, 4], [26, 4], [27, 2], [28, 2], [29, 2]] as [number, number][]) {
      if (bit(n)) { if (width === 2) block.u16(); else block.u32(); }
    }
    const group = bit(32) ? block.u32() : 0;
    // The ExtraDataBlock follows on a 4-byte boundary, Size first and then the strings.
    const extraStart = block.position + ((4 - ((block.position - start) % 4)) % 4);
    const extra = new Cursor(bytes, extraStart, extraStart);
    if (bit(8)) out.size = { cx: extra.i32(), cy: extra.i32() };
    if (value) out.value = extra.text(value);
    if (caption) out.caption = extra.text(caption);
    if (group) out.groupName = extra.text(group);
    // cb states PropMask + DataBlock + ExtraDataBlock. If the walk did not land exactly there,
    // some field's width is wrong and every value read is suspect, so none of them are kept.
    const cb = dv.getUint16(at + 2, true);
    if (extra.position - (at + 4) !== cb) return { kind };
    return out;
  }

  // ScrollBar and SpinButton are NOT parsed for properties. Their mask's contents are published
  // but its bit ORDER could not be confirmed, and the one real sample here has a bit set that is
  // unaccounted for: reading it with a guessed layout produced a width of -1, which is exactly the
  // plausible-but-wrong answer this codebase refuses to ship. The kind is still trustworthy, and
  // the cb check the MorphData path uses would catch it anyway if the layout were ever added.

  // Every other Forms 2.0 kind persists through MorphDataControl, whose mask is a different
  // layout again. Its shape is published, but there is no real file here to check an
  // implementation against, and a property list read with the wrong mask would be plausible and
  // wrong. The kind is from the class id and is trustworthy on its own.
  return out;
}
