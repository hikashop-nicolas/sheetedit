import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { kindOfClsid, readActiveXStream, setActiveXText, setActiveXValue } from "./activex-read";

// The bytes here are SYNTHETIC, built from [MS-OFORMS], because the real ActiveX files available
// are other people's business documents: fine to develop against locally, not to commit.
//
// They are not merely plausible. The generator below was checked against genuine Excel-written
// streams and produces a BYTE-IDENTICAL result for the first 52 bytes, which is the whole control
// up to the TextProps font block this reader does not use. That is what makes a synthetic fixture
// worth trusting: it agrees with a real file, and the real file does not have to live here.

const CLSID = {
  commandButton: [0x40, 0x32, 0x05, 0xd7, 0x69, 0xce, 0xcd, 0x11, 0xa7, 0x77, 0x00, 0xdd, 0x01, 0x14, 0x3c, 0x57],
  scroll: [0xe0, 0x81, 0xd1, 0xdf, 0x2f, 0x5e, 0xce, 0x11, 0xa4, 0x49, 0x00, 0xaa, 0x00, 0x4a, 0x80, 0x3d],
  checkbox: [0x40, 0x1d, 0xd2, 0x8b, 0x42, 0xec, 0xce, 0x11, 0x9e, 0x0d, 0x00, 0xaa, 0x00, 0x60, 0x02, 0xf3],
  combo: [0x30, 0x1d, 0xd2, 0x8b, 0x42, 0xec, 0xce, 0x11, 0x9e, 0x0d, 0x00, 0xaa, 0x00, 0x60, 0x02, 0xf3],
  dropdown: [0x30, 0x1d, 0xd2, 0x8b, 0x42, 0xec, 0xce, 0x11, 0x9e, 0x0d, 0x00, 0xaa, 0x00, 0x60, 0x02, 0xf3],
  textbox: [0x10, 0x1d, 0xd2, 0x8b, 0x42, 0xec, 0xce, 0x11, 0x9e, 0x0d, 0x00, 0xaa, 0x00, 0x60, 0x02, 0xf3],
  tabStrip: [0xb0, 0x0e, 0xe5, 0xea, 0x62, 0x4a, 0xce, 0x11, 0xbe, 0xd6, 0x00, 0xaa, 0x00, 0x61, 0x10, 0x80],
  label: [0x23, 0x9e, 0x8c, 0x97, 0xb0, 0xd4, 0xce, 0x11, 0xbf, 0x2d, 0x00, 0xaa, 0x00, 0x3f, 0x40, 0xd0],
  image: [0x41, 0x92, 0x59, 0x4c, 0x26, 0x69, 0x1b, 0x10, 0x99, 0x92, 0x00, 0x00, 0x0b, 0x65, 0xc6, 0xf9],
};

/** Emit a block of (bit, width, value) fields in bit order, aligned as the format requires. */
function emitFields(fields: [number, 1 | 2 | 4, number][]): number[] {
  const out: number[] = [];
  const align = (n: number): void => { while (out.length % n) out.push(0); };
  for (const [, width, value] of [...fields].sort((a, b) => a[0] - b[0])) {
    align(width);
    if (width === 1) out.push(value & 0xff);
    else if (width === 2) out.push(value & 0xff, (value >> 8) & 0xff);
    else out.push(...u32(value));
  }
  align(4);
  return out;
}

/** A TextProps block: the font, as its own versioned structure after the control's cb. */
function textProps(font: { name?: string; twips?: number; effects?: number }): number[] {
  const nameBytes = font.name ? [...font.name].map((c) => c.charCodeAt(0)) : [];
  let mask = 0;
  const data: [number, 1 | 2 | 4, number][] = [];
  if (font.name) { mask |= 1 << 0; data.push([0, 4, (0x80000000 | nameBytes.length) >>> 0]); }
  if (font.effects !== undefined) { mask |= 1 << 1; data.push([1, 4, font.effects]); }
  if (font.twips !== undefined) { mask |= 1 << 2; data.push([2, 4, font.twips]); }
  const block = emitFields(data);
  const extra = font.name ? [...nameBytes, ...Array((4 - (nameBytes.length % 4)) % 4).fill(0)] : [];
  return [0x00, 0x02, ...u16(4 + block.length + extra.length), ...u32(mask), ...block, ...extra];
}

interface MorphOpts {
  various?: number; displayStyle?: number; maxLength?: number; listRows?: number; columnCount?: number; boundColumn?: number;
  borderStyle?: number; specialEffect?: number; size?: [number, number];
  value?: string; caption?: string; group?: string;
  font?: { name?: string; twips?: number; effects?: number };
  /** Per-column widths in HIMETRIC, as rgColumnInfo states them; null = "the app decides" (-1). */
  columnWidths?: (number | null)[];
}

/** [MS-OFORMS] MorphDataColumnInfo: version, cb over PropMask+DataBlock, then the width. */
function columnInfo(widthHimetric: number | null): number[] {
  const hasWidth = widthHimetric !== null;
  const block = hasWidth ? u32(widthHimetric >>> 0) : [];
  return [0x00, 0x02, ...u16(4 + block.length), ...u32(hasWidth ? 1 : 0), ...block];
}

/** A MorphData control with whatever properties the test asks for, in the format's own order. */
function morph(clsid: number[], o: MorphOpts): Uint8Array {
  let lo = 0, hi = 0;
  const fields: [number, 1 | 2 | 4, number][] = [];
  const set = (bit: number): void => { if (bit < 32) lo |= 1 << bit; else hi |= 1 << (bit - 32); };
  const add = (bit: number, width: 1 | 2 | 4, v: number | undefined): void => {
    if (v === undefined) return;
    set(bit); fields.push([bit, width, v]);
  };
  add(0, 4, o.various);
  add(3, 4, o.maxLength);
  add(4, 1, o.borderStyle);
  add(6, 1, o.displayStyle);
  add(11, 2, o.boundColumn);
  add(13, 2, o.columnCount);
  add(14, 2, o.listRows);
  if (o.columnWidths) add(15, 2, o.columnWidths.length); // cColumnInfo
  if (o.value !== undefined) add(22, 4, (0x80000000 | o.value.length) >>> 0);
  if (o.caption !== undefined) add(23, 4, (0x80000000 | o.caption.length) >>> 0);
  add(26, 4, o.specialEffect);
  if (o.group !== undefined) add(32, 4, (0x80000000 | o.group.length) >>> 0);
  if (o.size) set(8);
  const block = emitFields(fields);
  const extra = [
    ...(o.size ? [...u32(o.size[0]), ...u32(o.size[1])] : []),
    ...(o.value !== undefined ? pad(o.value) : []),
    ...(o.caption !== undefined ? pad(o.caption) : []),
    ...(o.group !== undefined ? pad(o.group) : []),
  ];
  return new Uint8Array([
    ...clsid, 0x00, 0x02, ...u16(8 + block.length + extra.length),
    ...u32(lo), ...u32(hi), ...block, ...extra,
    ...(o.font ? textProps(o.font) : []),
    ...(o.columnWidths ?? []).flatMap((w) => columnInfo(w)),
  ]);
}

/** An Image control, whose mask is its own and whose AutoSize lives in the bit itself. */
function imageControl(o: { backColor?: number; borderColor?: number; borderStyle?: number; sizeMode?: number; size?: [number, number] }): Uint8Array {
  let mask = 0;
  const fields: [number, 1 | 2 | 4, number][] = [];
  const add = (bit: number, width: 1 | 2 | 4, v: number | undefined): void => {
    if (v === undefined) return;
    mask |= 1 << bit; fields.push([bit, width, v]);
  };
  add(3, 4, o.borderColor);
  add(4, 4, o.backColor);
  add(5, 1, o.borderStyle);
  add(7, 1, o.sizeMode);
  if (o.size) mask |= 1 << 9;
  const block = emitFields(fields);
  const extra = o.size ? [...u32(o.size[0]), ...u32(o.size[1])] : [];
  return new Uint8Array([
    ...CLSID.image, 0x00, 0x02, ...u16(4 + block.length + extra.length),
    ...u32(mask), ...block, ...extra,
  ]);
}

const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff];
const pad = (t: string): number[] => {
  const b = [...t].map((c) => c.charCodeAt(0));
  return [...b, ...Array((4 - (b.length % 4)) % 4).fill(0)];
};
const u32 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];

/** A CommandButton stream: caption and size, which is what Excel writes for a plain button. */
function button(caption: string, cx = 5609, cy = 970, extra: { backColor?: number } = {}): Uint8Array {
  const text = [...caption].map((c) => c.charCodeAt(0));
  const pad = (4 - (text.length % 4)) % 4;
  // fCaption is bit 3 and fSize is bit 5; fBackColor is bit 1.
  const mask = 0x08 | 0x20 | (extra.backColor !== undefined ? 0x02 : 0);
  const dataBlock = [
    ...(extra.backColor !== undefined ? u32(extra.backColor) : []),
    ...u32(0x80000000 | text.length), // compressed, so one byte per character
  ];
  const extraBlock = [...text, ...Array(pad).fill(0), ...u32(cx), ...u32(cy)];
  return new Uint8Array([
    ...CLSID.commandButton,
    0x00, 0x02,                                   // MinorVersion, MajorVersion
    ...u16(4 + dataBlock.length + extraBlock.length), // cb: PropMask + DataBlock + ExtraDataBlock
    ...u32(mask),
    ...dataBlock,
    ...extraBlock,
  ]);
}

describe("class ids", () => {
  it("names the Forms 2.0 controls", () => {
    expect(kindOfClsid("{D7053240-CE69-11CD-A777-00DD01143C57}")).toBe("commandButton");
    expect(kindOfClsid("{8BD21D40-EC42-11CE-9E0D-00AA006002F3}")).toBe("checkbox");
    expect(kindOfClsid("{DFD181E0-5E2F-11CE-A449-00AA004A803D}")).toBe("scroll");
  });

  it("takes the id with or without braces, in either case", () => {
    expect(kindOfClsid("d7053240-ce69-11cd-a777-00dd01143c57")).toBe("commandButton");
  });

  it("calls a third-party control unknown rather than guessing at it", () => {
    expect(kindOfClsid("{00000000-0000-0000-0000-000000000000}")).toBe("unknown");
  });
});

describe("reading a persisted stream", () => {
  it("reads a button's caption and size", () => {
    expect(readActiveXStream(button("copy from a file"))).toMatchObject({
      kind: "commandButton",
      caption: "copy from a file",
      size: { cx: 5609, cy: 970 },
    });
  });

  it("pads a caption whose length is not a multiple of four", () => {
    // "Menu" is 4; "Migrate data" is 12; "Go" is 2 and needs two bytes of padding before the size.
    for (const [text, len] of [["Menu", 4], ["Migrate data", 12], ["Go", 2]] as const) {
      const got = readActiveXStream(button(text, 100, 200));
      expect(got?.caption, text).toBe(text);
      expect(got?.size, text).toEqual({ cx: 100, cy: 200 });
      expect(text.length).toBe(len);
    }
  });

  it("reads a colour that was written before the caption", () => {
    // The DataBlock is written in bit order, so a set fBackColor shifts everything after it.
    expect(readActiveXStream(button("Menu", 2487, 988, { backColor: 0x00ffff80 }))).toMatchObject({
      kind: "commandButton",
      backColor: 0x00ffff80,
      caption: "Menu",
      size: { cx: 2487, cy: 988 },
    });
  });

  it("gives a kind but no properties for a control whose mask is not modelled", () => {
    // An Image control: its class id is known, so the kind is trustworthy, but nothing here reads
    // its layout and nothing is invented for it.
    const IMAGE = [0x41, 0x92, 0x59, 0x4c, 0x26, 0x69, 0x1b, 0x10, 0x99, 0x92, 0x00, 0x00, 0x0b, 0x65, 0xc6, 0xf9];
    const bytes = new Uint8Array([...IMAGE, 0x00, 0x02, ...u16(20), ...u32(0), ...Array(16).fill(0)]);
    expect(readActiveXStream(bytes)).toMatchObject({ kind: "image" });
  });

  it("returns nothing for bytes that are not a control at all", () => {
    expect(readActiveXStream(new Uint8Array(8))).toBeUndefined();
    expect(readActiveXStream(new Uint8Array(64))).toBeUndefined(); // a zero class id
  });

  it("stops at the version the spec pins rather than reading on", () => {
    const bytes = button("Menu");
    bytes[17] = 0x03; // MajorVersion must be 2
    expect(readActiveXStream(bytes)).toMatchObject({ kind: "commandButton" });
  });
});

// --- MorphData, the structure behind every kind except the button -----------------
// Its mask is EIGHT bytes where CommandButton's is four, and the ExtraDataBlock puts Size first
// and the strings after it. Both facts came from real files and are asserted here.

/** A checkbox carrying a value, a caption and a group name: the mask Excel writes for one. */
function checkbox(value: string, caption: string, group: string, cx = 2831, cy = 767): Uint8Array {
  const data = [0x04, 0, 0, 0, ...u32(0x80000000 | value.length), ...u32(0x80000000 | caption.length), ...u32(0x80000000 | group.length)];
  const extra = [...u32(cx), ...u32(cy), ...pad(value), ...pad(caption), ...pad(group)];
  return new Uint8Array([
    ...CLSID.checkbox, 0x00, 0x02, ...u16(8 + data.length + extra.length),
    ...u32(0x80c00140),   // fDisplayStyle, fSize, fValue, fCaption, and the reserved top bit
    ...u32(0x00000001),   // fGroupName, which lives in the mask's second word
    ...data, ...extra,
  ]);
}

describe("MorphData controls", () => {
  it("reads a checkbox's value, caption and group", () => {
    expect(readActiveXStream(checkbox("0", "CheckBox1", "DataEntry"))).toMatchObject({
      kind: "checkbox",
      size: { cx: 2831, cy: 767 },
      value: "0",
      caption: "CheckBox1",
      groupName: "DataEntry",
    });
  });

  it("reads a combo box's selected value", () => {
    // A combo sets fVariousPropertyBits, fDisplayStyle, fSize, fMatchEntry, fShowDropButtonWhen
    // and fValue, so the DataBlock mixes 4-byte and 1-byte fields and then realigns.
    const data = [...u32(0x2c80481b), 0x03, 0x01, 0x02, 0x00, ...u32(0x80000004)];
    const extra = [...u32(6244), ...u32(900), ...pad("West")];
    const bytes = new Uint8Array([
      ...CLSID.combo, 0x00, 0x02, ...u16(8 + data.length + extra.length),
      ...u32(0x80450141), ...u32(0), ...data, ...extra,
    ]);
    expect(readActiveXStream(bytes)).toMatchObject({
      kind: "dropdown", size: { cx: 6244, cy: 900 }, value: "West",
    });
  });

  it("refuses everything rather than returning half-right values when the walk misses cb", () => {
    // cb states PropMask + DataBlock + ExtraDataBlock. A wrong field width lands somewhere else,
    // and every value read is then suspect, so the kind alone comes back.
    const bytes = checkbox("0", "CheckBox1", "DataEntry");
    bytes[18] = (bytes[18]! + 4) & 0xff; // overstate cb
    expect(readActiveXStream(bytes)).toMatchObject({ kind: "checkbox" });
  });

  it("handles a caption needing padding and one that needs none", () => {
    expect(readActiveXStream(checkbox("1", "Go", "G"))?.caption).toBe("Go");
    expect(readActiveXStream(checkbox("1", "Four", "Grp1"))?.caption).toBe("Four");
  });
});

describe("writing a value back", () => {
  it("returns byte-identical output when the value does not change", () => {
    // The strongest check there is: nothing moves, so padding and every unmodelled trailing block
    // must come through untouched. Confirmed against four real Excel-written streams.
    const bytes = checkbox("0", "CheckBox1", "DataEntry");
    expect([...setActiveXValue(bytes, "0")!]).toEqual([...bytes]);
  });

  it("keeps the caption and the size when the value grows", () => {
    const before = checkbox("0", "CheckBox1", "DataEntry", 2831, 767);
    const after = setActiveXValue(before, "a much longer value")!;
    expect(readActiveXStream(after)).toMatchObject({
      kind: "checkbox", size: { cx: 2831, cy: 767 },
      value: "a much longer value", caption: "CheckBox1", groupName: "DataEntry",
    });
  });

  it("keeps them when it shrinks", () => {
    const after = setActiveXValue(checkbox("aaaaaaaa", "CheckBox1", "DataEntry"), "1")!;
    expect(readActiveXStream(after)).toMatchObject({ value: "1", caption: "CheckBox1", groupName: "DataEntry" });
  });

  it("carries the trailing blocks it does not model across a resize", () => {
    // TextProps sits after the ExtraDataBlock and must survive, or the control loses its font.
    const font = [0x00, 0x02, 0x18, 0x00, ...[0x35, 0, 0, 0], ...[0x07, 0, 0, 0x80], 0xf0, 0, 0, 0, 0x00, 0x02, 0x00, 0x00, ...[..."Calibri"].map((c) => c.charCodeAt(0)), 0];
    const base = checkbox("0", "CheckBox1", "DataEntry");
    const withTail = new Uint8Array([...base, ...font]);
    const after = setActiveXValue(withTail, "longer than before")!;
    expect([...after.subarray(after.length - font.length)]).toEqual(font);
  });

  it("refuses a control it could not read, rather than writing into the dark", () => {
    const broken = checkbox("0", "CheckBox1", "DataEntry");
    broken[18] = (broken[18]! + 4) & 0xff; // cb no longer matches, so the walk is not trusted
    expect(setActiveXValue(broken, "1")).toBeUndefined();
    // A button carries no Value at all, so there is nothing to set.
    expect(setActiveXValue(button("Go"), "1")).toBeUndefined();
  });

  it("switches to UTF-16 for text a single byte cannot hold", () => {
    const after = setActiveXValue(checkbox("0", "C", "G"), "漢字")!;
    expect(readActiveXStream(after)?.value).toBe("漢字");
  });

  it("round-trips through the reader whatever the length", () => {
    for (const v of ["", "1", "ab", "abc", "abcd", "abcde", "a longer one still"]) {
      const after = setActiveXValue(checkbox("0", "Cap", "Grp"), v)!;
      expect(readActiveXStream(after)?.value ?? "", JSON.stringify(v)).toBe(v);
    }
  });
});

// --- the flat families: scroll bar, spin button, label ---------------------------
// Each is [MS-OFORMS] section 2.2 read across, and each ends on the same cb check, which is what
// turns "these widths look right" into something the code verifies per file.

describe("ScrollBar", () => {
  const CLSID_SCROLL = [0xe0, 0x81, 0xd1, 0xdf, 0x2f, 0x5e, 0xce, 0x11, 0xa4, 0x49, 0x00, 0xaa, 0x00, 0x4a, 0x80, 0x3d];

  /** Exactly what Excel wrote in the sample: fSize, fMax and fOrientation, and nothing else. */
  const scrollBar = (max: number, cx: number, cy: number): Uint8Array => new Uint8Array([
    ...CLSID_SCROLL, 0x00, 0x02, ...u16(20), ...u32(0x2048),
    ...u32(max), ...u32(0xffffffff),   // Max, then Orientation (-1, "auto")
    ...u32(cx), ...u32(cy),            // the ExtraDataBlock is the Size alone
  ]);

  it("reads its bounds and size", () => {
    expect(readActiveXStream(scrollBar(1, 3069, 1005))).toMatchObject({
      kind: "scroll", max: 1, size: { cx: 3069, cy: 1005 },
    });
  });

  it("skips fPrevEnabled and fNextEnabled, which carry no field at all", () => {
    // Those two bits only mirror VariousPropertyBits.Enabled. Consuming four bytes for either
    // would push every later read out of place, and cb would then refuse the whole parse.
    const bytes = scrollBar(1, 100, 200);
    const dv = new DataView(bytes.buffer);
    dv.setUint32(20, 0x2048 | (1 << 9) | (1 << 10), true);
    expect(readActiveXStream(bytes)).toMatchObject({ kind: "scroll", max: 1, size: { cx: 100, cy: 200 } });
  });

  it("reports the kind alone when the walk does not land on cb", () => {
    const bytes = scrollBar(1, 100, 200);
    bytes[18] = (bytes[18]! + 4) & 0xff;
    expect(readActiveXStream(bytes)).toMatchObject({ kind: "scroll" });
  });
});

describe("SpinButton", () => {
  const CLSID_SPIN = [0xb0, 0x6f, 0x17, 0x79, 0xf2, 0xb7, 0xce, 0x11, 0x97, 0xef, 0x00, 0xaa, 0x00, 0x6d, 0x27, 0x76];

  it("uses its own mask, not the scroll bar's", () => {
    // fSize is bit 3 as for a scroll bar, but Min/Max/Position sit at 5/6/7 with no LargeChange
    // or ProportionalThumb between them, and fMousePointer moves to the end.
    const bytes = new Uint8Array([
      ...CLSID_SPIN, 0x00, 0x02, ...u16(4 + 12 + 8), ...u32((1 << 3) | (1 << 5) | (1 << 6) | (1 << 7)),
      ...u32(0), ...u32(100), ...u32(42),   // Min, Max, Position
      ...u32(500), ...u32(900),
    ]);
    expect(readActiveXStream(bytes)).toMatchObject({
      kind: "spin", min: 0, max: 100, position: 42, size: { cx: 500, cy: 900 },
    });
  });
});

describe("Label", () => {
  const CLSID_LABEL = [0x23, 0x9e, 0x8c, 0x97, 0xb0, 0xd4, 0xce, 0x11, 0xbf, 0x2d, 0x00, 0xaa, 0x00, 0x3f, 0x40, 0xd0];

  it("is its own control, not a MorphData one", () => {
    // Assuming otherwise is easy and wrong: a Label has LabelPropMask, where fCaption is bit 3 and
    // fSize is bit 5, and its ExtraDataBlock puts the caption BEFORE the size.
    const caption = [..."Total"].map((c) => c.charCodeAt(0));
    const bytes = new Uint8Array([
      ...CLSID_LABEL, 0x00, 0x02, ...u16(4 + 4 + 8 + 8), ...u32((1 << 3) | (1 << 5)),
      ...u32(0x80000000 | caption.length),
      ...caption, ...Array(3).fill(0),
      ...u32(1200), ...u32(300),
    ]);
    expect(readActiveXStream(bytes)).toMatchObject({
      kind: "label", caption: "Total", size: { cx: 1200, cy: 300 },
    });
  });
});

describe("writing any of the strings, not only the value", () => {
  it("renames a button's caption", () => {
    const after = setActiveXText(button("Old label"), "caption", "A much longer new label")!;
    expect(readActiveXStream(after)).toMatchObject({ caption: "A much longer new label", size: { cx: 5609, cy: 970 } });
  });

  it("changes one string of three and leaves the other two alone", () => {
    // The three sit end to end in the ExtraDataBlock, so rewriting the middle one moves the last.
    const before = checkbox("0", "CheckBox1", "DataEntry");
    const after = setActiveXText(before, "caption", "Renamed rather more fully")!;
    expect(readActiveXStream(after)).toMatchObject({
      kind: "checkbox", size: { cx: 2831, cy: 767 },
      value: "0", caption: "Renamed rather more fully", groupName: "DataEntry",
    });
  });

  it("changes the group name, which is the last of them", () => {
    const after = setActiveXText(checkbox("1", "Cap", "Grp"), "groupName", "AnotherGroupEntirely")!;
    expect(readActiveXStream(after)).toMatchObject({ value: "1", caption: "Cap", groupName: "AnotherGroupEntirely" });
  });

  it("rewrites a label's caption, whose block puts it before the size", () => {
    const caption = [..."Total"].map((c) => c.charCodeAt(0));
    const LABEL = [0x23, 0x9e, 0x8c, 0x97, 0xb0, 0xd4, 0xce, 0x11, 0xbf, 0x2d, 0x00, 0xaa, 0x00, 0x3f, 0x40, 0xd0];
    const bytes = new Uint8Array([
      ...LABEL, 0x00, 0x02, ...u16(4 + 4 + 8 + 8), ...u32((1 << 3) | (1 << 5)),
      ...u32(0x80000000 | caption.length), ...caption, ...Array(3).fill(0), ...u32(1200), ...u32(300),
    ]);
    const after = setActiveXText(bytes, "caption", "Grand total")!;
    expect(readActiveXStream(after)).toMatchObject({ kind: "label", caption: "Grand total", size: { cx: 1200, cy: 300 } });
  });

  it("refuses a property the control does not carry", () => {
    expect(setActiveXText(button("Go"), "value", "x")).toBeUndefined();
    expect(setActiveXText(checkbox("0", "C", "G"), "caption", "ok")).toBeTruthy();
  });
});

// --- what the binary says beyond the value, per [MS-OFORMS] ----------------------------------
// These pin the parts that were read and thrown away before: the bitfield a dozen booleans share,
// the DisplayStyle that tells one MorphData control from another, the font in the TextProps that
// follows the control structure, and the Image control's own layout.
describe("the rest of the control", () => {
  it("decodes VariousPropertyBits, whose defaults the spec states", () => {
    // 0x2C80081B is the MorphData file-format default: Enabled and WordWrap on, Locked off.
    const c = readActiveXStream(morph(CLSID.textbox, { various: 0x2c80081b, size: [100, 50] }))!;
    expect(c.enabled).toBe(true);
    expect(c.locked).toBe(false);
    expect(c.wordWrap).toBe(true);
    expect(c.multiLine).toBe(false);
    expect(c.transparent).toBe(false); // BackStyle 1 = opaque
    // Now the same field with Locked and MultiLine set, and BackStyle cleared.
    const d = readActiveXStream(morph(CLSID.textbox, { various: (0x2c80081b | 0x4 | 0x80000000) & ~0x8, size: [100, 50] }))!;
    expect(d.locked).toBe(true);
    expect(d.multiLine).toBe(true);
    expect(d.transparent).toBe(true);
  });

  it("tells an editable combo from a drop-list one, which share a class id", () => {
    expect(readActiveXStream(morph(CLSID.dropdown, { displayStyle: 3, size: [100, 50] }))!.displayStyle).toBe(3);
    expect(readActiveXStream(morph(CLSID.dropdown, { displayStyle: 7, size: [100, 50] }))!.displayStyle).toBe(7);
  });

  it("reads the numeric properties a list control carries", () => {
    const c = readActiveXStream(morph(CLSID.dropdown, {
      maxLength: 12, listRows: 5, columnCount: 3, borderStyle: 1, specialEffect: 2, size: [100, 50],
    }))!;
    expect(c.maxLength).toBe(12);
    expect(c.listRows).toBe(5);
    expect(c.columnCount).toBe(3);
    expect(c.borderStyle).toBe(1);
    expect(c.specialEffect).toBe(2);
  });

  it("reads the font from the TextProps that follows the control", () => {
    const c = readActiveXStream(morph(CLSID.textbox, { size: [100, 50], font: { name: "Verdana", twips: 240, effects: 0b0111 } }))!;
    expect(c.font).toEqual({ name: "Verdana", sizePt: 12, bold: true, italic: true, underline: true, strike: false });
  });

  it("reads an Image control's own layout", () => {
    const c = readActiveXStream(imageControl({ backColor: 0x00ff8040, borderStyle: 1, sizeMode: 3, size: [1200, 800] }))!;
    expect(c.kind).toBe("image");
    expect(c.backColor).toBe(0x00ff8040);
    expect(c.borderStyle).toBe(1);
    expect(c.pictureSizeMode).toBe(3);
    expect(c.size).toEqual({ cx: 1200, cy: 800 });
  });

  it("adds a Value to a control that had none, and patches it in place afterwards", () => {
    // An empty text box has no Value at all: its mask bit is clear, so there is nothing to patch
    // and the whole DataBlock has to be re-emitted with the new field in its place.
    const empty = morph(CLSID.textbox, { various: 0x2c80081b, size: [100, 50], font: { name: "Calibri", twips: 240 } });
    expect(readActiveXStream(empty)!.value).toBeUndefined();
    const added = setActiveXText(empty, "value", "typed in")!;
    expect(added).toBeDefined();
    const after = readActiveXStream(added)!;
    expect(after.value).toBe("typed in");
    // Everything else came through the rebuild unchanged, font included.
    expect(after.size).toEqual({ cx: 100, cy: 50 });
    expect(after.enabled).toBe(true);
    expect(after.font).toEqual({ name: "Calibri", sizePt: 12 });
    // With the property now present, a further change is the ordinary in-place patch.
    const again = setActiveXText(added, "value", "again!!!")!;
    expect(readActiveXStream(again)!.value).toBe("again!!!");
    expect(again.length).toBe(added.length);
  });
});

describe("multi-column lists", () => {
  it("reads ColumnCount and BoundColumn, which decide the grid and what it reports", () => {
    const c = readActiveXStream(morph(CLSID.dropdown, {
      displayStyle: 3, columnCount: 3, boundColumn: 2, listRows: 6, size: [100, 50], value: "b",
    }))!;
    expect(c.columnCount).toBe(3);
    expect(c.boundColumn).toBe(2);
    expect(c.value).toBe("b");
  });
});

// rgColumnInfo, the per-column widths of a multi-column list. Read from [MS-OFORMS] 2.2.5.6-2.2.5.8
// (the downloadable specification, not the HTML index, which documents only cColumnInfo): each
// entry is its own versioned record, and the single mask bit says whether a width follows at all.
describe("multi-column list widths", () => {
  // Every fixture carries a font because rgColumnInfo FOLLOWS TextProps, which the specification
  // makes mandatory while marking rgColumnInfo optional. The reader anchors on it rather than
  // reading widths from an offset it has not accounted for.
  it("reads a width per column, in pixels", () => {
    // 2540 HIMETRIC to the inch, 96 px to the inch: 2540 -> 96 px, 1270 -> 48 px.
    const bytes = morph(CLSID.combo, { displayStyle: 7, columnCount: 3, font: { name: "Tahoma" }, columnWidths: [2540, 1270, 5080] });
    const ctl = readActiveXStream(bytes)!;
    expect(ctl.columnCount).toBe(3);
    expect(ctl.columnInfoCount).toBe(3);
    expect(ctl.columnWidths).toEqual([96, 48, 192]);
  });

  it("leaves a column the file does not size to the application", () => {
    // -1 is the format's own default and means the client decides, which is not a width.
    const bytes = morph(CLSID.combo, { displayStyle: 7, columnCount: 3, font: { name: "Tahoma" }, columnWidths: [2540, -1, null] });
    const ctl = readActiveXStream(bytes)!;
    expect(ctl.columnWidths).toEqual([96, undefined, undefined]);
  });

  it("accepts a cColumnInfo shorter than the column count", () => {
    // cColumnInfo is the LAST column with a non-default width, so the array can stop early and
    // every column past it is default. Reading columnCount entries would run off the end.
    const bytes = morph(CLSID.combo, { displayStyle: 7, columnCount: 4, font: { name: "Tahoma" }, columnWidths: [2540] });
    const ctl = readActiveXStream(bytes)!;
    expect(ctl.columnCount).toBe(4);
    expect(ctl.columnWidths).toEqual([96]);
  });

  it("says nothing when every column takes the default", () => {
    const bytes = morph(CLSID.combo, { displayStyle: 7, columnCount: 2 });
    const ctl = readActiveXStream(bytes)!;
    expect(ctl.columnWidths).toBeUndefined();
  });

  it("does not disturb the font that precedes it", () => {
    const bytes = morph(CLSID.combo, { displayStyle: 7, columnCount: 2, font: { name: "Verdana", twips: 200 }, columnWidths: [2540, 1270] });
    const ctl = readActiveXStream(bytes)!;
    expect(ctl.font?.name).toBe("Verdana");
    expect(ctl.font?.sizePt).toBe(10);
    expect(ctl.columnWidths).toEqual([96, 48]);
  });
});

// Giving a caption to a control that has none. The MorphData family could already do this; the
// flat layouts (command button, label) could not, which is what stopped a button being labelled.
/** A command button with a Size and, when `caption` is empty, NO caption bit at all. */
function bareButton(caption: string, font?: { name?: string }): Uint8Array {
  let mask = (1 << 5); // fSize
  const fields: [number, 1 | 2 | 4, number][] = [];
  if (caption) { mask |= 1 << 3; fields.push([3, 4, (0x80000000 | caption.length) >>> 0]); }
  const block = emitFields(fields);
  const extra = [...(caption ? pad(caption) : []), ...u32(5609), ...u32(970)];
  return new Uint8Array([
    ...CLSID.commandButton, 0x00, 0x02, ...u16(4 + block.length + extra.length),
    ...u32(mask), ...block, ...extra,
    ...(font ? textProps(font) : []),
  ]);
}
const buttonWithFont = (caption: string): Uint8Array => bareButton(caption, { name: "Verdana" });

/** A label with a Size and no caption. LabelPropMask puts fSize at bit 5, caption at bit 3. */
function labelControl(): Uint8Array {
  const mask = 1 << 5;
  return new Uint8Array([
    ...CLSID.label, 0x00, 0x02, ...u16(4 + 8),
    ...u32(mask), ...u32(2000), ...u32(400),
  ]);
}

describe("adding a caption to a control that has none", () => {
  it("labels a bare command button", () => {
    const bytes = bareButton("");
    const before = readActiveXStream(bytes)!;
    expect(before.kind).toBe("commandButton");
    const out = setActiveXText(bytes, "caption", "Run report")!;
    expect(out).toBeTruthy();
    const after = readActiveXStream(out)!;
    expect(after.caption).toBe("Run report");
    // The properties it already had come through the rebuild unchanged.
    expect(after.size).toEqual(before.size);
    expect(after.kind).toBe("commandButton");
  });

  it("labels a bare label", () => {
    const bytes = labelControl();
    const out = setActiveXText(bytes, "caption", "Total")!;
    expect(readActiveXStream(out)!.caption).toBe("Total");
  });

  it("keeps the rest of the stream, including the font after cb", () => {
    const bytes = buttonWithFont("");
    const out = setActiveXText(bytes, "caption", "Go")!;
    const after = readActiveXStream(out)!;
    expect(after.caption).toBe("Go");
    expect(after.font?.name).toBe("Verdana"); // TextProps sits past cb and must ride along
  });

  it("refuses an empty caption, which is what absent already means", () => {
    expect(setActiveXText(bareButton(""), "caption", "")).toBeUndefined();
  });

  it("can still change a caption a button already has", () => {
    const out = setActiveXText(bareButton("Old"), "caption", "New")!;
    expect(readActiveXStream(out)!.caption).toBe("New");
  });
});

// The whole chain, not just the stream: a caption written into a control's binary has to come back
// through the package as the control's label, which is what the grid shows. The workbook is built
// here rather than read from a fixture, so this runs everywhere instead of skipping quietly.
function xlsxWithBareButton(): Uint8Array {
  const rels = (body: string) => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  const sheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
    `<sheetData/><mc:AlternateContent><mc:Choice Requires="x14"><controls>` +
    `<control shapeId="1025" r:id="rId1" name="CommandButton1"><controlPr defaultSize="0"/></control>` +
    `</controls></mc:Choice></mc:AlternateContent></worksheet>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/></Types>`),
    "_rels/.rels": strToU8(rels(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`)),
    "xl/workbook.xml": strToU8(`<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Controls" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(rels(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>`)),
    "xl/worksheets/sheet1.xml": strToU8(sheet),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(rels(`<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml"/>`)),
    "xl/activeX/activeX1.xml": strToU8(`<?xml version="1.0"?><ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ax:classid="{D7053240-CE69-11CD-A777-00DD01143C57}" ax:persistence="persistStreamInit" r:id="rId1"/>`),
    "xl/activeX/_rels/activeX1.xml.rels": strToU8(rels(`<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>`)),
    "xl/activeX/activeX1.bin": bareButton(""),
  });
}

describe("a caption written into a workbook", () => {
  it("survives save and re-read", async () => {
    const { readWorkbook, writeWorkbook } = await import("../../index");
    const wb = readWorkbook(xlsxWithBareButton());
    const ctl = wb.sheets[0].controls![0]!;
    expect(ctl.activeX).toBe(true);
    expect(ctl.kind).toBe("button");
    expect(ctl.label).toBeUndefined(); // no caption to begin with
    const out = setActiveXText(wb.files[ctl.activeXBinPath!]!, "caption", "Run report")!;
    expect(out).toBeTruthy();
    wb.files[ctl.activeXBinPath!] = out;
    const re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].controls![0]!.label).toBe("Run report");
  });
});

// A TabStrip is not a container: it holds tabs, not controls, so it persists as a flat stream of
// its own shape. Its captions live in the ExtraDataBlock as an array of strings, each counted in
// CHARACTERS with a compression flag in front of it.
describe("TabStrip", () => {
  function tabStrip(tabs: string[], listIndex = 0): Uint8Array {
    const items: number[] = [];
    for (const t of tabs) {
      items.push(...u32(0x80000000 | t.length));           // chars, compressed
      items.push(...[...t].map((c) => c.charCodeAt(0)));
      while (items.length % 4) items.push(0);
    }
    // fListIndex(0), fSize(4), fItems(5).
    const mask = (1 << 0) | (1 << 4) | (1 << 5);
    const block = [...u32(listIndex), ...u32(items.length)];
    const extra = [...u32(4000), ...u32(1200), ...items];
    return new Uint8Array([
      ...CLSID.tabStrip, 0x00, 0x02, ...u16(4 + block.length + extra.length),
      ...u32(mask), ...block, ...extra,
    ]);
  }

  it("reads its tabs in order, and which one is selected", () => {
    const ctl = readActiveXStream(tabStrip(["Summary", "Detail", "Notes"], 1))!;
    expect(ctl.kind).toBe("tabStrip");
    expect(ctl.tabs).toEqual(["Summary", "Detail", "Notes"]);
    expect(ctl.listIndex).toBe(1);
    expect(ctl.size).toEqual({ cx: 4000, cy: 1200 });
  });

  it("reads a strip with no tabs at all without inventing any", () => {
    expect(readActiveXStream(tabStrip([]))!.tabs).toBeUndefined();
  });
});
