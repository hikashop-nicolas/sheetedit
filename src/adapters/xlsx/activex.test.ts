import { describe, expect, it } from "vitest";
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
};

const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff];
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
    expect(readActiveXStream(button("copy from a file"))).toEqual({
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
    expect(readActiveXStream(button("Menu", 2487, 988, { backColor: 0x00ffff80 }))).toEqual({
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
    expect(readActiveXStream(bytes)).toEqual({ kind: "image" });
  });

  it("returns nothing for bytes that are not a control at all", () => {
    expect(readActiveXStream(new Uint8Array(8))).toBeUndefined();
    expect(readActiveXStream(new Uint8Array(64))).toBeUndefined(); // a zero class id
  });

  it("stops at the version the spec pins rather than reading on", () => {
    const bytes = button("Menu");
    bytes[17] = 0x03; // MajorVersion must be 2
    expect(readActiveXStream(bytes)).toEqual({ kind: "commandButton" });
  });
});

// --- MorphData, the structure behind every kind except the button -----------------
// Its mask is EIGHT bytes where CommandButton's is four, and the ExtraDataBlock puts Size first
// and the strings after it. Both facts came from real files and are asserted here.

const pad = (t: string): number[] => {
  const b = [...t].map((c) => c.charCodeAt(0));
  return [...b, ...Array((4 - (b.length % 4)) % 4).fill(0)];
};

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
    expect(readActiveXStream(checkbox("0", "CheckBox1", "DataEntry"))).toEqual({
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
    expect(readActiveXStream(bytes)).toEqual({
      kind: "dropdown", size: { cx: 6244, cy: 900 }, value: "West",
    });
  });

  it("refuses everything rather than returning half-right values when the walk misses cb", () => {
    // cb states PropMask + DataBlock + ExtraDataBlock. A wrong field width lands somewhere else,
    // and every value read is then suspect, so the kind alone comes back.
    const bytes = checkbox("0", "CheckBox1", "DataEntry");
    bytes[18] = (bytes[18]! + 4) & 0xff; // overstate cb
    expect(readActiveXStream(bytes)).toEqual({ kind: "checkbox" });
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
    expect(readActiveXStream(after)).toEqual({
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
    expect(readActiveXStream(scrollBar(1, 3069, 1005))).toEqual({
      kind: "scroll", max: 1, size: { cx: 3069, cy: 1005 },
    });
  });

  it("skips fPrevEnabled and fNextEnabled, which carry no field at all", () => {
    // Those two bits only mirror VariousPropertyBits.Enabled. Consuming four bytes for either
    // would push every later read out of place, and cb would then refuse the whole parse.
    const bytes = scrollBar(1, 100, 200);
    const dv = new DataView(bytes.buffer);
    dv.setUint32(20, 0x2048 | (1 << 9) | (1 << 10), true);
    expect(readActiveXStream(bytes)).toEqual({ kind: "scroll", max: 1, size: { cx: 100, cy: 200 } });
  });

  it("reports the kind alone when the walk does not land on cb", () => {
    const bytes = scrollBar(1, 100, 200);
    bytes[18] = (bytes[18]! + 4) & 0xff;
    expect(readActiveXStream(bytes)).toEqual({ kind: "scroll" });
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
    expect(readActiveXStream(bytes)).toEqual({
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
    expect(readActiveXStream(bytes)).toEqual({
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
    expect(readActiveXStream(after)).toEqual({
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
    expect(readActiveXStream(after)).toEqual({ kind: "label", caption: "Grand total", size: { cx: 1200, cy: 300 } });
  });

  it("refuses a property the control does not carry", () => {
    expect(setActiveXText(button("Go"), "value", "x")).toBeUndefined();
    expect(setActiveXText(checkbox("0", "C", "G"), "caption", "ok")).toBeTruthy();
  });
});
