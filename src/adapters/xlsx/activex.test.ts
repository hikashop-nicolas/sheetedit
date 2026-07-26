import { describe, expect, it } from "vitest";
import { kindOfClsid, readActiveXStream } from "./activex-read";

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
    // A ScrollBar: its mask's bit order could not be confirmed, so nothing is invented for it.
    const bytes = new Uint8Array([...CLSID.scroll, 0x00, 0x02, ...u16(20), ...u32(0x2048), ...Array(16).fill(0)]);
    expect(readActiveXStream(bytes)).toEqual({ kind: "scroll" });
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
