import { describe, expect, it } from "vitest";
import { isMetafileMime, metafileFromDataUri, metafileKind, metafileSize } from "./metafile";

// A metafile is a recorded list of GDI drawing calls, not an image, so a browser shows nothing for
// one. These cover the part that runs without a canvas: recognising them and working out how big
// they ask to be drawn, which the converter gets wrong for a placeable WMF if left to itself.

const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff];
const i16 = u16;
const u32 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];

/** A placeable WMF header: the 0x9AC6CDD7 key, a bounding box, and units per inch. */
const placeableWmf = (right: number, bottom: number, perInch = 1440): Uint8Array =>
  new Uint8Array([
    ...u32(0x9ac6cdd7), ...u16(0), ...i16(0), ...i16(0), ...i16(right), ...i16(bottom),
    ...u16(perInch), ...u32(0), ...u16(0),
    ...u16(1), ...u16(9), ...u16(0x0300), ...u32(0), ...u16(0), ...u32(0), ...u16(0), ...u32(0),
  ]);

/** An EMF header: record type 1, the " EMF" signature 40 bytes in, bounds and a frame in 0.01mm. */
function emfHeader(frameW: number, frameH: number): Uint8Array {
  const b = new Uint8Array(88);
  const dv = new DataView(b.buffer);
  dv.setUint32(0, 1, true);          // EMR_HEADER
  dv.setUint32(4, 88, true);         // record size
  dv.setInt32(8, 0, true); dv.setInt32(12, 0, true);
  dv.setInt32(16, 1000, true); dv.setInt32(20, 1000, true);  // rclBounds
  dv.setInt32(24, 0, true); dv.setInt32(28, 0, true);
  dv.setInt32(32, frameW, true); dv.setInt32(36, frameH, true); // rclFrame, hundredths of a mm
  dv.setUint32(40, 0x464d4520, true); // " EMF"
  return b;
}

describe("metafiles", () => {
  it("knows one by its header, not by the name it arrived under", () => {
    expect(metafileKind(placeableWmf(1000, 500))).toBe("wmf");
    expect(metafileKind(emfHeader(5000, 3000))).toBe("emf");
    // A PNG is not a metafile whatever anyone calls it.
    expect(metafileKind(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0, 0, 0]))).toBeNull();
  });

  it("recognises the MIME types Office writes for them", () => {
    for (const m of ["image/emf", "image/x-emf", "image/wmf", "image/x-wmf", "IMAGE/X-WMF"]) {
      expect(isMetafileMime(m), m).toBe(true);
    }
    expect(isMetafileMime("image/png")).toBe(false);
  });

  it("works out the size a placeable WMF asks for", () => {
    // 1440 units to the inch, so 1440x720 is one inch by half an inch: 96x48 CSS pixels.
    expect(metafileSize(placeableWmf(1440, 720))).toEqual({ width: 96, height: 48 });
    // The units per inch are the header's own, not an assumption.
    expect(metafileSize(placeableWmf(2540, 1270, 2540))).toEqual({ width: 96, height: 48 });
  });

  it("works out the size an EMF asks for, from its frame in hundredths of a millimetre", () => {
    // 2540 hundredths of a mm is one inch.
    expect(metafileSize(emfHeader(2540, 1270))).toEqual({ width: 96, height: 48 });
  });

  it("bounds a picture that declares an absurd frame", () => {
    // A hundred inches across would be a 9600px canvas; the long edge is capped and the aspect
    // ratio survives, which is the part that matters once it is drawn into a cell-sized box.
    const size = metafileSize(placeableWmf(1440 * 100, 1440 * 50))!;
    expect(size.width).toBeLessThanOrEqual(1600);
    expect(size.width / size.height).toBeCloseTo(2, 2);
  });

  it("pulls a metafile out of a data URI, and leaves an ordinary image alone", () => {
    const wmf = placeableWmf(1440, 720);
    const b64 = btoa(String.fromCharCode(...wmf));
    const got = metafileFromDataUri(`data:image/x-wmf;base64,${b64}`)!;
    expect(got.kind).toBe("wmf");
    expect(got.bytes.length).toBe(wmf.length);
    expect(metafileFromDataUri("data:image/png;base64,iVBORw0KGgo=")).toBeNull();
  });
});
