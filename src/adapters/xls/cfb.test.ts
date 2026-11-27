import { describe, expect, it } from "vitest";
import { readCfbStreams, writeCfbStreams } from "./cfb";

describe("compound file container", () => {
  it("round-trips small (mini stream) and large (FAT + DIFAT) streams", () => {
    const small = new Uint8Array(100).map((_, i) => i);
    const mid = new Uint8Array(5000).map((_, i) => i * 7);
    const big = new Uint8Array(8 * 1024 * 1024); // > 109 FAT sectors, so DIFAT sectors too
    for (let i = 0; i < big.length; i++) big[i] = (i * 31) & 255;
    const out = writeCfbStreams([{ name: "Workbook", data: big }, { name: "\u0005SummaryInformation", data: small }, { name: "Mid", data: mid }, { name: "Empty", data: new Uint8Array(0) }]);
    const { streams } = readCfbStreams(out);
    const wbk = streams.get("Workbook")!;
    expect(wbk.length === big.length && wbk.every((v, i) => v === big[i])).toBe(true);
    expect(streams.get("\u0005SummaryInformation")).toEqual(small);
    expect(streams.get("Mid")).toEqual(mid);
    expect(streams.get("Empty")?.length).toBe(0);
  });

  it("never hangs on a truncated or looping container", () => {
    const cut = writeCfbStreams([{ name: "Workbook", data: new Uint8Array(10000) }]).slice(0, 300);
    expect(readCfbStreams(cut).streams.get("Workbook")).toBeUndefined();
    // A FAT whose chain points back at itself is read once, not forever.
    const loop = writeCfbStreams([{ name: "Workbook", data: new Uint8Array(10000).fill(1) }]);
    const dv = new DataView(loop.buffer);
    const fat = 512 + dv.getUint32(76, true) * 512;
    for (let i = 0; i < 128; i++) if (dv.getUint32(fat + i * 4, true) === 0xfffffffe) dv.setUint32(fat + i * 4, 2, true);
    expect(() => readCfbStreams(loop)).not.toThrow();
  });
});
