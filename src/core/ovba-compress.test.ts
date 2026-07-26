import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { unzipSync } from "fflate";
import { compressOvba } from "./ovba-compress";
import { decompressOvba, readVbaProject } from "./vba";
import { readCfb } from "./cfb";

// Stage 5 of _plans/VBA_PLAN.md. The decompressor sheetedit already has, and which has been proven
// against real Excel-written files, is the oracle here: compress then decompress must be identity.

// Node's TextEncoder hands back a Buffer, whose prototype is not Uint8Array's, and toEqual counts
// that as a difference even when every byte matches. Re-wrap so the comparisons mean what they say.
const bytes = (s: string): Uint8Array => new Uint8Array(new TextEncoder().encode(s));
const roundTrip = (b: Uint8Array): Uint8Array => decompressOvba(compressOvba(b));

describe("compressOvba round-trips", () => {
  it("handles the empty input", () => {
    expect(roundTrip(new Uint8Array(0))).toEqual(new Uint8Array(0));
  });

  it("handles input too short to hold a match", () => {
    for (const s of ["a", "ab", "abc", "abcd"]) expect(roundTrip(bytes(s))).toEqual(bytes(s));
  });

  it("handles a long run, which is what copy tokens exist for", () => {
    const b = new Uint8Array(5000).fill(0x41);
    expect(roundTrip(b)).toEqual(b);
    // A run of one repeated byte should compress hard; anything near 1:1 means tokens are unused.
    expect(compressOvba(b).length).toBeLessThan(b.length / 4);
  });

  it("handles repeated text across a chunk boundary", () => {
    const b = bytes("Sub Test()\n    MsgBox \"hello\"\nEnd Sub\n".repeat(400));
    expect(roundTrip(b)).toEqual(b);
    expect(b.length).toBeGreaterThan(4096); // more than one chunk, which is the point
  });

  it("handles incompressible bytes by falling back to a raw chunk", () => {
    // A deterministic pseudo-random sequence: no repeats to find, so tokens would only expand it.
    const b = new Uint8Array(9000);
    let x = 12345;
    for (let i = 0; i < b.length; i++) { x = (x * 1103515245 + 12345) & 0x7fffffff; b[i] = (x >> 16) & 0xff; }
    expect(roundTrip(b)).toEqual(b);
    // Raw chunks cost 2 header bytes per 4096, and nothing more.
    expect(compressOvba(b).length).toBeLessThanOrEqual(b.length + 1 + 3 * 2);
  });

  it("handles exactly one chunk, and one byte either side of it", () => {
    for (const n of [4095, 4096, 4097, 8192]) {
      const b = new Uint8Array(n);
      for (let i = 0; i < n; i++) b[i] = i % 251;
      expect(roundTrip(b), `n=${n}`).toEqual(b);
    }
  });

  it("handles every byte value", () => {
    const b = new Uint8Array(256);
    for (let i = 0; i < 256; i++) b[i] = i;
    expect(roundTrip(b)).toEqual(b);
  });

  it("survives a spread of random-length inputs", () => {
    let x = 7;
    const rnd = (): number => (x = (x * 1103515245 + 12345) & 0x7fffffff);
    for (let trial = 0; trial < 60; trial++) {
      const n = rnd() % 3000;
      const b = new Uint8Array(n);
      // A small alphabet, so matches really do occur and the token path is exercised.
      for (let i = 0; i < n; i++) b[i] = 0x41 + (rnd() % 5);
      expect(roundTrip(b), `trial ${trial}, n=${n}`).toEqual(b);
    }
  });
});

describe("against what Excel actually wrote", () => {
  const streamsOf = (file: string): { path: string; data: Uint8Array }[] => {
    const bin = unzipSync(new Uint8Array(readFileSync(`src/fixtures/${file}`)))["xl/vbaProject.bin"]!;
    const cfb = readCfb(bin);
    return cfb.paths().map((p) => ({ path: p.path, data: cfb.read(p.index) }));
  };

  it("round-trips every compressed stream out of both real projects", () => {
    for (const file of ["macros-cp950.xlsm", "macros-cp1252.xlsm"]) {
      for (const s of streamsOf(file)) {
        if (s.data[0] !== 0x01) continue; // not an MS-OVBA container
        const plain = decompressOvba(s.data);
        if (!plain.length) continue;
        expect(roundTrip(plain), `${file} ${s.path}`).toEqual(plain);
      }
    }
  });

  it("compresses real macro source about as tightly as Excel did", () => {
    const bin = unzipSync(new Uint8Array(readFileSync("src/fixtures/macros-cp1252.xlsm")))["xl/vbaProject.bin"]!;
    const source = readVbaProject(bin)!.modules.find((m) => m.name === "Modul1")!.source;
    const plain = bytes(source);
    const packed = compressOvba(plain);
    expect(decompressOvba(packed)).toEqual(plain);
    // Not a fidelity requirement, just a check that tokens are being found at all.
    expect(packed.length).toBeLessThan(plain.length * 0.75);
  });
});
