import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { isCfb, readCfb } from "./cfb";
import { decompressOvba, readVbaProject, subNames } from "./vba";
import { unzipSync } from "fflate";

// Stage 0 of _plans/VBA_PLAN.md: reading a VBA project, no execution.
//
// The compound-file reader is checked against a real .xls produced by LibreOffice, so the fixture
// is a third-party file rather than something written to match the reader. The decompression
// vectors are computed by hand from [MS-OVBA] for the same reason.

const hex = (s: string): Uint8Array => new Uint8Array((s.match(/../g) ?? []).map((b) => parseInt(b, 16)));
const text = (b: Uint8Array): string => new TextDecoder().decode(b);

describe("compound file reader", () => {
  const bytes = new Uint8Array(readFileSync("src/fixtures/compound.xls"));

  it("recognises the signature", () => {
    expect(isCfb(bytes)).toBe(true);
    expect(isCfb(new Uint8Array([1, 2, 3]))).toBe(false);
    expect(isCfb(new Uint8Array(0))).toBe(false);
  });

  it("lists the streams a real .xls contains", () => {
    const paths = readCfb(bytes).paths().map((p) => p.path);
    expect(paths).toContain("/Workbook");
    expect(paths).toContain("/SummaryInformation");
    // A stream inside a storage keeps its parent in the path.
    expect(paths).toContain("/_SX_DB_CUR/0001");
  });

  it("reads a stream's bytes, at the size the directory declares", () => {
    const cfb = readCfb(bytes);
    const wb = cfb.paths().find((p) => p.path === "/Workbook")!;
    const data = cfb.read(wb.index);
    expect(data.length).toBe(cfb.entries[wb.index]!.size);
    // BIFF8 starts with a BOF record (0x0809).
    expect(data[0]! | (data[1]! << 8)).toBe(0x0809);
  });

  it("reads a small stream, which lives in the mini stream rather than the FAT", () => {
    const cfb = readCfb(bytes);
    const ole = cfb.paths().find((p) => /Ole$/.test(p.path))!;
    const data = cfb.read(ole.index);
    // 20 bytes is well under the 4096 cutoff, so this exercises the miniFAT path.
    expect(data.length).toBe(20);
    expect(data.length).toBeLessThan(4096);
  });

  it("finds a stream by path, and returns nothing for one that is absent", () => {
    const cfb = readCfb(bytes);
    expect(cfb.readPath("/Workbook")?.length).toBeGreaterThan(0);
    expect(cfb.readPath("/nope")).toBeUndefined();
  });

  it("refuses bytes that are not a compound file", () => {
    expect(() => readCfb(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).toThrow(/not a compound file/);
  });
});

describe("MS-OVBA decompression", () => {
  // Vectors computed by hand from the spec: header = size | (0b011 << 12) | (compressed << 15),
  // where size is the chunk's byte count less one, and the container starts with 0x01.

  it("decompresses a chunk of literals", () => {
    // 01 | header B002 | flag 00 | 'A' 'B'
    expect(text(decompressOvba(hex("0102b0004142")))).toBe("AB");
  });

  it("decompresses a copy token, including the overlap that encodes a run", () => {
    // After "AB" the difference is 2, so bitCount is 4: token 0x1001 is offset 2, length 4.
    // The copy overlaps its own output, which is how VBA's compression writes repeats.
    expect(text(decompressOvba(hex("0104b00441420110")))).toBe("ABABAB");
  });

  it("decompresses a raw (uncompressed) chunk", () => {
    // Bit 15 clear: the chunk's bytes are the output, with no flag byte or tokens.
    expect(text(decompressOvba(hex("0104304142434445")))).toBe("ABCDE");
  });

  it("returns nothing when the container signature is missing", () => {
    expect(decompressOvba(hex("00ffff")).length).toBe(0);
    expect(decompressOvba(new Uint8Array(0)).length).toBe(0);
  });

  it("stops rather than inventing bytes when a copy points before the output", () => {
    // The first token is a copy, at difference 0, so it can only point before the output.
    const bad = hex("0102b001ffff");
    expect(() => decompressOvba(bad)).not.toThrow();
    expect(decompressOvba(bad).length).toBe(0);
  });

  it("handles a chunk of exactly one literal", () => {
    expect(text(decompressOvba(hex("0101b00041")))).toBe("A");
  });
});

describe("reading a VBA project", () => {
  /** The vbaProject.bin out of one of the .xlsm fixtures (see src/fixtures/README.md). */
  const projectOf = (file: string) =>
    readVbaProject(unzipSync(new Uint8Array(readFileSync(`src/fixtures/${file}`)))["xl/vbaProject.bin"]!);

  it("returns nothing for bytes that are not a compound file", () => {
    expect(readVbaProject(new Uint8Array([1, 2, 3]))).toBeUndefined();
  });

  it("returns nothing for a compound file with no VBA storage", () => {
    // A plain .xls is a compound file, but it carries no /VBA.
    expect(readVbaProject(new Uint8Array(readFileSync("src/fixtures/compound.xls")))).toBeUndefined();
  });

  it("extracts the source of a real macro, whole", () => {
    const proj = projectOf("macros-cp950.xlsm")!;
    const mod = proj.modules.find((m) => m.name === "Module1")!;
    expect(mod.kind).toBe("standard");
    expect(mod.source).toContain('Sub Button1_Click()');
    expect(mod.source).toContain('MsgBox "bingo"');
    expect(mod.source).toContain("End Sub");
    expect(subNames(mod.source)).toEqual(["Button1_Click"]);
  });

  it("finds every module, including the ones bound to a sheet", () => {
    const proj = projectOf("macros-cp950.xlsm")!;
    expect(proj.modules.map((m) => m.name).sort()).toEqual(["Module1", "Sheet1", "Sheet2", "Sheet3", "ThisWorkbook"]);
    // A sheet's module is a class module; only Module1 is procedural.
    expect(proj.modules.filter((m) => m.kind === "standard").map((m) => m.name)).toEqual(["Module1"]);
  });

  it("reads the project's code page and decodes with it", () => {
    expect(projectOf("macros-cp950.xlsm")!.codePage).toBe(950);
    expect(projectOf("macros-cp1252.xlsm")!.codePage).toBe(1252);
  });

  it("handles a project whose names are not English", () => {
    const proj = projectOf("macros-cp1252.xlsm")!;
    expect(proj.modules.map((m) => m.name).sort()).toEqual(["DieseArbeitsmappe", "Modul1"]);
    expect(subNames(proj.modules.find((m) => m.name === "Modul1")!.source)).toEqual(["Plus1_Klicken", "Minus2_Klicken"]);
  });

  it("survives the PROJECTVERSION record, whose 4-byte field is not a size", () => {
    // Reading it as a size desynchronises the walk by two bytes and every later record is garbage,
    // which showed up as a project with no modules at all.
    for (const f of ["macros-cp950.xlsm", "macros-cp1252.xlsm"]) {
      expect(projectOf(f)!.modules.length).toBeGreaterThan(0);
    }
  });
});

describe("finding the macros a module declares", () => {
  it("lists Sub names, whatever their modifiers", () => {
    const src = [
      "Option Explicit",
      "Public Sub Alpha()",
      "End Sub",
      "Private Static Sub Beta()",
      "End Sub",
      "Sub Gamma(ByVal x As Long)",
      "End Sub",
      "Function NotASub() As Long",
      "End Function",
    ].join("\r\n");
    expect(subNames(src)).toEqual(["Alpha", "Beta", "Gamma"]);
  });

  it("ignores the word Sub inside other text", () => {
    expect(subNames("' Sub Hidden()\nDim Substring As String")).toEqual([]);
  });
});
