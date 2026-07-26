import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { unzipSync } from "fflate";
import { readCfb } from "./cfb";
import { writeCfb } from "./cfb-write";
import { readVbaProject } from "./vba";
import { setModuleSource, VbaWriteError } from "./vba-write";

// Stage 5 of _plans/VBA_PLAN.md. The reader is the oracle for the writer, as the decompressor is
// for the compressor: a rewritten project must read back the way the original did, minus exactly
// the one module that changed.

const projectOf = (file: string): Uint8Array =>
  unzipSync(new Uint8Array(readFileSync(`src/fixtures/${file}`)))["xl/vbaProject.bin"]!;

const FIXTURES = ["macros-cp950.xlsm", "macros-cp1252.xlsm"];

describe("the CFB writer", () => {
  it("rebuilds both real projects stream for stream", () => {
    for (const file of FIXTURES) {
      const src = readCfb(projectOf(file));
      const rebuilt = readCfb(writeCfb(src));
      const before = src.paths().map((p) => ({ path: p.path, data: [...src.read(p.index)] }));
      const after = rebuilt.paths().map((p) => ({ path: p.path, data: [...rebuilt.read(p.index)] }));
      expect(after.map((s) => s.path).sort(), file).toEqual(before.map((s) => s.path).sort());
      for (const b of before) {
        expect(after.find((a) => a.path === b.path)?.data, `${file} ${b.path}`).toEqual(b.data);
      }
    }
  });

  it("keeps the directory tree the original had", () => {
    const src = readCfb(projectOf("macros-cp1252.xlsm"));
    const rebuilt = readCfb(writeCfb(src));
    // Same entries, same order, same links: only start sector and size are allowed to move.
    expect(rebuilt.entries.map((e) => ({ name: e.name, type: e.type, left: e.left, right: e.right, child: e.child })))
      .toEqual(src.entries.map((e) => ({ name: e.name, type: e.type, left: e.left, right: e.right, child: e.child })));
  });

  it("is deterministic", () => {
    const src = readCfb(projectOf("macros-cp950.xlsm"));
    expect([...writeCfb(src)]).toEqual([...writeCfb(readCfb(projectOf("macros-cp950.xlsm")))]);
  });

  it("carries a replaced stream across at its new size", () => {
    const src = readCfb(projectOf("macros-cp950.xlsm"));
    // A stream large enough to leave the mini stream, so the mini/regular split is exercised.
    const big = new Uint8Array(9000).fill(0x5a);
    const target = src.paths().find((p) => /\/vba\//i.test(p.path))!;
    const rebuilt = readCfb(writeCfb(src, new Map([[target.index, big]])));
    const hit = rebuilt.paths().find((p) => p.path === target.path)!;
    expect([...rebuilt.read(hit.index)]).toEqual([...big]);
    // Everything else still reads.
    for (const p of src.paths()) {
      if (p.path === target.path) continue;
      const other = rebuilt.paths().find((q) => q.path === p.path)!;
      expect([...rebuilt.read(other.index)], p.path).toEqual([...src.read(p.index)]);
    }
  });
});

describe("setModuleSource", () => {
  it("replaces one module's source and leaves the others alone", () => {
    const bin = projectOf("macros-cp1252.xlsm");
    const before = readVbaProject(bin)!;
    const next = "Sub Plus1_Klicken()\r\n    Range(\"A1\").Value = 1\r\nEnd Sub\r\n";
    const out = setModuleSource(bin, "Modul1", next);

    const after = readVbaProject(out)!;
    expect(after.modules.find((m) => m.name === "Modul1")!.source).toBe(next);
    expect(after.codePage).toBe(before.codePage);
    for (const m of before.modules) {
      if (m.name === "Modul1") continue;
      expect(after.modules.find((x) => x.name === m.name)?.source, m.name).toBe(m.source);
    }
  });

  it("survives being written twice, so an edit can be edited", () => {
    const bin = projectOf("macros-cp950.xlsm");
    const once = setModuleSource(bin, "Module1", "Sub A()\r\nEnd Sub\r\n");
    const twice = setModuleSource(once, "Module1", "Sub B()\r\n    Debug.Print 1\r\nEnd Sub\r\n");
    expect(readVbaProject(twice)!.modules.find((m) => m.name === "Module1")!.source)
      .toBe("Sub B()\r\n    Debug.Print 1\r\nEnd Sub\r\n");
  });

  it("drops the performance cache, so the source is the only description of the macro", () => {
    const bin = projectOf("macros-cp950.xlsm");
    const out = setModuleSource(bin, "Module1", "Sub A()\r\nEnd Sub\r\n");
    // MODULEOFFSET is now 0, so nothing sits in front of the compressed source.
    const cfb = readCfb(out);
    const hit = cfb.paths().find((p) => /\/vba\/module1$/i.test(p.path))!;
    expect(cfb.read(hit.index)[0]).toBe(0x01); // the container signature, at byte 0
  });

  it("handles source much longer than the original", () => {
    const bin = projectOf("macros-cp950.xlsm");
    const long = `Sub Big()\r\n${"    Debug.Print \"line\"\r\n".repeat(800)}End Sub\r\n`;
    const out = setModuleSource(bin, "Module1", long);
    expect(readVbaProject(out)!.modules.find((m) => m.name === "Module1")!.source).toBe(long);
  });

  it("handles empty source", () => {
    const bin = projectOf("macros-cp950.xlsm");
    const out = setModuleSource(bin, "Module1", "");
    expect(readVbaProject(out)!.modules.find((m) => m.name === "Module1")!.source).toBe("");
  });

  it("names the module when there is no such module", () => {
    expect(() => setModuleSource(projectOf("macros-cp950.xlsm"), "NoSuchModule", "Sub A()\r\nEnd Sub\r\n"))
      .toThrow(/no module called NoSuchModule/);
  });

  it("refuses a character the workbook's code page cannot hold, rather than mangling it", () => {
    const bin = projectOf("macros-cp1252.xlsm"); // a single-byte code page
    expect(() => setModuleSource(bin, "Modul1", 'Sub A()\r\n    Debug.Print "漢字"\r\nEnd Sub\r\n'))
      .toThrow(VbaWriteError);
    // The original is untouched: setModuleSource returns new bytes, it never edits in place.
    expect(readVbaProject(bin)!.modules.find((m) => m.name === "Modul1")!.source).toMatch(/Plus1_Klicken/);
  });

  it("refuses bytes that are not a macro project at all", () => {
    expect(() => setModuleSource(new Uint8Array([1, 2, 3]), "M", "")).toThrow(/not a compound file/);
  });
});

describe("the verify-before-return rule", () => {
  it("round-trips every module of both fixtures through a rewrite", () => {
    // setModuleSource verifies internally and throws if the result does not read back; running it
    // over every module of both real projects is the broadest check available without Excel.
    for (const file of FIXTURES) {
      const bin = projectOf(file);
      for (const mod of readVbaProject(bin)!.modules) {
        const text = `${mod.source}\r\n' edited\r\n`;
        const out = setModuleSource(bin, mod.name, text);
        expect(readVbaProject(out)!.modules.find((m) => m.name === mod.name)!.source, `${file} ${mod.name}`).toBe(text);
      }
    }
  });
});
