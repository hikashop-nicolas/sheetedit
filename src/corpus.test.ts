import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { readWorkbook, writeWorkbook } from "./core/workbook";
import { getCell } from "./core/model";

// Real-file round-trip corpus. The committed fixtures under test/corpus are produced by a
// real engine (LibreOffice), so they carry genuine sharedStrings, styles, calcChain and
// number formats, not the hand-written XML the unit tests use. A dev can drop more files
// (including private/large ones) into the gitignored test/corpus/local and they run too.
const dir = join(__dirname, "../test/corpus");
const localDir = join(dir, "local");
const files = [
  ...(existsSync(dir) ? readdirSync(dir) : []),
  ...(existsSync(localDir) ? readdirSync(localDir).map((f) => join("local", f)) : []),
].filter((f) => /\.(xlsx|ods)$/i.test(f));

// Every cell of every sheet as a stable string, so a round-trip that drops, reorders or
// mutates a value or formula fails loudly.
const sig = (bytes: Uint8Array): string =>
  readWorkbook(bytes)
    .sheets.map(
      (s) =>
        `${s.name}:` +
        [...s.cells.entries()].map(([k, c]) => `${k}=${c.value}${c.formula ? "|" + c.formula : ""}`).sort().join(","),
    )
    .join(";;");

describe.skipIf(!files.length)("real-file round-trip corpus", () => {
  for (const name of files) {
    it(`preserves every cell value and formula through read -> write -> read: ${name}`, () => {
      const bytes = new Uint8Array(readFileSync(join(dir, name)));
      const once = sig(bytes);
      const twice = sig(writeWorkbook(readWorkbook(bytes)));
      expect(twice).toBe(once);
    });
  }
});

// Spot-check that the real-engine files actually parse to the right values and formulas, so
// a regression in real .xlsx/.ods reading (not just round-trip stability) is caught too.
describe("real-file value correctness", () => {
  for (const name of ["lo-formulas.xlsx", "lo-formulas.ods"]) {
    it(`reads LibreOffice formulas and values: ${name}`, () => {
      const path = join(dir, name);
      if (!existsSync(path)) return;
      const s = readWorkbook(new Uint8Array(readFileSync(path))).sheets[0]!;
      expect(getCell(s, 2, 4)?.value).toBe("28.5"); // D2 = B2*C2 = 3 * 9.5
      expect(getCell(s, 2, 4)?.formula).toContain("B2*C2");
      expect(getCell(s, 4, 4)?.value).toBe("44.25"); // D4 = SUM(D2:D3)
      expect(getCell(s, 4, 4)?.formula).toContain("SUM(D2:D3)");
    });
  }
});
