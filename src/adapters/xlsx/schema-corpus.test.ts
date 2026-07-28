import { describe, expect, it } from "vitest";
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { readWorkbook, writeWorkbook } from "../../index";
import { setCellInput } from "../../core/workbook";

// Not a test of behaviour: this writes the corpus that `npm run check:schema` validates against
// the ECMA-376 schemas. It lives here because only the project's own toolchain can compile the
// library, and it is a test file so it goes through that toolchain rather than needing a build.
//
// Each demo workbook is read, a cell is touched so the writer really re-emits the sheet rather
// than cloning it verbatim, and the result is written beside the original for comparison.

const OUT = join(process.cwd(), ".cache", "schema-corpus");

describe("schema-check corpus", () => {
  it("writes every demo workbook back for validation", () => {
    const demo = join(process.cwd(), "demo");
    if (!existsSync(demo)) return; // a consumer's checkout has no demo fixtures
    mkdirSync(OUT, { recursive: true });
    const names = readdirSync(demo).filter((n) => n.endsWith(".xlsx")).sort();
    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      const wb = readWorkbook(new Uint8Array(readFileSync(join(demo, name))));
      if (wb.sheets[0]) setCellInput(wb.sheets[0], 1, 1, "schema check");
      writeFileSync(join(OUT, name), writeWorkbook(wb));
    }
    expect(readdirSync(OUT).length).toBe(names.length);
  });
});
