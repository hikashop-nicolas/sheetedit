import { describe, expect, it } from "vitest";
import { editModuleSource } from "./vba-macro";
import type { Workbook } from "./model";

// Sharing macro source, and the line it does not cross.
//
// The source travels; running it never does. A macro is arbitrary code that reaches the
// whole workbook, so running a peer's on everyone's behalf would be strictly worse than
// refreshing their query, which only reaches the network. Whoever wants to run one does so
// knowingly, on code they can read first.
//
// Taking a peer's module goes through the same editor a person uses, which is what these
// pin: it parses first, and it rewrites the macro part rather than only the in-memory copy.

const workbookWithMacros = async (): Promise<Workbook> => {
  const { readFileSync } = await import("node:fs");
  const { unzipSync } = await import("fflate");
  const { readVbaProject } = await import("./vba");
  const files = unzipSync(new Uint8Array(readFileSync("src/fixtures/macros-cp950.xlsm")));
  const wb: Workbook = {
    kind: "xlsx",
    sheets: [{ name: "Sheet1", cells: new Map(), maxRow: 0, maxCol: 0 }],
    files: files as unknown as Record<string, Uint8Array>,
  } as unknown as Workbook;
  wb.vba = readVbaProject(files["xl/vbaProject.bin"]!);
  return wb;
};

describe("macro source, shared", () => {
  it("stores a peer's module, and writes it into the macro part", async () => {
    const wb = await workbookWithMacros();
    const name = wb.vba!.modules[0]!.name;
    const before = wb.files["xl/vbaProject.bin"];

    const res = editModuleSource(wb, name, 'Sub FromAPeer()\n  Range("A1").Value = 1\nEnd Sub\n');

    expect(res.ok).toBe(true);
    expect(wb.vba!.modules.find((m) => m.name === name)!.source).toContain("FromAPeer");
    expect(wb.files["xl/vbaProject.bin"], "the part was rewritten, not just the copy in memory")
      .not.toBe(before);
  });

  // Refused here rather than accepted and left to break the file later.
  it("refuses a module that will not parse, leaving the old source alone", async () => {
    const wb = await workbookWithMacros();
    const name = wb.vba!.modules[0]!.name;
    const original = wb.vba!.modules.find((m) => m.name === name)!.source;

    const res = editModuleSource(wb, name, "Sub Broken(\n");

    expect(res.ok, "not stored").toBe(false);
    expect(wb.vba!.modules.find((m) => m.name === name)!.source, "and what was there is untouched")
      .toBe(original);
  });

  it("ignores a module this workbook does not have", async () => {
    const wb = await workbookWithMacros();
    expect(editModuleSource(wb, "NoSuchModule", "Sub X()\nEnd Sub\n").ok).toBe(false);
  });
});
