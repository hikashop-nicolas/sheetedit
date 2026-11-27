import { readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { editModuleSource } from "../../core/vba-macro";
import { readWorkbook } from "../../core/workbook";
import { readCfbStorage, readCfbStreams, writeCfbStreams } from "./cfb";
import { records } from "./biff";
import { writeXls } from "./write";

const fixture = (): Uint8Array => new Uint8Array(readFileSync(join(__dirname, "../../fixtures/compound.xls")));
const vbaBin = (): Uint8Array => unzipSync(new Uint8Array(readFileSync(join(__dirname, "../../fixtures/macros-cp1252.xlsm"))))["xl/vbaProject.bin"]!;

/** A .xls carrying a macro project, built the way Excel lays one out. */
function macroXls(): Uint8Array {
  const wb = readWorkbook(fixture());
  wb.files["xl/vbaProject.bin"] = vbaBin();
  wb.sheets[0]!.codeName = "Sheet1";
  return writeXls(wb);
}

describe("compound file storages", () => {
  it("nests streams in storages and reads one storage back", () => {
    const out = writeCfbStreams([
      { name: "Workbook", data: new Uint8Array(5000).fill(1) },
      { name: "_VBA_PROJECT_CUR/PROJECT", data: new Uint8Array([1, 2, 3]) },
      { name: "_VBA_PROJECT_CUR/VBA/dir", data: new Uint8Array(6000).fill(9) },
      { name: "_VBA_PROJECT_CUR/VBA/Module1", data: new Uint8Array([4]) },
    ]);
    const { streams, storages } = readCfbStreams(out);
    expect(streams.get("Workbook")?.length).toBe(5000);
    expect(storages).toContain("_VBA_PROJECT_CUR");
    const vba = readCfbStorage(out, "_VBA_PROJECT_CUR");
    expect(vba.map((e) => e.name).sort()).toEqual(["PROJECT", "VBA/Module1", "VBA/dir"]);
    expect(vba.find((e) => e.name === "VBA/dir")!.data.every((v) => v === 9)).toBe(true);
  });
});

describe("xls macros", () => {
  it("reads the _VBA_PROJECT_CUR storage as the workbook's macro project", () => {
    const xls = macroXls();
    expect(readCfbStorage(xls, "_VBA_PROJECT_CUR").some((e) => e.name === "VBA/dir")).toBe(true);
    const wb = readWorkbook(xls);
    const fromXlsm = readWorkbook(new Uint8Array(readFileSync(join(__dirname, "../../fixtures/macros-cp1252.xlsm"))));
    expect(wb.vba?.modules.map((m) => [m.name, m.source])).toEqual(fromXlsm.vba?.modules.map((m) => [m.name, m.source]));
    expect(wb.sheets[0]!.codeName).toBe("Sheet1");
  });

  it("keeps the project verbatim on save, with OBPROJ and the code names", () => {
    const xls = macroXls();
    const before = readCfbStorage(xls, "_VBA_PROJECT_CUR");
    const again = writeXls(readWorkbook(xls));
    const after = readCfbStorage(again, "_VBA_PROJECT_CUR");
    expect(after.map((e) => e.name).sort()).toEqual(before.map((e) => e.name).sort());
    for (const e of before) expect(after.find((x) => x.name === e.name)!.data).toEqual(e.data);
    const stream = readCfbStreams(again).streams.get("Workbook")!;
    const types = [...records(stream)].map((r) => r.type);
    expect(types).toContain(0x00d3);
    expect(types.filter((t) => t === 0x01ba).length).toBeGreaterThanOrEqual(2); // workbook + Sheet1
  });

  it("writes an edited module back into the .xls", () => {
    const wb = readWorkbook(macroXls());
    const mod = wb.vba!.modules.find((m) => m.kind === "standard")!;
    const src = mod.source.replace(/\s*$/, "") + "\r\nSub Added()\r\nEnd Sub\r\n";
    expect(editModuleSource(wb, mod.name, src).ok).toBe(true);
    const again = readWorkbook(writeXls(wb));
    expect(again.vba!.modules.find((m) => m.name === mod.name)!.source).toContain("Sub Added()");
  });
});
