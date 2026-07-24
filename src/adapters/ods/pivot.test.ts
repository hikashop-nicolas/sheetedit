import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../index";

async function realBytes(name: string): Promise<Uint8Array> {
  const { readFileSync } = await import("node:fs");
  return new Uint8Array(readFileSync(`test/fixtures/${name}`));
}

describe("ods data-pilot (pivot) tables", () => {
  it("reads a data-pilot's definition from a real ODS", async () => {
    const wb = readWorkbook(await realBytes("pivot.ods"));
    const host = wb.sheets.find((s) => s.pivotTables?.length);
    expect(host).toBeTruthy();
    const pt = host!.pivotTables![0]!;
    expect(pt.rowFields).toEqual(["Region"]);
    expect(pt.colFields).toEqual(["Product"]);
    expect(pt.dataFields.map((d) => d.name)).toEqual(["Sales"]);
    expect(pt.dataFields[0]!.func).toBe("sum");
    expect(pt.sourceSheet).toBe("Data");
    expect(pt.sourceRange).toEqual({ r1: 1, c1: 1, r2: 7, c2: 3 });
    expect(pt.targetRange).toEqual({ r1: 1, c1: 1, r2: 7, c2: 4 });
  });

  it("attaches the pivot to its target (output) sheet", async () => {
    const wb = readWorkbook(await realBytes("pivot.ods"));
    const host = wb.sheets.find((s) => s.pivotTables?.length);
    expect(host!.name).toBe("Pivot");
  });

  it("preserves the data-pilot block through a read/write round-trip", async () => {
    const { writeWorkbook } = await import("../../core/workbook");
    const wb = readWorkbook(await realBytes("pivot.ods"));
    const out = writeWorkbook(wb);
    const back = readWorkbook(out);
    expect(back.sheets.find((s) => s.pivotTables?.length)?.pivotTables?.[0]?.rowFields).toEqual(["Region"]);
  });
});
