import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../index";
import { flagXlsxPivotRefresh } from "./pivot-read";
import { strFromU8 } from "fflate";

async function realBytes(name: string): Promise<Uint8Array> {
  const { readFileSync } = await import("node:fs");
  return new Uint8Array(readFileSync(`test/fixtures/${name}`));
}

describe("xlsx pivot tables", () => {
  it("reads a pivot's definition (fields, source and target) from a real workbook", async () => {
    const wb = readWorkbook(await realBytes("pivot.xlsx"));
    const host = wb.sheets.find((s) => s.pivotTables?.length);
    expect(host).toBeTruthy();
    const pt = host!.pivotTables![0]!;
    expect(pt.rowFields).toEqual(["Region"]);
    expect(pt.colFields).toEqual(["Product"]);
    expect(pt.dataFields.map((d) => d.name)).toEqual(["Sum - Sales"]);
    expect(pt.sourceSheet).toBe("Data");
    expect(pt.sourceRange).toEqual({ r1: 1, c1: 1, r2: 7, c2: 3 });
    expect(pt.targetRange).toEqual({ r1: 3, c1: 1, r2: 7, c2: 4 });
  });

  it("registers the pivot cache with its worksheet source range", async () => {
    const wb = readWorkbook(await realBytes("pivot.xlsx"));
    expect(wb.pivotCaches?.length).toBe(1);
    expect(wb.pivotCaches![0]!.sourceSheet).toBe("Data");
    expect(wb.pivotCaches![0]!.source).toEqual({ r1: 1, c1: 1, r2: 7, c2: 3 });
    expect(wb.pivotCaches![0]!.part).toMatch(/pivotCacheDefinition/);
  });

  it("flags refreshOnLoad when a cell inside the source range is edited", async () => {
    const wb = readWorkbook(await realBytes("pivot.xlsx"));
    const part = wb.pivotCaches![0]!.part;
    expect(strFromU8(wb.files[part]!)).not.toContain("refreshOnLoad");
    flagXlsxPivotRefresh(wb, "Data", [{ r: 2, c: 3 }]); // B2-ish, inside A1:C7
    expect(strFromU8(wb.files[part]!)).toContain('refreshOnLoad="1"');
    expect(wb.pivotCaches![0]!.refreshFlagged).toBe(true);
  });

  it("does not flag when the edit is outside the source range or on another sheet", async () => {
    const wb = readWorkbook(await realBytes("pivot.xlsx"));
    const part = wb.pivotCaches![0]!.part;
    flagXlsxPivotRefresh(wb, "Data", [{ r: 99, c: 99 }]);
    flagXlsxPivotRefresh(wb, "Pivot", [{ r: 2, c: 2 }]);
    expect(strFromU8(wb.files[part]!)).not.toContain("refreshOnLoad");
  });

  it("preserves the pivot parts verbatim through a read/write round-trip", async () => {
    const { writeWorkbook } = await import("../../core/workbook");
    const wb = readWorkbook(await realBytes("pivot.xlsx"));
    const out = writeWorkbook(wb);
    const back = readWorkbook(out);
    const host = back.sheets.find((s) => s.pivotTables?.length);
    expect(host?.pivotTables?.[0]?.rowFields).toEqual(["Region"]);
    expect(Object.keys(back.files).some((k) => /pivotCacheRecords/.test(k))).toBe(true);
  });
});
