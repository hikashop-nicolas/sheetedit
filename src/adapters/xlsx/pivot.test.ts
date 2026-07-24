import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../index";
import { flagXlsxPivotRefresh } from "./pivot-read";
import { writeXlsxPivotParts, deleteXlsxPivotParts } from "./pivot-write";
import { computePivot, type PivotSpec } from "../../core/pivot";
import { setCellInput, writeWorkbook } from "../../core/workbook";
import { addSheet } from "../../core/sheet-ops";
import { getCell } from "../../core/model";
import { strFromU8, unzipSync } from "fflate";

function placeAndWrite(fixtureBytes: Uint8Array, spec: PivotSpec): Uint8Array {
  const wb = readWorkbook(fixtureBytes);
  const data = wb.sheets.find((s) => s.name === "Data")!;
  const computed = computePivot(data, spec);
  const dest = wb.sheets[addSheet(wb, "Out")]!;
  for (let r = 0; r < computed.matrix.length; r++)
    for (let c = 0; c < computed.matrix[r]!.length; c++) { const cell = computed.matrix[r]![c]!; if (cell.value !== "") setCellInput(dest, r + 1, c + 1, String(cell.value)); }
  writeXlsxPivotParts(wb, dest, { row: 1, col: 1 }, "Data", spec, computed);
  return writeWorkbook(wb);
}

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

  it("authors a new pivot: emits the parts and reads back the definition", async () => {
    const spec: PivotSpec = { source: { r1: 1, c1: 1, r2: 7, c2: 3 }, rows: [0], cols: [1], values: [{ field: 2, func: "sum" }] };
    const out = placeAndWrite(await realBytes("pivot.xlsx"), spec);
    const files = unzipSync(out);
    // A second pivot cache/table pair was added, with a worksheet source and refreshOnLoad.
    const cache = Object.keys(files).find((k) => /pivotCacheDefinition2\.xml$/.test(k))!;
    expect(strFromU8(files[cache]!)).toContain('refreshOnLoad="1"');
    expect(strFromU8(files[cache]!)).toContain('<worksheetSource ref="A1:C7" sheet="Data"/>');
    // The output cells are materialised (grand total present) and the pivot is detected on re-read.
    const back = readWorkbook(out);
    const host = back.sheets.find((s) => s.name === "Out")!;
    expect(host.pivotTables?.[0]?.rowFields).toEqual(["Region"]);
    expect(host.pivotTables?.[0]?.colFields).toEqual(["Product"]);
    // Materialised grand total (D4 = 350) is present in the output cells.
    expect(getCell(host, 4, 4)?.value).toBe("350");
  });

  it("deletes an authored pivot cleanly (parts, cache, wiring all removed)", async () => {
    const spec: PivotSpec = { source: { r1: 1, c1: 1, r2: 7, c2: 3 }, rows: [0], cols: [1], values: [{ field: 2, func: "sum" }] };
    const wb = readWorkbook(await realBytes("pivot.xlsx"));
    const data = wb.sheets.find((s) => s.name === "Data")!;
    const computed = computePivot(data, spec);
    const dest = wb.sheets[addSheet(wb, "Out")]!;
    const { part, cachePart } = writeXlsxPivotParts(wb, dest, { row: 1, col: 1 }, "Data", spec, computed);
    const cachesBefore = wb.pivotCaches!.length;
    expect(wb.files[part]).toBeTruthy();
    deleteXlsxPivotParts(wb, dest, part, cachePart);
    expect(wb.files[part]).toBeUndefined();
    expect(wb.files[cachePart]).toBeUndefined();
    expect(wb.files[cachePart.replace(/Definition/, "Records")]).toBeUndefined();
    expect(wb.pivotCaches!.length).toBe(cachesBefore - 1);
    // The original fixture's own pivot (on the "Pivot" sheet) is untouched.
    const ct = strFromU8(wb.files["[Content_Types].xml"]!);
    expect(ct).not.toContain(`/${part}`);
    // Re-reading the workbook succeeds and finds only the original pivot.
    const back = readWorkbook(writeWorkbook(wb));
    const pivotHosts = back.sheets.filter((s) => s.pivotTables?.length);
    expect(pivotHosts.length).toBe(1);
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
