import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../index";
import { writeOdsPivotDef } from "./write";
import { computePivot, type PivotSpec } from "../../core/pivot";
import { setCellInput, writeWorkbook } from "../../core/workbook";
import { addSheet } from "../../core/sheet-ops";
import { getCell } from "../../core/model";

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

  it("reconstructs the authoring spec from a file-read data-pilot (so it is editable)", async () => {
    const wb = readWorkbook(await realBytes("pivot.ods"));
    const pt = wb.sheets.find((s) => s.pivotTables?.length)!.pivotTables![0]!;
    expect(pt.authorSpec).toBeTruthy();
    expect(pt.authorSpec!.rows).toEqual([0]); // Region
    expect(pt.authorSpec!.cols).toEqual([1]); // Product
    expect(pt.authorSpec!.values).toEqual([{ field: 2, func: "sum" }]);
  });

  it("attaches the pivot to its target (output) sheet", async () => {
    const wb = readWorkbook(await realBytes("pivot.ods"));
    const host = wb.sheets.find((s) => s.pivotTables?.length);
    expect(host!.name).toBe("Pivot");
  });

  it("authors a new data-pilot with two row fields and materialised output", async () => {
    const wb = readWorkbook(await realBytes("pivot.ods"));
    const data = wb.sheets.find((s) => s.name === "Data")!;
    const spec: PivotSpec = { source: { r1: 1, c1: 1, r2: 7, c2: 3 }, rows: [0, 1], cols: [], values: [{ field: 2, func: "sum" }] };
    const computed = computePivot(data, spec);
    const dest = wb.sheets[addSheet(wb, "Out")]!;
    dest.odsDirty = true;
    for (let r = 0; r < computed.matrix.length; r++)
      for (let c = 0; c < computed.matrix[r]!.length; c++) { const cell = computed.matrix[r]![c]!; if (cell.value !== "") setCellInput(dest, r + 1, c + 1, String(cell.value)); }
    writeOdsPivotDef(wb, dest.name, "Data", spec, computed);
    const back = readWorkbook(writeWorkbook(wb));
    const host = back.sheets.find((s) => s.name === "Out")!;
    expect(host.pivotTables?.[0]?.rowFields).toEqual(["Region", "Product"]);
    expect(host.pivotTables?.[0]?.sourceSheet).toBe("Data");
    // Grand total row is materialised at the bottom of the output.
    expect(getCell(host, computed.height, computed.width)?.value).toBe("350");
  });

  it("preserves the data-pilot block through a read/write round-trip", async () => {
    const { writeWorkbook } = await import("../../core/workbook");
    const wb = readWorkbook(await realBytes("pivot.ods"));
    const out = writeWorkbook(wb);
    const back = readWorkbook(out);
    expect(back.sheets.find((s) => s.pivotTables?.length)?.pivotTables?.[0]?.rowFields).toEqual(["Region"]);
  });
});
