import { strFromU8 } from "fflate";
import { describe, expect, it } from "vitest";
import { evaluateSection } from "mlang";
import { readWorkbookQueries } from "mlang/qdeff";
import type { MValue } from "mlang";
// Shared with the demo fixture so the test exercises the exact same workbook.
// eslint-disable-next-line
// @ts-ignore - plain .mjs helper, transpiled by vitest without types
import { buildPqXlsx } from "../../../scripts/gen-pq-fixture.mjs";
import { getCell } from "../../core/model";
import { readWorkbook } from "../../core/workbook";
import { applyQueryResult, listWorkbookTables, tableForQuery, tableValue, touchedPositions, workbookHasQueries } from "./tables";

const fixtureBytes = (): Uint8Array => buildPqXlsx() as Uint8Array;

describe("Power Query integration (fixture workbook)", () => {
  it("sniffs the DataMashup payload", () => {
    const wb = readWorkbook(fixtureBytes());
    expect(wb.kind).toBe("xlsx");
    expect(workbookHasQueries(wb.files)).toBe(true);
  });

  it("lists tables with sheet, range and headers", () => {
    const wb = readWorkbook(fixtureBytes());
    const tables = listWorkbookTables(wb);
    const names = tables.map((t) => t.displayName).sort();
    expect(names).toEqual(["Output", "Sales"]);
    const sales = tables.find((t) => t.displayName === "Sales")!;
    expect([sales.r1, sales.c1, sales.r2, sales.c2]).toEqual([2, 2, 5, 4]);
    expect(sales.sheetIndex).toBe(0);
    expect(sales.headerRows).toBe(1);
  });

  it("exposes a table as an mlang value (Excel.CurrentWorkbook shape)", () => {
    const wb = readWorkbook(fixtureBytes());
    const sales = listWorkbookTables(wb).find((t) => t.displayName === "Sales")!;
    const v = tableValue(wb, sales);
    expect(v.columns).toEqual(["Product", "Qty", "Price"]);
    expect(v.rows.length).toBe(3);
    expect(v.rows[0]![0]).toEqual({ kind: "text", value: "Apples" });
    expect(v.rows[2]![1]).toEqual({ kind: "number", value: 20 });
  });

  it("refreshes the Output query end to end and resizes the table part", async () => {
    const wb = readWorkbook(fixtureBytes());
    const q = readWorkbookQueries(wb.files)!;
    expect(q.mashup.sectionM).toContain("shared Output");

    const tables = listWorkbookTables(wb);
    const host = {
      "Excel.CurrentWorkbook": {
        kind: "function",
        name: "Excel.CurrentWorkbook",
        params: [],
        call: (): MValue => ({
          kind: "table",
          columns: ["Name", "Content"],
          rows: tables.map((tb) => [{ kind: "text", value: tb.displayName } as MValue, tableValue(wb, tb) as MValue]),
        }),
      } as MValue,
    };
    const section = await evaluateSection(q.mashup.sectionM, host);
    expect(section.names).toEqual(["Output"]);
    const result = await section.run("Output");
    if (result.kind !== "table") throw new Error("expected a table");
    expect(result.columns).toEqual(["Product", "Quantity", "Total"]);
    expect(result.rows.map((r) => r.map((v) => (v.kind === "text" ? v.value : v.kind === "number" ? v.value : null)))).toEqual([
      ["Cherries", 20, 100],
      ["Apples", 10, 25],
    ]);

    const target = tableForQuery(tables, "Output")!;
    expect(target.displayName).toBe("Output");
    const touched = touchedPositions(target, result);
    expect(touched.length).toBeGreaterThan(0);
    const { rows } = applyQueryResult(wb, target, result);
    expect(rows).toBe(2);

    // Cells written through the model (F2:H4): header + two data rows, stale row replaced.
    const sheet = wb.sheets[0]!;
    expect(getCell(sheet, 2, 6)?.value).toBe("Product");
    expect(getCell(sheet, 3, 6)?.value).toBe("Cherries");
    expect(getCell(sheet, 3, 7)?.value).toBe("20");
    expect(getCell(sheet, 3, 8)?.value).toBe("100");
    expect(getCell(sheet, 4, 6)?.value).toBe("Apples");
    expect(getCell(sheet, 4, 8)?.value).toBe("25");

    // The table part grew from F2:H3 to F2:H4 (autoFilter too).
    const tableXml = strFromU8(wb.files["xl/tables/table2.xml"]!);
    expect(tableXml).toContain('ref="F2:H4"');
    expect(tableXml.match(/ref="F2:H4"/g)!.length).toBe(2);

    // The DataMashup payload was not rewritten by the refresh.
    expect(workbookHasQueries(wb.files)).toBe(true);

    // Save round trip: the written workbook keeps the refreshed cells, the resized table
    // part AND the query definitions, so Excel can still refresh it.
    const { writeWorkbookAsync } = await import("../../core/workbook");
    const saved = await writeWorkbookAsync(wb);
    const wb2 = readWorkbook(saved);
    expect(getCell(wb2.sheets[0]!, 3, 6)?.value).toBe("Cherries");
    expect(getCell(wb2.sheets[0]!, 4, 8)?.value).toBe("25");
    expect(strFromU8(wb2.files["xl/tables/table2.xml"]!)).toContain('ref="F2:H4"');
    expect(workbookHasQueries(wb2.files)).toBe(true);
    expect(readWorkbookQueries(wb2.files)!.mashup.sectionM).toBe(q.mashup.sectionM);
  });

  it("an external-source query raises a typed missing-connector error", async () => {
    const { isMissingConnector, missingConnectorName } = await import("mlang");
    const section = await evaluateSection(`section Section1;\nshared Q = Web.Contents("https://x");`, {});
    try {
      await section.run("Q");
      throw new Error("should have thrown");
    } catch (e) {
      expect(isMissingConnector(e)).toBe(true);
      expect(missingConnectorName(e as never)).toBe("Web.Contents");
    }
  });

  it("a host Web.Contents connector feeds a query (async resolve-by-replay)", async () => {
    const { evaluateSection: evalSec, asyncConnector, toJS } = await import("mlang");
    const host = {
      "Web.Contents": asyncConnector("Web.Contents", async () => ({ kind: "binary" as const, bytes: new TextEncoder().encode("Name,Qty\nApples,10") })),
    };
    const m = `section Section1;\nshared Q = Table.PromoteHeaders(Csv.Document(Web.Contents("https://x/d.csv")));`;
    const section = await evalSec(m, host);
    const out = toJS(await section.run("Q")) as { columns: string[]; rows: unknown[][] };
    expect(out.columns).toEqual(["Name", "Qty"]);
    expect(out.rows).toEqual([["Apples", "10"]]);
  });

  it("a host OData.Feed connector expands the value array (with paging)", async () => {
    const { evaluateSection: evalSec, asyncConnector, tableFromJson, toJS } = await import("mlang");
    const pages: Record<string, { value: unknown[]; "@odata.nextLink"?: string }> = {
      "svc/Products": { value: [{ ID: 1, Name: "Bread" }], "@odata.nextLink": "svc/Products?skip=1" },
      "svc/Products?skip=1": { value: [{ ID: 2, Name: "Milk" }] },
    };
    const host = {
      "OData.Feed": asyncConnector("OData.Feed", async (args) => {
        const records: unknown[] = [];
        let next: string | null = (args[0] as { value: string }).value;
        while (next) {
          const body = pages[next]!;
          records.push(...body.value);
          next = body["@odata.nextLink"] ?? null;
        }
        return tableFromJson(records);
      }),
    };
    const section = await evalSec(`section Section1;\nshared Q = OData.Feed("svc/Products");`, host);
    const out = toJS(await section.run("Q")) as { columns: string[]; rows: unknown[][] };
    expect(out.columns).toEqual(["ID", "Name"]);
    expect(out.rows).toEqual([[1, "Bread"], [2, "Milk"]]);
  });

  it("sniffs a REAL Excel workbook (UTF-16 DataMashup item) and reads its queries", async () => {
    const { readFileSync } = await import("node:fs");
    // cwd-relative: this suite runs under jsdom, where import.meta.url is not file://.
    const real = new Uint8Array(readFileSync("test/fixtures/msft-simple-query.xlsx"));
    const wb = readWorkbook(real);
    expect(wb.kind).toBe("xlsx");
    expect(workbookHasQueries(wb.files)).toBe(true);
    const q = readWorkbookQueries(wb.files);
    expect(q?.mashup.sectionM).toContain("shared Query1");
  });

  it("shrink: a smaller result clears stale rows and shrinks the ref", async () => {
    const wb = readWorkbook(fixtureBytes());
    const tables = listWorkbookTables(wb);
    const target = tableForQuery(tables, "Output")!;
    const one: Extract<MValue, { kind: "table" }> = {
      kind: "table",
      columns: ["Product", "Quantity", "Total"],
      rows: [[{ kind: "text", value: "Only" }, { kind: "number", value: 1 }, { kind: "number", value: 2 }]],
    };
    applyQueryResult(wb, target, one);
    const sheet = wb.sheets[0]!;
    expect(getCell(sheet, 3, 6)?.value).toBe("Only");
    expect(getCell(sheet, 4, 6)?.value ?? "").toBe(""); // no stale second row
    expect(strFromU8(wb.files["xl/tables/table2.xml"]!)).toContain('ref="F2:H3"');
  });
});
