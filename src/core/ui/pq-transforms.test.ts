import { describe, expect, it } from "vitest";
import { evaluate } from "mlang";
import { toJS } from "mlang";
import { TRANSFORMS, quoteName, strLit, fieldRef, nameList } from "./pq-transforms";

const T = (id: string) => TRANSFORMS.find((t) => t.id === id)!;
// Evaluate `let Source = <table> in <buildM output>` so each transform is checked end to end
// against the real mlang engine, not just string-matched.
async function run(prev: string, m: string): Promise<unknown> {
  return toJS(await evaluate(`let ${prev} = #table({"Product", "Qty", "Price"}, {{"Apples", 10, 2}, {"Pears", 4, 3}, {"Cherries", 20, 5}}) in ${m}`));
}
const cols = (v: unknown): string[] => (v as { columns: string[] }).columns;
const rows = (v: unknown): number => (v as { rows: unknown[] }).rows.length;

describe("pq-transforms: M helpers", () => {
  it("quotes and escapes correctly", () => {
    expect(quoteName("Qty")).toBe("Qty");
    expect(quoteName("Unit Price")).toBe('#"Unit Price"');
    expect(strLit('a"b')).toBe('"a""b"');
    expect(fieldRef("Unit Price")).toBe('[#"Unit Price"]');
    expect(nameList(["A", "B"])).toBe('{"A", "B"}');
  });
});

describe("pq-transforms: generated M runs on the engine", () => {
  it("removeColumns / chooseColumns", async () => {
    expect(cols(await run("S", T("removeColumns").buildM("S", { cols: ["Price"] })))).toEqual(["Product", "Qty"]);
    expect(cols(await run("S", T("chooseColumns").buildM("S", { cols: ["Product"] })))).toEqual(["Product"]);
  });
  it("renameColumn", async () => {
    expect(cols(await run("S", T("renameColumn").buildM("S", { column: "Qty", name: "Quantity" })))).toEqual(["Product", "Quantity", "Price"]);
  });
  it("filterRows: numeric and text conditions", async () => {
    expect(rows(await run("S", T("filterRows").buildM("S", { column: "Qty", op: "gt", value: "5" })))).toBe(2);
    expect(rows(await run("S", T("filterRows").buildM("S", { column: "Product", op: "contains", value: "err" })))).toBe(1); // Cherries
    expect(rows(await run("S", T("filterRows").buildM("S", { column: "Product", op: "eq", value: "Pears" })))).toBe(1);
  });
  it("sort descending", async () => {
    const v = await run("S", T("sort").buildM("S", { column: "Qty", dir: "Order.Descending" }));
    expect((v as { rows: unknown[][] }).rows[0][0]).toBe("Cherries");
  });
  it("keep/remove rows", async () => {
    expect(rows(await run("S", T("keepTop").buildM("S", { n: "2" })))).toBe(2);
    expect(rows(await run("S", T("keepBottom").buildM("S", { n: "1" })))).toBe(1);
    expect(rows(await run("S", T("removeTop").buildM("S", { n: "1" })))).toBe(2);
  });
  it("removeDuplicates / reverse", async () => {
    expect(rows(await run("S", T("removeDuplicates").buildM("S", {})))).toBe(3);
    const rev = await run("S", T("reverse").buildM("S", {}));
    expect((rev as { rows: unknown[][] }).rows[0][0]).toBe("Cherries");
  });
  it("changeType to whole number", async () => {
    // Just assert it evaluates and keeps the shape (type tracking is mlang's concern).
    expect(cols(await run("S", T("changeType").buildM("S", { column: "Qty", type: "Int64.Type" })))).toEqual(["Product", "Qty", "Price"]);
  });
  it("replaceValues", async () => {
    const v = await run("S", T("replaceValues").buildM("S", { column: "Product", find: "Pears", replace: "Plums" }));
    expect((v as { rows: unknown[][] }).rows[1][0]).toBe("Plums");
  });
  it("splitColumn adds parts", async () => {
    const m = T("splitColumn").buildM("S", { column: "Product", delimiter: "e", parts: "2" });
    expect(cols(await run("S", m))).toContain("Product.1");
  });
  it("groupBy: count and sum", async () => {
    const cnt = await run("S", T("groupBy").buildM("S", { column: "Product", agg: "count", valueColumn: "" }));
    expect(cols(cnt)).toEqual(["Product", "Count"]);
    const sum = await run("S", T("groupBy").buildM("S", { column: "Product", agg: "sum", valueColumn: "Qty" }));
    expect(cols(sum)).toEqual(["Product", "Sum of Qty"]);
  });
  it("unpivotOthers", async () => {
    const v = await run("S", T("unpivotOthers").buildM("S", { keep: ["Product"] }));
    expect(cols(v)).toEqual(["Product", "Attribute", "Value"]);
    expect(rows(v)).toBe(6); // 3 rows x 2 unpivoted columns
  });
  it("customColumn / indexColumn / transpose / promoteHeaders", async () => {
    expect(cols(await run("S", T("customColumn").buildM("S", { name: "Total", expr: "[Qty] * [Price]" })))).toContain("Total");
    expect(cols(await run("S", T("indexColumn").buildM("S", { start: "1" })))).toContain("Index");
    expect(rows(await run("S", T("transpose").buildM("S", {})))).toBe(3); // 3 columns -> 3 rows
    // promoteHeaders on a headerless table (just check it evaluates)
    expect(await run("S", T("promoteHeaders").buildM("S", {}))).toBeTruthy();
  });
});
