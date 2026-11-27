import { describe, expect, it } from "vitest";
import { getCell, type Cell, type Sheet, type Workbook } from "./model";
import { recalc } from "./recalc";

// A whole column (A:A) used to be read as all 1,048,576 rows, so every VLOOKUP(x,Data!A:C,..)
// filled down a sheet built a million-row array. Functions that blanks cannot change now read
// only the used rows; the ones that count blanks or rows keep the whole column.

function sheet(name: string, rows: (string | number)[][]): Sheet {
  const s: Sheet = { name, cells: new Map<string, Cell>(), maxRow: 0, maxCol: 0 };
  rows.forEach((row, i) =>
    row.forEach((v, j) => {
      s.cells.set(`${i + 1}:${j + 1}`, { row: i + 1, col: j + 1, value: String(v), kind: typeof v === "number" ? "n" : "s" });
      s.maxRow = Math.max(s.maxRow, i + 1);
      s.maxCol = Math.max(s.maxCol, j + 1);
    }),
  );
  return s;
}

const data = (): Sheet =>
  sheet("Data", [
    ["k1", 10, 1],
    ["k2", 20, 2],
    ["k3", 30, 3],
    ["x4", 40, 4],
    ["k5", 50, 5],
  ]);

function ev(formulas: string[]): string[] {
  const out = sheet("Out", []);
  formulas.forEach((f, i) => out.cells.set(`${i + 1}:1`, { row: i + 1, col: 1, value: "", kind: "blank", formula: f }));
  out.maxRow = formulas.length;
  out.maxCol = 1;
  const wb: Workbook = { kind: "xlsx", files: {}, sheets: [data(), out] };
  recalc(wb);
  return formulas.map((_, i) => getCell(out, i + 1, 1)?.value ?? "");
}

describe("whole-column references", () => {
  it("give the same answers when cut to the used rows", () => {
    expect(
      ev([
        "SUM(Data!B:B)",
        "COUNTA(Data!A:A)",
        "AVERAGE(Data!B:B)",
        "MAX(Data!B:C)",
        'VLOOKUP("k3",Data!A:C,2,FALSE)',
        'MATCH("x4",Data!A:A,0)',
        'INDEX(Data!B:B,MATCH("k5",Data!A:A,0))',
        'SUMIF(Data!A:A,"k*",Data!B:B)',
        'SUMIFS(Data!B:B,Data!A:A,"k*",Data!C:C,">1")',
        "SUMPRODUCT(Data!B:B,Data!C:C)",
        'IFERROR(VLOOKUP("zz",Data!A:B,2,FALSE),-1)',
        "ROUND(AVERAGE(Data!C:C),1)",
      ]),
    ).toEqual(["150", "5", "30", "50", "30", "4", "50", "110", "100", "550", "-1", "3"]);
  });

  it("keep the full column where blanks or rows count", () => {
    expect(ev(["COUNTBLANK(Data!B:B)", "ROWS(Data!B:B)", "INDEX(Data!B:B,100)"])).toEqual([String(1048576 - 5), "1048576", "0"]);
  });

  it("are cheap enough to fill a lookup down a sheet", () => {
    const formulas = Array.from({ length: 300 }, (_, i) => `VLOOKUP("k${(i % 5) + 1}",Data!A:C,3,FALSE)`);
    const t = performance.now();
    const got = ev(formulas);
    expect(performance.now() - t).toBeLessThan(2000);
    expect(got[2]).toBe("3");
  });
});
