import { describe, expect, it } from "vitest";
import { buildChart, defaultAnchor } from "./chart-build";

describe("buildChart", () => {
  const rect = { r1: 2, c1: 2, r2: 5, c2: 4 }; // B2:D5

  it("first column labels + first row headers -> categories and named series", () => {
    const m = buildChart("Data", "column", rect, { firstRowHeader: true, firstColLabels: true }, "c1", defaultAnchor(rect));
    expect(m.kind).toBe("column");
    expect(m.categories).toEqual({ ref: "Data!$B$3:$B$5" }); // col B, rows 3-5 (header row skipped)
    expect(m.series).toHaveLength(2); // C and D
    expect(m.series[0].values).toEqual({ ref: "Data!$C$3:$C$5" });
    expect(m.series[0].name).toEqual({ ref: "Data!$C$2" }); // header cell
    expect(m.series[1].values).toEqual({ ref: "Data!$D$3:$D$5" });
    expect(m.dirty).toBe(true);
  });

  it("no headers -> no series names, categories from the first column", () => {
    const m = buildChart("Data", "line", rect, { firstRowHeader: false, firstColLabels: true }, "c2", defaultAnchor(rect));
    expect(m.categories).toEqual({ ref: "Data!$B$2:$B$5" });
    expect(m.series[0].name).toBeUndefined();
    expect(m.series[0].values).toEqual({ ref: "Data!$C$2:$C$5" });
  });

  it("scatter uses the first value column as x and the next as y", () => {
    const m = buildChart("Data", "scatter", rect, { firstRowHeader: true, firstColLabels: false }, "c3", defaultAnchor(rect));
    expect(m.series).toHaveLength(1);
    expect(m.series[0].xValues).toEqual({ ref: "Data!$B$3:$B$5" });
    expect(m.series[0].values).toEqual({ ref: "Data!$C$3:$C$5" });
  });

  it("defaultAnchor sits to the right of the selection", () => {
    const a = defaultAnchor(rect);
    expect(a.fromCol).toBe(6); // c2(4) + 2
    expect(a.toCol).toBe(14);
    expect(a.fromRow).toBe(2);
  });
});
