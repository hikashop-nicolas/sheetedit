import { describe, expect, it } from "vitest";
import { chartConfig } from "./chart-overlay";
import { buildChart, defaultAnchor } from "../chart-build";
import type { ChartKind, Workbook } from "../model";

// A minimal workbook with the data ranges the built charts reference, so chartConfig resolves
// real values. Guards that EVERY chart kind produces a valid Chart.js config (a real regression:
// pie/doughnut once threw here because their series values weren't resolved).
function wb(): Workbook {
  const cell = (v: string) => ({ row: 0, col: 0, value: v, kind: "n" as const });
  void cell;
  const sheet = {
    name: "Sheet1",
    cells: new Map<string, { row: number; col: number; value: string; kind: "n" | "inlineStr" }>(),
    maxRow: 4,
    maxCol: 3,
  };
  const set = (r: number, c: number, v: string, k: "n" | "inlineStr") => sheet.cells.set(`${r}:${c}`, { row: r, col: c, value: v, kind: k });
  set(1, 1, "Product", "inlineStr"); set(1, 2, "Qty", "inlineStr"); set(1, 3, "Price", "inlineStr");
  set(2, 1, "Apples", "inlineStr"); set(2, 2, "10", "n"); set(2, 3, "2.5", "n");
  set(3, 1, "Pears", "inlineStr"); set(3, 2, "4", "n"); set(3, 3, "3", "n");
  set(4, 1, "Cherries", "inlineStr"); set(4, 2, "20", "n"); set(4, 3, "5", "n");
  return { kind: "xlsx", sheets: [sheet as unknown as Workbook["sheets"][number]], files: {} };
}

const KINDS: ChartKind[] = ["column", "bar", "line", "area", "pie", "doughnut", "scatter", "bubble", "radar"];

describe("chartConfig covers every chart kind", () => {
  const rect = { r1: 1, c1: 1, r2: 4, c2: 3 };
  for (const kind of KINDS) {
    it(`${kind} builds a config with data`, () => {
      const model = buildChart("Sheet1", kind, rect, { firstRowHeader: true, firstColLabels: true }, "c", defaultAnchor(rect));
      const cfg = chartConfig(model, wb()) as { type: string; data: { datasets: { data: unknown[] }[] } };
      expect(cfg.type).toBeTruthy();
      expect(cfg.data.datasets.length).toBeGreaterThan(0);
      // Each dataset has resolved, non-empty data.
      for (const ds of cfg.data.datasets) expect((ds.data as unknown[]).length).toBeGreaterThan(0);
    });
  }

  it("pie / doughnut colour every slice", () => {
    const model = buildChart("Sheet1", "pie", rect, { firstRowHeader: true, firstColLabels: true }, "p", defaultAnchor(rect));
    const cfg = chartConfig(model, wb()) as { data: { datasets: { backgroundColor: unknown }[] } };
    expect(Array.isArray(cfg.data.datasets[0].backgroundColor)).toBe(true);
  });
});
