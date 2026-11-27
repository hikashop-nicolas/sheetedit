import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { buildChart } from "../../core/chart-build";
import type { ChartKind, ChartModel } from "../../core/chart-model";
import type { Sheet, SheetShape, Workbook } from "../../core/model";
import { readWorkbook, setCellInput } from "../../core/workbook";
import { xlsSaveWarnings } from ".";
import { xlsCanWrite } from "./drawing";
import { writeXls } from "./write";

const fixture = (): Workbook => readWorkbook(new Uint8Array(readFileSync(join(__dirname, "../../fixtures/compound.xls"))));
const reread = (wb: Workbook): Workbook => readWorkbook(writeXls(wb));
const anchor = (c: number, r: number) => ({ fromCol: c, fromRow: r, fromColOff: 0, fromRowOff: 0, toCol: c + 5, toRow: r + 12, toColOff: 0, toRowOff: 0 });

/** A sheet with a small table at J30:M34 (labels in J, series A..C). */
function withData(): { wb: Workbook; s: Sheet } {
  const wb = fixture();
  const s = wb.sheets[0]!;
  const rows = [["", "A", "B", "C"], ["Jan", "3", "5", "2"], ["Feb", "4", "2", "6"], ["Mar", "6", "3", "4"], ["Apr", "2", "7", "5"]];
  rows.forEach((row, r) => row.forEach((v, c) => setCellInput(s, 30 + r, 10 + c, v)));
  return { wb, s };
}
const chart = (s: Sheet, kind: ChartKind, i = 0): ChartModel =>
  buildChart(s.name, kind, { r1: 30, c1: 10, r2: 34, c2: 13 }, { firstRowHeader: true, firstColLabels: true }, `c${i}`, anchor(2 + i * 6, 40));

describe("xls charts", () => {
  it("writes every chart kind and reads it back", () => {
    const { wb, s } = withData();
    const kinds: ChartKind[] = ["column", "bar", "line", "area", "pie", "doughnut", "scatter", "bubble", "radar", "stock"];
    s.charts = kinds.map((k, i) => ({ ...chart(s, k, i), title: `${k} chart` }));
    const back = reread(wb).sheets[0]!.charts!;
    expect(back.map((c) => c.kind)).toEqual(kinds);
    back.forEach((c, i) => {
      expect(c.title).toBe(`${kinds[i]} chart`);
      expect(c.series.map((x) => x.values.ref)).toEqual(s.charts![i]!.series.map((x) => x.values.ref));
    });
    expect(back[0]!.categories?.ref).toBe(s.charts[0]!.categories?.ref);
    expect(back[6]!.series[0]!.xValues?.ref).toBe(s.charts[6]!.series[0]!.xValues?.ref);
    expect(back[0]!.anchor).toEqual(s.charts[0]!.anchor);
  });

  it("keeps grouping, colours, axes, legend, labels, combos and trendlines", () => {
    const { wb, s } = withData();
    const col = { ...chart(s, "column"), stacked: true, percent: true, gapWidth: 80, overlap: 100, legend: { show: true, pos: "bottom" as const } };
    col.series[0]!.color = "#ff0000";
    col.series[1]!.trendline = { type: "poly", order: 3, dispEq: true };
    col.axes = { x: { title: "Month" }, y: { title: "Value", min: 0, max: 20, numFmt: "0.0" } };
    const pie = { ...chart(s, "doughnut", 1), holeSize: 60, rotation: 90, dataLabels: true };
    pie.series[0]!.pointColors = ["#ff0000", "#00ff00", "#0000ff", "#ffff00"];
    const line = chart(s, "line", 2);
    line.series[1]!.secondaryAxis = true;
    line.series[2]!.type = "column";
    line.series[0]!.smooth = true;
    line.series[0]!.marker = { symbol: "diamond", size: 7 };
    s.charts = [col, pie, line];
    const [c, p, l] = reread(wb).sheets[0]!.charts!;
    expect([c!.stacked, c!.percent, c!.gapWidth, c!.overlap, c!.legend?.pos]).toEqual([true, true, 80, 100, "bottom"]);
    expect(c!.series[0]!.color).toBe("#ff0000");
    expect(c!.series[1]!.trendline).toMatchObject({ type: "poly", order: 3, dispEq: true });
    expect(c!.axes?.x?.title).toBe("Month");
    expect(c!.axes?.y).toMatchObject({ title: "Value", min: 0, max: 20, numFmt: "0.0" });
    expect([p!.holeSize, p!.rotation, p!.dataLabels]).toEqual([60, 90, true]);
    expect(p!.series[0]!.pointColors).toEqual(["#ff0000", "#00ff00", "#0000ff", "#ffff00"]);
    expect(l!.series[1]!.secondaryAxis).toBe(true);
    expect(l!.series[2]!.type).toBe("column");
    expect(l!.series[0]!.smooth).toBe(true);
    expect(l!.series[0]!.marker).toEqual({ symbol: "diamond", size: 7 });
  });

  it("writes a chart sheet as its own substream", () => {
    const { wb, s } = withData();
    const cs: Sheet = { name: "Chart1", cells: new Map(), maxRow: 0, maxCol: 0, chartSheet: true };
    cs.charts = [{ ...chart(s, "pie"), anchor: { fromCol: 1, fromRow: 1, fromColOff: 8, fromRowOff: 8, toCol: 1, toRow: 1, toColOff: 600, toRowOff: 400 } }];
    wb.sheets.push(cs);
    const back = reread(wb);
    const sheet = back.sheets.find((x) => x.name === "Chart1")!;
    expect(sheet.chartSheet).toBe(true);
    expect(sheet.charts?.[0]?.kind).toBe("pie");
    expect(sheet.charts?.[0]?.series[0]?.values.ref).toBe(s.charts?.[0]?.series[0]?.values.ref ?? cs.charts[0]!.series[0]!.values.ref);
  });

  it("keeps a series given only as cached values", () => {
    const { wb, s } = withData();
    const m = chart(s, "column");
    m.series = [{ name: "Lit", values: { cache: [1, 2, 3] } }];
    m.categories = { cache: ["x", "y", "z"] };
    s.charts = [m];
    const back = reread(wb).sheets[0]!.charts![0]!;
    expect(back.series[0]!.values.cache).toEqual([1, 2, 3]);
    expect(back.series[0]!.name).toBe("Lit");
    expect(back.categories?.cache).toEqual(["x", "y", "z"]);
  });
});

describe("xls shapes", () => {
  const shape = (geom: SheetShape["geom"], extra: Partial<SheetShape> = {}): SheetShape => ({ geom, anchor: anchor(2, 2), created: true, dirty: true, ...extra });

  it("writes shapes from the model and reads them back", () => {
    const wb = fixture();
    const s = wb.sheets[0]!;
    s.shapes = [
      shape("roundRect", { fill: "#ff0000", stroke: "#00ff00", strokeWidth: 3, text: "Hi\nthere", textAlign: "center", textValign: "middle", textSize: 14, textColor: "#0000ff" }),
      shape("ellipse", { fill: "#123456", rotation: 30, flipH: true, fillOpacity: 0.5 }),
      shape("line", { stroke: "#000000", dash: "dash", tailEnd: "triangle", headEnd: "oval", flipV: true }),
      shape("rightArrow", { fillGradient: { angle: 90, stops: [{ pos: 0, color: "#ffffff" }, { pos: 1, color: "#000000" }] }, fill: "#ffffff" }),
    ];
    const back = reread(wb).sheets[0]!.shapes!;
    expect(back.map((x) => x.geom)).toEqual(["roundRect", "ellipse", "line", "rightArrow"]);
    expect(back[0]).toMatchObject({ fill: "#ff0000", stroke: "#00ff00", strokeWidth: 3, text: "Hi\nthere", textAlign: "center", textValign: "middle", textSize: 14, textColor: "#0000ff" });
    expect(back[1]).toMatchObject({ fill: "#123456", rotation: 30, flipH: true, fillOpacity: 0.5 });
    expect(back[1]!.stroke).toBeUndefined();
    expect(back[2]).toMatchObject({ stroke: "#000000", dash: "dash", tailEnd: "triangle", headEnd: "oval", flipV: true });
    expect(back[3]!.fillGradient?.stops.map((x) => x.color)).toEqual(["#ffffff", "#000000"]);
    expect(back[0]!.anchor).toEqual(s.shapes[0]!.anchor);
  });

  it("writes an untouched shape back byte for byte, and a moved one with its new anchor", () => {
    const wb = fixture();
    wb.sheets[0]!.shapes = [shape("diamond", { fill: "#ff00ff", text: "x" }), shape("hexagon", { fill: "#00ffff" })];
    const once = reread(wb);
    const [a, b] = once.sheets[0]!.shapes!;
    b!.anchor = { ...b!.anchor, fromRow: 10, toRow: 20 };
    b!.dirty = true;
    const twice = reread(once).sheets[0]!.shapes!;
    expect(twice[0]!.xls!.raw.slice(16)).toEqual(a!.xls!.raw.slice(16)); // past the shape id
    expect(twice[1]!.anchor.fromRow).toBe(10);
    expect(twice[1]!.fill).toBe("#00ffff");
  });

  it("offers only the shapes OfficeArt has, and says when one is saved as a stand-in", () => {
    expect(xlsCanWrite("rect") && xlsCanWrite("rightArrow") && xlsCanWrite("flowDocument")).toBe(true);
    expect(xlsCanWrite("heptagon") || xlsCanWrite("gear6")).toBe(false);
    const wb = fixture();
    wb.sheets[0]!.shapes = [shape("heptagon"), shape("rect")];
    expect(xlsSaveWarnings(wb)).toContainEqual({ key: "xlsShapeStandIn", n: 1 });
    expect(reread(wb).sheets[0]!.shapes!.map((x) => x.geom)).toEqual(["rect", "rect"]);
  });
});
