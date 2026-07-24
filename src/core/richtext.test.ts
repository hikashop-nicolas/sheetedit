import { describe, expect, it } from "vitest";
import type { Cell } from "./model";
import { applyRunStyle, cellRuns, normalizeRuns, runsUniform, setRunStyle } from "./richtext";

function cell(value: string, extra?: Partial<Cell>): Cell {
  return { row: 1, col: 1, value, kind: "s", ...extra } as Cell;
}

describe("richtext run authoring", () => {
  it("splits a plain cell into a styled sub-run and back", () => {
    const c = cell("hello world");
    let runs = applyRunStyle(cellRuns(c), 6, 11, { bold: true }); // bold "world"
    expect(runs).toEqual([{ text: "hello " }, { text: "world", bold: true }]);
    // toggling the same range off collapses back to a single plain run
    runs = applyRunStyle(runs, 6, 11, { bold: true });
    expect(normalizeRuns(runs)).toEqual([{ text: "hello world" }]);
  });

  it("applies overlapping ranges and merges adjacent identical runs", () => {
    const c = cell("abcdef");
    let runs = applyRunStyle(cellRuns(c), 0, 3, { italic: true }); // "abc" italic
    runs = applyRunStyle(runs, 2, 5, { bold: true }); // "cde" bold -> "c" both
    expect(runs).toEqual([
      { text: "ab", italic: true },
      { text: "c", italic: true, bold: true },
      { text: "de", bold: true },
      { text: "f" },
    ]);
  });

  it("sets colour and size on a sub-range and clears with an empty value", () => {
    const c = cell("total");
    let runs = applyRunStyle(cellRuns(c), 0, 5, { color: "#ff0000", fontSize: 14 });
    expect(runs).toEqual([{ text: "total", color: "#ff0000", size: 14 }]);
    runs = applyRunStyle(runs, 0, 5, { color: "" });
    expect(runs).toEqual([{ text: "total", size: 14 }]);
  });

  it("toggles based on the whole selection's current state", () => {
    const c = cell("ab");
    let runs = applyRunStyle(cellRuns(c), 0, 1, { bold: true }); // only "a" bold
    // selecting both: not uniformly bold -> sets bold on all
    runs = applyRunStyle(runs, 0, 2, { bold: true });
    expect(runs).toEqual([{ text: "ab", bold: true }]);
    // now uniformly bold -> clears
    runs = applyRunStyle(runs, 0, 2, { bold: true });
    expect(runs).toEqual([{ text: "ab" }]);
  });

  it("setRunStyle sets an absolute value (no toggle) and folds a whole-cell change", () => {
    const runs = [{ text: "a", bold: true }, { text: "b" }];
    // whole-cell "make bold" -> both bold regardless of prior state
    const out = setRunStyle(runs, 0, 2, { bold: true });
    expect(out).toEqual([{ text: "ab", bold: true }]);
  });

  it("runsUniform drops richRuns that equal the base cell style", () => {
    const c = cell("hi", { cellStyle: { bold: true } });
    // both runs bold == the cell's whole-cell bold -> uniform, richRuns can be dropped
    expect(runsUniform([{ text: "hi", bold: true }], c)).toBe(true);
    expect(runsUniform([{ text: "h", bold: true }, { text: "i" }], c)).toBe(false);
  });

  it("cellRuns derives the base run from the cell's whole-cell style", () => {
    const c = cell("x", { cellStyle: { italic: true, color: "#00ff00", fontSize: 12, fontFamily: "Arial" } });
    expect(cellRuns(c)).toEqual([{ text: "x", italic: true, color: "#00ff00", size: 12, font: "Arial" }]);
  });
});
