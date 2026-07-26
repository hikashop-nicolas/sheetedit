import { describe, expect, it } from "vitest";
import { setCellInput } from "./workbook";
import { recalc } from "./recalc";
import { conditionOf } from "./if-fn";
import type { Sheet, Workbook } from "./model";

// IF with a bare cell reference used to take the true branch whatever the cell held, because the
// parser hands the condition over as a reference object and an object is truthy. These pin the
// behaviour to Excel's, condition coercion included.

function bookWith(cells: Record<string, string>): { wb: Workbook; sheet: Sheet } {
  const sheet: Sheet = { name: "S", cells: new Map(), maxRow: 20, maxCol: 10 };
  const wb: Workbook = { kind: "xlsx", sheets: [sheet], files: {} };
  for (const [ref, raw] of Object.entries(cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
    const col = m[1]!.split("").reduce((a, ch) => a * 26 + (ch.charCodeAt(0) - 64), 0);
    setCellInput(sheet, Number(m[2]), col, raw);
  }
  return { wb, sheet };
}

/** Evaluate one formula in A10 against the given cells. */
function evalIn(cells: Record<string, string>, formula: string): string {
  const { wb, sheet } = bookWith(cells);
  setCellInput(sheet, 10, 1, formula);
  recalc(wb);
  return sheet.cells.get("10:1")!.value;
}

describe("IF condition coercion", () => {
  it("treats a blank as false rather than an error", () => {
    expect(conditionOf(null)).toBe(false);
    expect(conditionOf(undefined)).toBe(false);
  });

  it("passes a boolean through and reads a number as non-zero", () => {
    expect(conditionOf(true)).toBe(true);
    expect(conditionOf(false)).toBe(false);
    expect(conditionOf(0)).toBe(false);
    expect(conditionOf(-1)).toBe(true);
    expect(conditionOf(0.5)).toBe(true);
  });

  it("accepts only text that names a boolean, and errors on anything else", () => {
    expect(conditionOf("TRUE")).toBe(true);
    expect(conditionOf("false")).toBe(false);
    // Excel reports #VALUE! for text that is not a boolean, including the empty string.
    expect(conditionOf("abc")).toEqual({ _error: "#VALUE!" });
    expect(conditionOf("")).toEqual({ _error: "#VALUE!" });
  });

  it("propagates an error condition instead of coercing it", () => {
    expect(conditionOf({ _error: "#DIV/0!" })).toEqual({ _error: "#DIV/0!" });
  });
});

describe("IF over cell references", () => {
  it("takes the false branch for a FALSE cell", () => {
    expect(evalIn({ A1: "FALSE" }, '=IF(A1,"yes","no")')).toBe("no");
  });

  it("takes the true branch for a TRUE cell", () => {
    expect(evalIn({ A1: "TRUE" }, '=IF(A1,"yes","no")')).toBe("yes");
  });

  it("reads a numeric cell as Excel does", () => {
    expect(evalIn({ A1: "0" }, '=IF(A1,"yes","no")')).toBe("no");
    expect(evalIn({ A1: "7" }, '=IF(A1,"yes","no")')).toBe("yes");
  });

  it("reads a blank cell as false", () => {
    expect(evalIn({}, '=IF(A1,"yes","no")')).toBe("no");
  });

  it("still handles a comparison, which always worked", () => {
    expect(evalIn({ A1: "5" }, '=IF(A1>3,"yes","no")')).toBe("yes");
    expect(evalIn({ A1: "1" }, '=IF(A1>3,"yes","no")')).toBe("no");
  });

  it("returns a referenced value from the chosen branch", () => {
    expect(evalIn({ A1: "TRUE", B1: "hit", C1: "miss" }, "=IF(A1,B1,C1)")).toBe("hit");
    expect(evalIn({ A1: "FALSE", B1: "hit", C1: "miss" }, "=IF(A1,B1,C1)")).toBe("miss");
  });

  it("leaves the untaken branch unevaluated, so its error does not surface", () => {
    // The whole point of IF: guarding a division that would otherwise fail.
    expect(evalIn({ A1: "0" }, '=IF(A1=0,"n/a",1/A1)')).toBe("n/a");
    expect(evalIn({ A1: "4" }, '=IF(A1=0,"n/a",1/A1)')).toBe("0.25");
  });

  it("nests", () => {
    expect(evalIn({ A1: "FALSE", A2: "TRUE" }, '=IF(A1,"a",IF(A2,"b","c"))')).toBe("b");
    expect(evalIn({ A1: "FALSE", A2: "FALSE" }, '=IF(A1,"a",IF(A2,"b","c"))')).toBe("c");
  });

  it("errors on a text condition, as Excel reports it", () => {
    expect(evalIn({ A1: "hello" }, '=IF(A1,"yes","no")')).toBe("#VALUE!");
  });

  it("omitting the false branch yields FALSE, not blank", () => {
    expect(evalIn({ A1: "FALSE" }, '=IF(A1,"yes")')).toBe("FALSE");
  });

  it("recalculates when the condition cell flips", () => {
    const { wb, sheet } = bookWith({ A1: "TRUE" });
    setCellInput(sheet, 2, 1, '=IF(A1,"yes","no")');
    recalc(wb);
    expect(sheet.cells.get("2:1")!.value).toBe("yes");
    setCellInput(sheet, 1, 1, "FALSE");
    recalc(wb);
    expect(sheet.cells.get("2:1")!.value).toBe("no");
  });
});
