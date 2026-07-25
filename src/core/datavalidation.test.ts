import { describe, expect, it } from "vitest";
import { validateCell } from "./datavalidation";
import type { DataValidation } from "./model";

const dv = (o: Partial<DataValidation>): DataValidation => ({ ranges: [], ...o });

describe("validateCell", () => {
  it("list membership (on the display text)", () => {
    const v = dv({ type: "list", values: ["Yes", "No"] });
    expect(validateCell(v, "Yes", "Yes", ["Yes", "No"])).toBe(true);
    expect(validateCell(v, "Maybe", "Maybe", ["Yes", "No"])).toBe(false);
    expect(validateCell(v, "", "", ["Yes", "No"])).toBe(true); // blank ok
  });

  it("whole number between (rejects decimals and out-of-range)", () => {
    const v = dv({ type: "whole", operator: "between", formula1: "1", formula2: "10" });
    expect(validateCell(v, "5", "5", [])).toBe(true);
    expect(validateCell(v, "5.5", "5.5", [])).toBe(false); // not whole
    expect(validateCell(v, "11", "11", [])).toBe(false); // out of range
    expect(validateCell(v, "abc", "abc", [])).toBe(false); // not a number
  });

  it("decimal with a single-operand operator", () => {
    const v = dv({ type: "decimal", operator: "greaterThanOrEqual", formula1: "0" });
    expect(validateCell(v, "0", "0", [])).toBe(true);
    expect(validateCell(v, "-1", "-1", [])).toBe(false);
    expect(validateCell(v, "2.5", "2.5", [])).toBe(true);
  });

  it("text length compares the display length", () => {
    const v = dv({ type: "textLength", operator: "lessThanOrEqual", formula1: "3" });
    expect(validateCell(v, "ab", "ab", [])).toBe(true);
    expect(validateCell(v, "abcd", "abcd", [])).toBe(false);
  });

  it("date compares the raw serial", () => {
    const v = dv({ type: "date", operator: "greaterThanOrEqual", formula1: "45000" });
    expect(validateCell(v, "45001", "2023-03-16", [])).toBe(true); // raw serial >= 45000
    expect(validateCell(v, "44000", "2020-06-20", [])).toBe(false);
  });

  it("custom rules are not evaluated live (treated valid)", () => {
    expect(validateCell(dv({ type: "custom", formula1: "A1>0" }), "-5", "-5", [])).toBe(true);
  });
});
