import { describe, expect, it } from "vitest";
import { extractFormula } from "./formula-assist";

describe("extractFormula", () => {
  it("returns a clean formula unchanged", () => {
    expect(extractFormula("=SUM(A1:A10)")).toBe("=SUM(A1:A10)");
    expect(extractFormula('=IF(A1>0,"y","n")')).toBe('=IF(A1>0,"y","n")');
  });

  it("adds a leading = when the model drops it", () => {
    expect(extractFormula("SUM(A:A)")).toBe("=SUM(A:A)");
  });

  it("pulls the formula out of surrounding prose", () => {
    expect(extractFormula("Sure! Use =AVERAGE(B2:B99) here")).toBe("=AVERAGE(B2:B99) here");
  });

  it("strips code fences", () => {
    expect(extractFormula("```\n=COUNTIF(A:A,\"yes\")\n```")).toBe('=COUNTIF(A:A,"yes")');
  });

  it("returns empty for empty input", () => {
    expect(extractFormula("")).toBe("");
    expect(extractFormula("   ")).toBe("");
  });
});
