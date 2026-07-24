import { describe, expect, it } from "vitest";
import { capabilitiesFor } from "./capabilities";

describe("workbook capabilities", () => {
  it("xlsx advertises every authoring feature", () => {
    const c = capabilitiesFor("xlsx");
    expect(Object.values(c).every(Boolean)).toBe(true);
  });

  it("ods advertises charts only (for now)", () => {
    const c = capabilitiesFor("ods");
    expect(c.charts).toBe(true);
    expect(c.sparklines).toBe(false);
    expect(c.autofilter).toBe(false);
    expect(c.hyperlinks).toBe(false);
    expect(c.dataValidation).toBe(false);
    expect(c.comments).toBe(false);
    expect(c.conditionalFormat).toBe(false);
  });

  it("csv advertises no advanced features", () => {
    const c = capabilitiesFor("csv");
    expect(Object.values(c).some(Boolean)).toBe(false);
  });
});
