import { describe, expect, it } from "vitest";
import { capabilitiesFor } from "./capabilities";

describe("workbook capabilities", () => {
  it("xlsx advertises every authoring feature", () => {
    const c = capabilitiesFor("xlsx");
    expect(Object.values(c).every(Boolean)).toBe(true);
  });

  it("ods advertises charts, hyperlinks and comments, but not the Excel-only features", () => {
    const c = capabilitiesFor("ods");
    expect(c.charts).toBe(true);
    expect(c.hyperlinks).toBe(true);
    expect(c.comments).toBe(true);
    // No Excel-style sparklines or the x14 autofilter model.
    expect(c.sparklines).toBe(false);
    expect(c.autofilter).toBe(false);
  });

  it("csv advertises no advanced features", () => {
    const c = capabilitiesFor("csv");
    expect(Object.values(c).some(Boolean)).toBe(false);
  });
});
