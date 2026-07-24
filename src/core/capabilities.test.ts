import { describe, expect, it } from "vitest";
import { capabilitiesFor } from "./capabilities";

describe("workbook capabilities", () => {
  it("xlsx advertises every authoring feature", () => {
    const c = capabilitiesFor("xlsx");
    expect(Object.values(c).every(Boolean)).toBe(true);
  });

  it("ods advertises the ODF-native features, but not the Excel-only ones", () => {
    const c = capabilitiesFor("ods");
    expect(c.charts).toBe(true);
    expect(c.hyperlinks).toBe(true);
    expect(c.comments).toBe(true);
    expect(c.dataValidation).toBe(true);
    expect(c.conditionalFormat).toBe(true);
    expect(c.autofilter).toBe(true);
    expect(c.pivots).toBe(true);
    // sparklines author via the LibreOffice calcext extension.
    expect(c.sparklines).toBe(true);
  });

  it("csv advertises no advanced features", () => {
    const c = capabilitiesFor("csv");
    expect(Object.values(c).some(Boolean)).toBe(false);
  });
});
