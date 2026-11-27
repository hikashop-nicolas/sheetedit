import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { capabilitiesFor, capabilitiesOf } from "./capabilities";
import { readWorkbook } from "./workbook";

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
    expect(Object.values(c).some((v) => v === true)).toBe(false);
  });
});

describe("legacy .xls capabilities", () => {
  it("hides what BIFF8 cannot store and clamps the grid", () => {
    const c = capabilitiesFor("xls");
    expect(c.tables || c.sparklines || c.pivots || c.charts || c.themes || c.shapes || c.formControls || c.conditionalFormatAdvanced).toBe(false);
    expect(c.comments && c.hyperlinks && c.dataValidation && c.conditionalFormat && c.protection && c.printSetup && c.freezePanes && c.definedNames).toBe(true);
    expect([c.maxRows, c.maxCols]).toEqual([65536, 256]);
    const wb = readWorkbook(new Uint8Array(readFileSync(join(__dirname, "../fixtures/compound.xls"))));
    expect(capabilitiesOf(wb).maxCols).toBe(256);
  });
});
