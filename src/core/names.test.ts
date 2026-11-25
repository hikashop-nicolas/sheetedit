import { afterEach, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { setRegionalLocale } from "./regional";
import { readWorkbook } from "./workbook";
import { recalc } from "./recalc";
import { absoluteRangeRef, expandValueNames, formulaUsesName, looksLikeCellRef, nameError, parseRangeRef, refersFromLocal, refersToLocal, renameInFormula } from "./names";

afterEach(() => setRegionalLocale(null));

describe("name validation, as Excel", () => {
  it("accepts letters, underscore or backslash first, then letters, digits, _ and .", () => {
    for (const ok of ["Sales", "_tax", "\\path", "Tax.Rate_2026", "Prix_TTC", "ventes2026", "ABCD1"]) expect(nameError(ok, [])).toBeNull();
  });
  it("refuses spaces, symbols, a digit first and the empty name", () => {
    expect(nameError("", [])).toBe("empty");
    for (const bad of ["my name", "1st", "a-b", "x+y", "_xlnm.Print_Area"]) expect(nameError(bad, [])).toBe("chars");
  });
  it("refuses anything that reads as a cell reference", () => {
    for (const ref of ["A1", "xfd1048576", "R1C1", "R", "C", "rc", "C12"]) expect(nameError(ref, [])).toBe("cellRef");
    expect(looksLikeCellRef("XFE1")).toBe(false);
  });
  it("is unique case-insensitively, except for the name being edited", () => {
    expect(nameError("sales", ["Sales"])).toBe("duplicate");
    expect(nameError("SALES", ["Sales"], "Sales")).toBeNull();
  });
});

describe("references and formulas", () => {
  it("builds an absolute sheet-qualified range and parses it back", () => {
    expect(absoluteRangeRef("Sheet1", { r1: 2, c1: 1, r2: 5, c2: 2 })).toBe("Sheet1!$A$2:$B$5");
    expect(absoluteRangeRef("My data", { r1: 1, c1: 3, r2: 1, c2: 3 })).toBe("'My data'!$C$1");
    expect(parseRangeRef("'It''s'!$B$2:$A$1")).toEqual({ sheet: "It's", rect: { r1: 1, c1: 1, r2: 2, c2: 2 } });
    expect(parseRangeRef("0.5")).toBeNull();
  });
  it("renames a name in formulas, leaving strings, functions and sheet names alone", () => {
    const f = `SUM(sales)+Sales*2&"Sales"+Sales!A1+SALESX(1)+'Sales'!B2`;
    expect(formulaUsesName(f, "Sales")).toBe(true);
    expect(renameInFormula(f, "Sales", "Revenue")).toBe(`SUM(Revenue)+Revenue*2&"Sales"+Sales!A1+SALESX(1)+'Sales'!B2`);
    expect(formulaUsesName("SUM(A1:B2)", "A")).toBe(false);
  });
  it("shows a refers-to the local way and stores it invariant", () => {
    setRegionalLocale("fr");
    expect(refersToLocal("SUM(Feuil1!$A$1:$A$3,0.5)")).toBe("=SOMME(Feuil1!$A$1:$A$3;0,5)");
    expect(refersFromLocal("=SOMME(Feuil1!$A$1:$A$3;0,5)")).toBe("SUM(Feuil1!$A$1:$A$3,0.5)");
    expect(refersFromLocal("0,25")).toBe("0.25");
  });
  it("inlines constant and formula names, nested", () => {
    const names = new Map([["Rate", "0.2"], ["Gross", "Net*(1+Rate)"], ["Net", "S!$A$1"]]);
    expect(expandValueNames("Gross+rate", names)).toBe("(Net*(1+(0.2)))+(0.2)");
  });
});

describe("names in recalc", () => {
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  it("evaluates a constant name, any case, and a sheet-scoped name on its own sheet", () => {
    const body = `<sheetData><row r="1"><c r="A1"><v>10</v></c><c r="B1"><f>A1*rate</f><v>0</v></c><c r="C1"><f>Here*2</f><v>0</v></c></row></sheetData>`;
    const wb = readWorkbook(zipSync({
      "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
      "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Rate">0.5</definedName><definedName name="Here" localSheetId="0">S!$A$1</definedName></definedNames></workbook>`),
      "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
      "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}">${body}</worksheet>`),
    }));
    expect(wb.localNames).toEqual([{ name: "Here", sheet: "S", ref: "S!$A$1" }]);
    recalc(wb);
    expect(wb.sheets[0]!.cells.get("1:2")?.value).toBe("5");
    expect(wb.sheets[0]!.cells.get("1:3")?.value).toBe("20");
  });
});
