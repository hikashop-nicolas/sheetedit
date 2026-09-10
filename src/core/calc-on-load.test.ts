import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { getCell, readWorkbook } from "../index";
import { needsCalcOnLoad, recalc } from "./recalc";

// A workbook may legally carry formulas with no cached results: openpyxl and other
// generators write the <f> and leave the <v> out, and set calcPr fullCalcOnLoad to ask the
// reader to compute. The grid draws cached values, so without a calculation on load every
// one of those cells renders empty. See issue #37.

const sheetXml = (rows: string): string =>
  `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <sheetData>${rows}</sheetData>
</worksheet>`;

function makeXlsx(rows: string, opts: { fullCalcOnLoad?: boolean } = {}): Uint8Array {
  const calcPr = opts.fullCalcOnLoad ? `<calcPr fullCalcOnLoad="1"/>` : "";
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>${calcPr}</workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(sheetXml(rows)),
  });
}

// A1=2, B1=5, C1 is a formula with no cached value, A2 depends on C1.
const UNCACHED = `<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>5</v></c><c r="C1"><f>A1+B1</f></c></row><row r="2"><c r="A2"><f>C1*10</f></c></row>`;
// The same sheet as a normal producer writes it: every result cached.
const CACHED = `<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>5</v></c><c r="C1"><f>A1+B1</f><v>7</v></c></row><row r="2"><c r="A2"><f>C1*10</f><v>70</v></c></row>`;

describe("deciding whether a workbook needs computing on load", () => {
  it("says yes when a formula has no cached result", () => {
    expect(needsCalcOnLoad(readWorkbook(makeXlsx(UNCACHED)))).toBe(true);
  });

  it("says yes when the producer asked for a full recalculation", () => {
    expect(needsCalcOnLoad(readWorkbook(makeXlsx(CACHED, { fullCalcOnLoad: true })))).toBe(true);
  });

  // The point of asking at all: a workbook that carries its results must not be recomputed
  // on the way in, which on a large file is the difference between opening and hanging.
  it("says no for an ordinary workbook whose results are all cached", () => {
    expect(needsCalcOnLoad(readWorkbook(makeXlsx(CACHED)))).toBe(false);
  });

  it("says no for a workbook with no formulas at all", () => {
    const plain = `<row r="1"><c r="A1"><v>2</v></c><c r="B1" t="s"><v>0</v></c></row>`;
    expect(needsCalcOnLoad(readWorkbook(makeXlsx(plain)))).toBe(false);
  });
});

describe("computing a workbook that shipped without results", () => {
  it("fills in the formula cells, following the chain between them", () => {
    const wb = readWorkbook(makeXlsx(UNCACHED, { fullCalcOnLoad: true }));
    const sheet = wb.sheets[0]!;
    expect(getCell(sheet, 1, 3)?.value).toBe(""); // nothing cached to show
    expect(getCell(sheet, 2, 1)?.value).toBe("");

    recalc(wb);

    expect(getCell(sheet, 1, 3)?.value).toBe("7"); // A1+B1
    expect(getCell(sheet, 2, 1)?.value).toBe("70"); // C1*10, so the chain resolved in order
  });

  it("leaves the formulas themselves untouched", () => {
    const wb = readWorkbook(makeXlsx(UNCACHED));
    recalc(wb);
    const sheet = wb.sheets[0]!;
    expect(getCell(sheet, 1, 3)?.formula).toBe("A1+B1");
    expect(getCell(sheet, 2, 1)?.formula).toBe("C1*10");
  });
});
