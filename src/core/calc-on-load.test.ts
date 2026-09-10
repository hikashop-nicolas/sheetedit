import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { getCell, readWorkbook } from "../index";
import { needsCalcOnLoad, recalc } from "./recalc";

// A workbook may legally carry formulas with no cached results: openpyxl and other generators
// write the <f> and leave the <v> out. The grid draws cached values, so without a calculation on
// load every one of those cells renders empty. See issue #37. The load pass fills those blanks
// and only those: a result the file came with is the producer's, computed by a real engine, and
// must survive opening untouched.

const sheetXml = (rows: string): string =>
  `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <sheetData>${rows}</sheetData>
</worksheet>`;

function makeXlsx(rows: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
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

  // The point of asking at all: a workbook that carries its results must not be recomputed
  // on the way in, which on a large file is the difference between opening and hanging.
  it("says no for an ordinary workbook whose results are all cached", () => {
    expect(needsCalcOnLoad(readWorkbook(makeXlsx(CACHED)))).toBe(false);
  });

  // "" is a result. IF(x="","",...) computes to it, and Excel stores that as an empty <v>, which
  // looks exactly like a blank cell unless the reader keeps the difference.
  it("says no when a formula's cached result IS the empty string", () => {
    const blanked = `<row r="1"><c r="A1"/><c r="B1" t="str"><f>IF(A1="","",A1-1)</f><v/></c></row>`;
    expect(needsCalcOnLoad(readWorkbook(makeXlsx(blanked)))).toBe(false);
  });

  it("says no for a workbook with no formulas at all", () => {
    const plain = `<row r="1"><c r="A1"><v>2</v></c><c r="B1" t="s"><v>0</v></c></row>`;
    expect(needsCalcOnLoad(readWorkbook(makeXlsx(plain)))).toBe(false);
  });
});

describe("computing a workbook that shipped without results", () => {
  it("fills in the formula cells, following the chain between them", () => {
    const wb = readWorkbook(makeXlsx(UNCACHED));
    const sheet = wb.sheets[0]!;
    expect(getCell(sheet, 1, 3)?.value).toBe(""); // nothing cached to show
    expect(getCell(sheet, 2, 1)?.value).toBe("");

    recalc(wb, { keepCached: true });

    expect(getCell(sheet, 1, 3)?.value).toBe("7"); // A1+B1
    expect(getCell(sheet, 2, 1)?.value).toBe("70"); // C1*10, so the chain resolved in order
  });

  it("leaves the formulas themselves untouched", () => {
    const wb = readWorkbook(makeXlsx(UNCACHED));
    recalc(wb, { keepCached: true });
    const sheet = wb.sheets[0]!;
    expect(getCell(sheet, 1, 3)?.formula).toBe("A1+B1");
    expect(getCell(sheet, 2, 1)?.formula).toBe("C1*10");
  });

  // The load pass must not become a rewrite. A result this engine gets wrong (an unsupported
  // function, a reference form it reads differently) would otherwise replace a correct value
  // the producer computed, on a file the user only meant to look at.
  it("keeps a cached result even when recomputing would give a different one", () => {
    // B1's cached 99 disagrees with A1+1; loading must not "correct" it.
    const wb = readWorkbook(makeXlsx(`<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1+1</f><v>99</v></c><c r="C1"><f>B1*2</f></c></row>`));
    recalc(wb, { keepCached: true });
    const sheet = wb.sheets[0]!;
    expect(getCell(sheet, 1, 2)?.value).toBe("99");
    expect(getCell(sheet, 1, 3)?.value).toBe("198"); // the blank is filled, from the cached 99
  });

  it("leaves a deliberately blanked cell blank", () => {
    const wb = readWorkbook(makeXlsx(`<row r="1"><c r="A1"/><c r="B1" t="str"><f>IF(A1="","",A1-1)</f><v/></c></row>`));
    recalc(wb, { keepCached: true });
    expect(getCell(wb.sheets[0]!, 1, 2)?.value).toBe("");
  });

  // An explicit recalculation (an edit, or the user asking) still recomputes everything.
  it("does replace cached results when not in load mode", () => {
    const wb = readWorkbook(makeXlsx(`<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1+1</f><v>99</v></c></row>`));
    recalc(wb);
    expect(getCell(wb.sheets[0]!, 1, 2)?.value).toBe("3");
  });
});

// The load pass is not something the user asked for, so it must not decorate the grid with its
// own complaints about cells it was never going to change.
describe("what the load pass reports", () => {
  it("does not flag a cell that carries a result when the formula will not evaluate", () => {
    // NOTSAFUNCTION does not exist; B1 has a cached value all the same.
    const wb = readWorkbook(makeXlsx(`<row r="1"><c r="B1"><f>NOTSAFUNCTION(1)</f><v>5</v></c><c r="C1"><f>B1+1</f></c></row>`));
    recalc(wb, { keepCached: true });
    const sheet = wb.sheets[0]!;
    expect(getCell(sheet, 1, 2)?.calcFailed).toBeUndefined();
    expect(getCell(sheet, 1, 2)?.value).toBe("5");
    expect(getCell(sheet, 1, 3)?.value).toBe("6"); // the blank still gets filled
  });

  it("still flags one with nothing to show instead", () => {
    const wb = readWorkbook(makeXlsx(`<row r="1"><c r="B1"><f>NOTSAFUNCTION(1)</f></c></row>`));
    recalc(wb, { keepCached: true });
    expect(getCell(wb.sheets[0]!, 1, 2)?.calcFailed).toBe("name");
  });
});
