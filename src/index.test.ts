import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { a1ToOdf, odfToA1, readWorkbook, recalc, setCellInput, writeWorkbook } from "./index";

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const SHEET1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <sheetData>
  <row r="1"><c r="A1" s="2"><v>2</v></c><c r="B1" t="s"><v>0</v></c><c r="C1"><f>A1*3</f><v>6</v></c></row>
  <row r="2"><c r="A2"><v>5</v></c><c r="B2"><f>SUM(A1:A2)</f><v>7</v></c></row>
 </sheetData>
</worksheet>`;

function makeXlsx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(SHEET1),
    "xl/sharedStrings.xml": strToU8(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>hello</t></si></sst>`,
    ),
    "xl/styles.xml": strToU8("<<STYLES-MARKER>>"),
    "extra.bin": new Uint8Array([9, 8, 7]),
  });
}

const ODS_CONTENT = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:spreadsheet>
  <table:table table:name="Sheet1">
   <table:table-column table:number-columns-repeated="3"/>
   <table:table-row>
    <table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>
    <table:table-cell office:value-type="string" office:string-value="hello"><text:p>hello</text:p></table:table-cell>
    <table:table-cell table:formula="of:=[.A1]*3" office:value-type="float" office:value="6"><text:p>6</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>
    <table:table-cell table:formula="of:=SUM([.A1:.A2])" office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>`;

function makeOds(): Uint8Array {
  const repacked: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
    "content.xml": strToU8(ODS_CONTENT),
    "meta.xml": strToU8("<<META-MARKER>>"),
  };
  return zipSync(repacked as Record<string, Uint8Array>);
}

// ---------------------------------------------------------------------------
// Formula reference translation
// ---------------------------------------------------------------------------

describe("ODF <-> A1 formula refs", () => {
  it("reads ODF formulas to A1", () => {
    expect(odfToA1("of:=[.A1]*3")).toBe("A1*3");
    expect(odfToA1("of:=SUM([.A1:.A2])")).toBe("SUM(A1:A2)");
    expect(odfToA1("of:=[.A1]+[Sheet2.B2]")).toBe("A1+Sheet2!B2");
    expect(odfToA1('of:=IF([.A1]>0;"yes";"no")')).toBe('IF(A1>0,"yes","no")');
  });
  it("writes A1 formulas to ODF", () => {
    expect(a1ToOdf("A1*3")).toBe("of:=[.A1]*3");
    expect(a1ToOdf("SUM(A1:A2)")).toBe("of:=SUM([.A1:.A2])");
    expect(a1ToOdf('IF(A1>0,"yes","no")')).toBe('of:=IF([.A1]>0;"yes";"no")');
    expect(a1ToOdf("LOG10(A1)")).toBe("of:=LOG10([.A1])");
  });
});

// ---------------------------------------------------------------------------
// xlsx
// ---------------------------------------------------------------------------

describe("xlsx", () => {
  it("reads literals, shared strings and formulas", () => {
    const wb = readWorkbook(makeXlsx());
    const s = wb.sheets[0]!;
    expect(s.name).toBe("Sheet1");
    expect(s.cells.get("1:1")).toMatchObject({ value: "2", kind: "n", style: "2" });
    expect(s.cells.get("1:2")).toMatchObject({ value: "hello", kind: "s" });
    expect(s.cells.get("1:3")?.formula).toBe("A1*3");
  });

  it("recalculates formulas in dependency order", () => {
    const wb = readWorkbook(makeXlsx());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 1, "10"); // A1 = 10
    recalc(wb);
    expect(s.cells.get("1:3")?.value).toBe("30"); // C1 = A1*3
    expect(s.cells.get("2:2")?.value).toBe("15"); // B2 = SUM(A1:A2) = 10+5
  });

  it("writes edits in place, preserving styles, other parts and formulas", () => {
    const wb = readWorkbook(makeXlsx());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 1, "10");
    const out = writeWorkbook(wb);
    const files = unzipSync(out);

    // untouched parts survive byte-for-byte
    expect(strFromU8(files["xl/styles.xml"])).toBe("<<STYLES-MARKER>>");
    expect(Array.from(files["extra.bin"])).toEqual([9, 8, 7]);
    const ws = strFromU8(files["xl/worksheets/sheet1.xml"]);
    expect(ws).toContain('s="2"'); // A1 kept its style index
    expect(ws).toContain("<f>A1*3</f>"); // C1 formula preserved

    const wb2 = readWorkbook(out);
    const s2 = wb2.sheets[0]!;
    expect(s2.cells.get("1:1")?.value).toBe("10");
    expect(s2.cells.get("1:3")?.value).toBe("30"); // cached recompute persisted
    expect(s2.cells.get("1:2")?.value).toBe("hello");
  });

  it("adds a string into a previously empty cell as an inline string", () => {
    const wb = readWorkbook(makeXlsx());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 5, "added"); // E1, new cell
    const out = writeWorkbook(wb);
    const ws = strFromU8(unzipSync(out)["xl/worksheets/sheet1.xml"]);
    expect(ws).toContain('r="E1"');
    expect(ws).toContain("added");
    expect(readWorkbook(out).sheets[0]!.cells.get("1:5")?.value).toBe("added");
  });
});

// ---------------------------------------------------------------------------
// ods
// ---------------------------------------------------------------------------

describe("ods", () => {
  it("reads literals and formulas, expanding the table", () => {
    const wb = readWorkbook(makeOds());
    const s = wb.sheets[0]!;
    expect(s.name).toBe("Sheet1");
    expect(s.cells.get("1:1")).toMatchObject({ value: "2", kind: "n" });
    expect(s.cells.get("1:2")?.value).toBe("hello");
    expect(s.cells.get("1:3")?.formula).toBe("A1*3");
    expect(s.cells.get("2:2")?.formula).toBe("SUM(A1:A2)");
  });

  it("recalculates", () => {
    const wb = readWorkbook(makeOds());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 1, "10");
    recalc(wb);
    expect(s.cells.get("1:3")?.value).toBe("30");
    expect(s.cells.get("2:2")?.value).toBe("15");
  });

  it("writes edits, keeping mimetype first/stored and other parts", () => {
    const wb = readWorkbook(makeOds());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 1, "10");
    const out = writeWorkbook(wb);
    const files = unzipSync(out);
    expect(strFromU8(files["meta.xml"])).toBe("<<META-MARKER>>");
    expect(strFromU8(files["mimetype"])).toBe("application/vnd.oasis.opendocument.spreadsheet");

    const wb2 = readWorkbook(out);
    const s2 = wb2.sheets[0]!;
    expect(s2.cells.get("1:1")?.value).toBe("10");
    expect(s2.cells.get("1:3")?.value).toBe("30");
    expect(s2.cells.get("1:2")?.value).toBe("hello");
    expect(s2.cells.get("1:3")?.formula).toBe("A1*3");
  });
});
