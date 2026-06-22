import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { a1ToOdf, odfToA1, readWorkbook, recalc, setCellInput, setXlsxCellStyle, writeWorkbook } from "./index";

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
// number format display layer
// ---------------------------------------------------------------------------

const STYLED_SHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <sheetData>
  <row r="1"><c r="A1" s="1"><v>45000</v></c><c r="B1" s="2"><v>1234.5</v></c><c r="C1" s="2"><f>B1</f><v>1234.5</v></c><c r="D1"><v>1.5</v></c></row>
 </sheetData>
</worksheet>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0.00"/></numFmts>
 <cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/></cellXfs>
</styleSheet>`;

function makeStyledXlsx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(STYLED_SHEET),
    "xl/styles.xml": strToU8(STYLES_XML),
  });
}

describe("number format display layer", () => {
  it("formats xlsx cells via their number format, keeping the raw value editable", () => {
    const wb = readWorkbook(makeStyledXlsx());
    const s = wb.sheets[0]!;
    const a1 = s.cells.get("1:1")!;
    expect(a1.value).toBe("45000"); // raw serial preserved for editing
    expect(a1.display).toBe("3/15/23"); // built-in date format (id 14)
    const b1 = s.cells.get("1:2")!;
    expect(b1.value).toBe("1234.5");
    expect(b1.display).toBe("$1,234.50"); // custom currency code
    const d1 = s.cells.get("1:4")!;
    expect(d1.display).toBeUndefined(); // General format -> no display, raw value shown
  });

  it("reformats a formula cell's display when it recomputes", () => {
    const wb = readWorkbook(makeStyledXlsx());
    const s = wb.sheets[0]!;
    expect(s.cells.get("1:3")!.display).toBe("$1,234.50"); // C1 = B1, currency
    setCellInput(s, 1, 2, "1000.5"); // B1
    recalc(wb);
    expect(s.cells.get("1:3")!.display).toBe("$1,000.50"); // C1 reformatted
    expect(s.cells.get("1:2")!.display).toBe("$1,000.50"); // typed value keeps the format
  });

  it("uses the ODF text:p as the display for formatted .ods cells", () => {
    const content = `<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:spreadsheet><table:table table:name="S">
  <table:table-row><table:table-cell office:value-type="float" office:value="1234.5"><text:p>1,234.50</text:p></table:table-cell></table:table-row>
 </table:table></office:spreadsheet></office:body></office:document-content>`;
    const bytes = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
      "content.xml": strToU8(content),
    } as Record<string, Uint8Array>);
    const cell = readWorkbook(bytes).sheets[0]!.cells.get("1:1")!;
    expect(cell.value).toBe("1234.5"); // raw, editable
    expect(cell.display).toBe("1,234.50"); // producer-formatted text
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

// ---------------------------------------------------------------------------
// xlsx style resolution (display)
// ---------------------------------------------------------------------------

const STYLES = `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <fonts count="2">
  <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
  <font><b/><color rgb="FFFF0000"/></font>
 </fonts>
 <fills count="3">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
 </fills>
 <borders count="2">
  <border><left/><right/><top/><bottom/></border>
  <border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom style="thin"><color rgb="FF000000"/></bottom></border>
 </borders>
 <cellXfs count="3">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"><alignment horizontal="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
 </cellXfs>
</styleSheet>`;

const VSTYLE_SHEET = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <cols><col min="2" max="2" width="20" customWidth="1"/></cols>
 <sheetData>
  <row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="2"><v>5</v></c></row>
 </sheetData>
</worksheet>`;

function makeVisualXlsx(): Uint8Array {
  return zipSync({
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(VSTYLE_SHEET),
    "xl/sharedStrings.xml": strToU8(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Title</t></si></sst>`,
    ),
    "xl/styles.xml": strToU8(STYLES),
  });
}

describe("xlsx cell styles", () => {
  it("resolves font, fill, alignment and column width", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    const a1 = sheet.cells.get("1:1")!.cellStyle!;
    expect(a1.bold).toBe(true);
    expect(a1.color).toBe("#ff0000");
    expect(a1.bg).toBe("#ffff00");
    expect(a1.align).toBe("center");
    // Column B has width 20 (chars) -> ~145px.
    expect(sheet.colWidths?.get(2)).toBe(145);
  });

  it("resolves borders per side", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const b1 = wb.sheets[0]!.cells.get("1:2")!.cellStyle!;
    expect(b1.borders?.left).toBe("#000000");
    expect(b1.borders?.bottom).toBe("#000000");
    expect(b1.borders?.top).toBeUndefined();
  });

  it("writes a style change into the pools and round-trips", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    // A1 starts bold/red/yellow/centre; turn bold off and change the fill.
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:1")!, { bold: false, bg: "#00ff00" });
    const wb2 = readWorkbook(writeWorkbook(wb));
    const a1 = wb2.sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.bold).toBeFalsy();
    expect(a1.bg).toBe("#00ff00");
    expect(a1.color).toBe("#ff0000"); // unchanged attributes are preserved
    expect(a1.align).toBe("center");
  });

  it("adds an all-sides border via a style change", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:1")!, { border: true });
    const a1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.borders?.top).toBe("#000000");
    expect(a1.borders?.right).toBe("#000000");
  });
});
