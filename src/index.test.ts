import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  a1ToOdf,
  odfToA1,
  readWorkbook,
  recalc,
  setCellInput,
  shiftFormula,
  setOdsCellStyle,
  setOdsColWidth,
  setOdsMerge,
  setOdsRowHeight,
  setXlsxCellStyle,
  setXlsxColWidth,
  setXlsxMerge,
  setXlsxRowHeight,
  writeWorkbook,
  writeWorkbookAsync,
} from "./index";

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

  it("readWorkbook from a pre-inflated map matches reading from raw bytes", () => {
    const xlsx = makeXlsx();
    const a = readWorkbook(xlsx);
    const b = readWorkbook(xlsx, {}, unzipSync(xlsx));
    const cells = (wb: ReturnType<typeof readWorkbook>) => [...wb.sheets[0]!.cells.entries()].map(([k, c]) => `${k}=${c.value}`).sort();
    expect(cells(b)).toEqual(cells(a));
    expect(b.sheets.map((s) => s.name)).toEqual(a.sheets.map((s) => s.name));
  });

  it("writeWorkbookAsync (off-thread zip) matches the synchronous writer", async () => {
    const wb1 = readWorkbook(makeXlsx());
    setCellInput(wb1.sheets[0]!, 1, 1, "10");
    const wb2 = readWorkbook(makeXlsx());
    setCellInput(wb2.sheets[0]!, 1, 1, "10");
    const [sync, asyncOut] = [writeWorkbook(wb1), await writeWorkbookAsync(wb2)];
    const a = unzipSync(sync);
    const b = unzipSync(asyncOut);
    expect(Object.keys(b).sort()).toEqual(Object.keys(a).sort());
    for (const k of Object.keys(a)) expect(Array.from(b[k]!)).toEqual(Array.from(a[k]!));
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
 <fonts count="3">
  <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
  <font><b/><color rgb="FFFF0000"/></font>
  <font><u/><strike/><sz val="16"/><name val="Georgia"/></font>
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
 <cellXfs count="4">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"><alignment horizontal="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
 </cellXfs>
</styleSheet>`;

const VSTYLE_SHEET = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <cols><col min="2" max="2" width="20" customWidth="1"/></cols>
 <sheetData>
  <row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="2"><v>5</v></c><c r="C1" s="3"><v>7</v></c></row>
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

  it("sets a single border side via borderSides and round-trips", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:1")!, { borderSides: { left: true } });
    const a1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.borders?.left).toBe("#000000");
    expect(a1.borders?.top).toBeUndefined();
    expect(a1.borders?.right).toBeUndefined();
  });

  it("resolves underline, strike, font size/name, valign and wrap", () => {
    const c1 = readWorkbook(makeVisualXlsx()).sheets[0]!.cells.get("1:3")!.cellStyle!;
    expect(c1.underline).toBe(true);
    expect(c1.strike).toBe(true);
    expect(c1.fontSize).toBe(16);
    expect(c1.fontFamily).toBe("Georgia");
    expect(c1.valign).toBe("top");
    expect(c1.wrap).toBe(true);
  });

  it("suppresses font size/name equal to the workbook default", () => {
    // A1 uses font 1, which has no explicit sz/name (inherits Calibri 11 visuals).
    const a1 = readWorkbook(makeVisualXlsx()).sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.fontSize).toBeUndefined();
    expect(a1.fontFamily).toBeUndefined();
  });

  it("writes underline/strike/size/family/valign/wrap and round-trips", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:1")!, {
      underline: true,
      strike: true,
      fontSize: 20,
      fontFamily: "Georgia",
      valign: "middle",
      wrap: true,
    });
    const a1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.underline).toBe(true);
    expect(a1.strike).toBe(true);
    expect(a1.fontSize).toBe(20);
    expect(a1.fontFamily).toBe("Georgia");
    expect(a1.valign).toBe("middle");
    expect(a1.wrap).toBe(true);
    expect(a1.bold).toBe(true); // pre-existing font traits survive
    expect(a1.align).toBe("center");
  });

  it("turns underline and strike back off", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:3")!, { underline: false, strike: false });
    const c1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:3")!.cellStyle!;
    expect(c1.underline).toBeFalsy();
    expect(c1.strike).toBeFalsy();
    expect(c1.fontSize).toBe(16); // untouched traits stay
  });

  it("mints styles.xml when the workbook has none", () => {
    // e.g. a workbook from the CSV converter: no xl/styles.xml at all.
    const bare = zipSync({
      "[Content_Types].xml": strToU8(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>',
      ),
      "_rels/.rels": strToU8("<Relationships/>"),
      "xl/workbook.xml": strToU8(
        `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      ),
      "xl/_rels/workbook.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      ),
      "xl/worksheets/sheet1.xml": strToU8(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
      ),
    });
    const wb = readWorkbook(bare);
    const sheet = wb.sheets[0]!;
    setXlsxCellStyle(wb, sheet, sheet.cells.get("1:1")!, { bold: true, underline: true });
    const a1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.bold).toBe(true);
    expect(a1.underline).toBe(true);
  });

  it("writes a column width and round-trips", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    setXlsxColWidth(sheet, 3, 215); // ~30 chars
    const out = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect(out.colWidths?.get(3)).toBe(215);
    // The pre-existing width on column B is preserved.
    expect(out.colWidths?.get(2)).toBe(145);
  });

  it("writes a row height and round-trips", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    setXlsxRowHeight(sheet, 1, 40);
    const out = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect(out.rowHeights?.get(1)).toBe(40);
  });

  it("adds and removes a merged range, round-tripping", () => {
    const wb = readWorkbook(makeVisualXlsx());
    const sheet = wb.sheets[0]!;
    setXlsxMerge(sheet, 2, 1, 2, 3, true); // merge A2:C2
    const merged = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect(merged.merges).toContainEqual({ r1: 2, c1: 1, r2: 2, c2: 3 });
    // Unmerge it again.
    setXlsxMerge(sheet, 2, 1, 2, 3, false);
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect((back.merges ?? []).some((m) => m.r1 === 2 && m.c1 === 1 && m.r2 === 2 && m.c2 === 3)).toBe(false);
  });
});

const ODS_STYLED = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
 <office:automatic-styles>
  <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style>
  <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="1cm"/></style:style>
  <style:style style:name="ce1" style:family="table-cell">
   <style:table-cell-properties fo:background-color="#4472c4" fo:border="0.5pt solid #000000" fo:wrap-option="wrap" style:vertical-align="middle"/>
   <style:text-properties fo:color="#ffffff" fo:font-weight="bold" style:text-underline-style="solid" style:text-line-through-style="solid" fo:font-size="14pt" fo:font-family="Georgia"/>
   <style:paragraph-properties fo:text-align="center"/>
  </style:style>
 </office:automatic-styles>
 <office:body><office:spreadsheet>
  <table:table table:name="Sheet1">
   <table:table-column table:style-name="co1"/>
   <table:table-column table:number-columns-repeated="2"/>
   <table:table-row table:style-name="ro1">
    <table:table-cell table:style-name="ce1" table:number-columns-spanned="2" office:value-type="string" office:string-value="Title"><text:p>Title</text:p></table:table-cell>
    <table:covered-table-cell/>
    <table:table-cell office:value-type="float" office:value="9"><text:p>9</text:p></table:table-cell>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>`;

function makeStyledOds(): Uint8Array {
  const repacked: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
    "content.xml": strToU8(ODS_STYLED),
  };
  return zipSync(repacked as Record<string, Uint8Array>);
}

describe("ods cell styles", () => {
  it("resolves cell style, column width, row height and merges", () => {
    const sheet = readWorkbook(makeStyledOds()).sheets[0]!;
    const a1 = sheet.cells.get("1:1")!.cellStyle!;
    expect(a1.bold).toBe(true);
    expect(a1.color).toBe("#ffffff");
    expect(a1.bg).toBe("#4472c4");
    expect(a1.align).toBe("center");
    expect(a1.borders?.top).toBe("#000000");
    // 3cm ~ 113px.
    expect(sheet.colWidths?.get(1)).toBe(113);
    // 1cm ~ 38px.
    expect(sheet.rowHeights?.get(1)).toBe(38);
    expect(sheet.merges).toContainEqual({ r1: 1, c1: 1, r2: 1, c2: 2 });
  });

  it("writes a cell style and round-trips", () => {
    const wb = readWorkbook(makeStyledOds());
    const sheet = wb.sheets[0]!;
    setOdsCellStyle(wb, sheet, sheet.cells.get("1:3")!, { bold: true, bg: "#ff0000", align: "right" });
    const c1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:3")!.cellStyle!;
    expect(c1.bold).toBe(true);
    expect(c1.bg).toBe("#ff0000");
    expect(c1.align).toBe("right");
  });

  it("writes a per-side border and round-trips", () => {
    const wb = readWorkbook(makeStyledOds());
    const sheet = wb.sheets[0]!;
    setOdsCellStyle(wb, sheet, sheet.cells.get("1:3")!, { borderSides: { bottom: true } });
    const c1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:3")!.cellStyle!;
    expect(c1.borders?.bottom).toBe("#000000");
    expect(c1.borders?.top).toBeUndefined();
  });

  it("resolves underline, strike, font size/family, valign and wrap", () => {
    const a1 = readWorkbook(makeStyledOds()).sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.underline).toBe(true);
    expect(a1.strike).toBe(true);
    expect(a1.fontSize).toBe(14);
    expect(a1.fontFamily).toBe("Georgia");
    expect(a1.valign).toBe("middle");
    expect(a1.wrap).toBe(true);
  });

  it("writes underline/strike/size/family/valign/wrap and round-trips", () => {
    const wb = readWorkbook(makeStyledOds());
    const sheet = wb.sheets[0]!;
    setOdsCellStyle(wb, sheet, sheet.cells.get("1:3")!, {
      underline: true,
      strike: true,
      fontSize: 18,
      fontFamily: "Verdana",
      valign: "bottom",
      wrap: true,
    });
    const c1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:3")!.cellStyle!;
    expect(c1.underline).toBe(true);
    expect(c1.strike).toBe(true);
    expect(c1.fontSize).toBe(18);
    expect(c1.fontFamily).toBe("Verdana");
    expect(c1.valign).toBe("bottom");
    expect(c1.wrap).toBe(true);
  });

  it("turns underline and strike back off, keeping the rest", () => {
    const wb = readWorkbook(makeStyledOds());
    const sheet = wb.sheets[0]!;
    setOdsCellStyle(wb, sheet, sheet.cells.get("1:1")!, { underline: false, strike: false });
    const a1 = readWorkbook(writeWorkbook(wb)).sheets[0]!.cells.get("1:1")!.cellStyle!;
    expect(a1.underline).toBeFalsy();
    expect(a1.strike).toBeFalsy();
    expect(a1.fontSize).toBe(14);
    expect(a1.bold).toBe(true);
    expect(a1.wrap).toBe(true);
  });

  it("writes a column width and a row height, round-tripping", () => {
    const wb = readWorkbook(makeStyledOds());
    const sheet = wb.sheets[0]!;
    setOdsColWidth(wb, sheet, 3, 200);
    setOdsRowHeight(wb, sheet, 1, 50);
    const out = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect(out.colWidths?.get(3)).toBe(200);
    expect(out.rowHeights?.get(1)).toBe(50);
    // Column 1's original 3cm width is preserved.
    expect(out.colWidths?.get(1)).toBe(113);
  });

  it("adds and removes a merge, round-tripping", () => {
    const wb = readWorkbook(makeOds());
    const sheet = wb.sheets[0]!;
    setOdsMerge(sheet, 1, 1, 1, 3, true);
    const merged = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect(merged.merges).toContainEqual({ r1: 1, c1: 1, r2: 1, c2: 3 });
    expect(merged.cells.get("1:1")!.value).toBe("2"); // top-left value kept
    setOdsMerge(sheet, 1, 1, 1, 3, false);
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect((back.merges ?? []).length).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// shared formulas (xlsx <f t="shared">)
// ---------------------------------------------------------------------------

const SHARED_SHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <sheetData>
  <row r="1"><c r="A1"><v>2</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*10</f><v>20</v></c></row>
  <row r="2"><c r="A2"><v>3</v></c><c r="B2"><f t="shared" si="0"/><v>30</v></c></row>
  <row r="3"><c r="A3"><v>4</v></c><c r="B3"><f t="shared" si="0"/><v>40</v></c></row>
 </sheetData>
</worksheet>`;

function makeSharedXlsx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8("<Relationships/>"),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(SHARED_SHEET),
    "xl/calcChain.xml": strToU8(`<calcChain/>`),
  });
}

describe("shiftFormula", () => {
  it("shifts relative refs and keeps absolute parts", () => {
    expect(shiftFormula("A1*10", 1, 0)).toBe("A2*10");
    expect(shiftFormula("$A$1+B2", 2, 3)).toBe("$A$1+E4");
    expect(shiftFormula("A$1+$B2", 1, 1)).toBe("B$1+$B3");
    expect(shiftFormula("SUM(A1:B2)", 1, 1)).toBe("SUM(B2:C3)");
  });
  it("leaves strings, function names and defined names alone", () => {
    expect(shiftFormula('IF(A1>0,"A1 ok","no")', 1, 0)).toBe('IF(A2>0,"A1 ok","no")');
    expect(shiftFormula("LOG10(A1)", 1, 0)).toBe("LOG10(A2)");
    expect(shiftFormula("myname1+A1", 1, 0)).toBe("myname1+A2");
    expect(shiftFormula("ZZZ9999999+A1", 1, 0)).toBe("ZZZ9999999+A2");
  });
  it("shifts sheet-qualified refs", () => {
    expect(shiftFormula("Sheet2!A1+1", 1, 1)).toBe("Sheet2!B2+1");
    expect(shiftFormula("'My Sheet'!A1", 0, 1)).toBe("'My Sheet'!B1");
  });
  it("marks out-of-range results as #REF!", () => {
    expect(shiftFormula("A1", -1, 0)).toBe("#REF!");
  });
});

describe("xlsx shared formulas", () => {
  it("resolves child formulas from the master and recalcs them", () => {
    const wb = readWorkbook(makeSharedXlsx());
    const s = wb.sheets[0]!;
    expect(s.cells.get("2:2")?.formula).toBe("A2*10");
    expect(s.cells.get("3:2")?.formula).toBe("A3*10");
    setCellInput(s, 2, 1, "7"); // A2 = 7
    recalc(wb);
    expect(s.cells.get("2:2")?.value).toBe("70"); // child recalculated
  });

  it("keeps shared <f> attributes when only cached values change", () => {
    const wb = readWorkbook(makeSharedXlsx());
    setCellInput(wb.sheets[0]!, 2, 1, "7"); // value edit, not a formula edit
    const ws = strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]);
    expect(ws).toContain('t="shared" ref="B1:B3" si="0"');
    expect(ws).toContain("<v>70</v>");
  });

  it("de-shares the whole group when a member formula is edited", () => {
    const wb = readWorkbook(makeSharedXlsx());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 2, "=A1*100"); // edit the master's formula
    const out = writeWorkbook(wb);
    const ws = strFromU8(unzipSync(out)["xl/worksheets/sheet1.xml"]);
    expect(ws).not.toContain('t="shared"');
    expect(ws).not.toContain("si=");
    // children keep working, as plain translated formulas
    const wb2 = readWorkbook(out);
    expect(wb2.sheets[0]!.cells.get("2:2")?.formula).toBe("A2*10");
    expect(wb2.sheets[0]!.cells.get("1:2")?.formula).toBe("A1*100");
  });

  it("drops calcChain.xml and sets fullCalcOnLoad when a formula changes", () => {
    const wb = readWorkbook(makeSharedXlsx());
    setCellInput(wb.sheets[0]!, 1, 4, "=A1+A2"); // new formula
    const files = unzipSync(writeWorkbook(wb));
    expect(files["xl/calcChain.xml"]).toBeUndefined();
    expect(strFromU8(files["[Content_Types].xml"])).not.toContain("calcChain");
    expect(strFromU8(files["xl/_rels/workbook.xml.rels"])).not.toContain("calcChain");
    expect(strFromU8(files["xl/workbook.xml"])).toContain('fullCalcOnLoad="1"');
  });

  it("keeps calcChain.xml when no formula changed", () => {
    const wb = readWorkbook(makeSharedXlsx());
    setCellInput(wb.sheets[0]!, 2, 1, "7"); // value edit only
    const files = unzipSync(writeWorkbook(wb));
    expect(files["xl/calcChain.xml"]).toBeDefined();
    expect(strFromU8(files["xl/workbook.xml"])).not.toContain("fullCalcOnLoad");
  });
});

// ---------------------------------------------------------------------------
// ods preservation (untouched sheets verbatim; touched sheets keep structure)
// ---------------------------------------------------------------------------

const ODS_PRESERVE = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:spreadsheet>
  <table:table table:name="Touched">
   <table:table-header-rows>
    <table:table-row><table:table-cell office:value-type="string" office:string-value="head"><text:p>head</text:p></table:table-cell></table:table-row>
   </table:table-header-rows>
   <table:table-row table:visibility="collapse" table:default-cell-style-name="Default"><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>
   <table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="1" office:value-type="string" office:string-value="merged"><text:p>merged</text:p></table:table-cell><table:covered-table-cell office:value-type="float" office:value="99"><text:p>99</text:p></table:covered-table-cell></table:table-row>
  </table:table>
  <table:table table:name="Untouched">
   <table:table-row table:number-rows-repeated="3"><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell></table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>`;

function makePreserveOds(): Uint8Array {
  const repacked: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
    "content.xml": strToU8(ODS_PRESERVE),
  };
  return zipSync(repacked as Record<string, Uint8Array>);
}

describe("ods preservation", () => {
  it("keeps untouched sheets verbatim (repeats not expanded)", () => {
    const wb = readWorkbook(makePreserveOds());
    setCellInput(wb.sheets[0]!, 1, 1, "edited"); // touch only the first sheet
    const content = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]);
    const untouched = content.slice(content.indexOf('table:name="Untouched"'));
    expect(untouched).toContain('table:number-rows-repeated="3"');
  });

  it("keeps row attributes on the touched sheet", () => {
    const wb = readWorkbook(makePreserveOds());
    setCellInput(wb.sheets[0]!, 1, 1, "edited");
    const content = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]);
    const touched = content.slice(content.indexOf('table:name="Touched"'), content.indexOf('table:name="Untouched"'));
    expect(touched).toContain('table:visibility="collapse"');
    expect(touched).toContain('table:default-cell-style-name="Default"');
  });

  it("re-wraps header rows in table:table-header-rows on the touched sheet", () => {
    const wb = readWorkbook(makePreserveOds());
    setCellInput(wb.sheets[0]!, 2, 1, "edited");
    const content = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]);
    const touched = content.slice(content.indexOf('table:name="Touched"'), content.indexOf('table:name="Untouched"'));
    expect(touched).toContain("<table:table-header-rows>");
    // the header row (with "head") sits inside the group
    const hdrStart = touched.indexOf("<table:table-header-rows>");
    const hdrEnd = touched.indexOf("</table:table-header-rows>");
    expect(touched.slice(hdrStart, hdrEnd)).toContain("head");
  });

  it("keeps covered-cell content on the touched sheet", () => {
    const wb = readWorkbook(makePreserveOds());
    setCellInput(wb.sheets[0]!, 1, 1, "edited");
    const content = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]);
    const touched = content.slice(content.indexOf('table:name="Touched"'), content.indexOf('table:name="Untouched"'));
    expect(touched).toContain("covered-table-cell");
    expect(touched).toContain('office:value="99"'); // hidden merged-away value survives
  });

  it("preserves the tail of a content run repeated beyond the expansion cap", () => {
    const content = ODS_PRESERVE.replace(
      'table:number-rows-repeated="3"',
      'table:number-rows-repeated="1030"',
    );
    const repacked: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {
      mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
      "content.xml": strToU8(content),
    };
    const wb = readWorkbook(zipSync(repacked as Record<string, Uint8Array>));
    setCellInput(wb.sheets[1]!, 1, 1, "edited"); // touch the repeated sheet itself
    const out = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]);
    const sheet2 = out.slice(out.indexOf('table:name="Untouched"'));
    expect(sheet2).toContain('table:number-rows-repeated="6"'); // 1030 - 1024 tail kept
    expect(sheet2).toContain('office:value="7"');
  });
});

// ---------------------------------------------------------------------------
// error surface
// ---------------------------------------------------------------------------

describe("error surface", () => {
  it("rejects CFB containers (encrypted xlsx / legacy xls) with a clear message", () => {
    const cfb = new Uint8Array(16);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0]);
    expect(() => readWorkbook(cfb)).toThrow(/password-protected or legacy/);
  });

  it("rejects a non-archive with a clear message", () => {
    expect(() => readWorkbook(new Uint8Array([1, 2, 3, 4, 5]))).toThrow(/not a valid workbook/);
  });

  it("rejects corrupt essential XML with a clear message", () => {
    const bad = zipSync({
      "xl/workbook.xml": strToU8("<workbook><unclosed"),
    });
    expect(() => readWorkbook(bad)).toThrow(/corrupt workbook XML/);
  });

  it("degrades gracefully when an optional part is corrupt", () => {
    // styles.xml is already the invalid "<<STYLES-MARKER>>" in the base fixture.
    const wb = readWorkbook(makeXlsx());
    expect(wb.sheets[0]!.cells.get("1:1")?.value).toBe("2");
  });
});
