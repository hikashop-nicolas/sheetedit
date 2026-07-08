import { afterEach, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { getCell } from "../core/model";
import { isoToSerial } from "../core/dates";
import { setLocale } from "../core/i18n";
import { readCsv } from "./csv/read";
import { writeCsv } from "./csv/write";
import { readWorkbook, setCellInput, writeWorkbook } from "../core/workbook";

afterEach(() => setLocale("en"));

// --- xlsx fixtures ---------------------------------------------------------

const makeXlsx = (sheetXml: string) =>
  zipSync({
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
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetXml}</sheetData></worksheet>`,
    ),
  });

describe("xlsx t=\"d\" date cells", () => {
  it("read as serials and display as dates", () => {
    const wb = readWorkbook(makeXlsx('<row r="1"><c r="A1" t="d"><v>2026-07-08</v></c></row>'));
    const cell = getCell(wb.sheets[0]!, 1, 1)!;
    expect(cell.kind).toBe("n");
    expect(Number(cell.value)).toBe(isoToSerial("2026-07-08"));
    expect(cell.display).toBe("2026-07-08");
  });

  it("stay untouched on save when not edited", () => {
    const wb = readWorkbook(makeXlsx('<row r="1"><c r="A1" t="d"><v>2026-07-08</v></c></row>'));
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]!);
    expect(xml).toContain('t="d"');
    expect(xml).toContain("2026-07-08");
  });

  it("an edit converts to a serial with a persisted date format", () => {
    const wb = readWorkbook(makeXlsx('<row r="1"><c r="A1" t="d"><v>2026-07-08</v></c></row>'));
    setCellInput(wb.sheets[0]!, 1, 1, "2026-12-25");
    const files = unzipSync(writeWorkbook(wb));
    const xml = strFromU8(files["xl/worksheets/sheet1.xml"]!);
    expect(xml).not.toContain('t="d"');
    expect(xml).toContain(`<v>${isoToSerial("2026-12-25")}</v>`);
    const styles = strFromU8(files["xl/styles.xml"]!);
    expect(styles).toContain('formatCode="yyyy-mm-dd"');
    expect(xml).toMatch(/<c r="A1" s="\d+"/);
  });
});

describe("typed dates in xlsx", () => {
  it("store a serial and mint styles.xml when the workbook has none", () => {
    const wb = readWorkbook(makeXlsx('<row r="1"><c r="A1"><v>1</v></c></row>'));
    setCellInput(wb.sheets[0]!, 1, 2, "12/25/2026"); // en locale: m/d/y
    const cell = getCell(wb.sheets[0]!, 1, 2)!;
    expect(cell.kind).toBe("n");
    expect(Number(cell.value)).toBe(isoToSerial("2026-12-25"));
    expect(cell.display).toBe("12/25/2026");
    const files = unzipSync(writeWorkbook(wb));
    expect(strFromU8(files["xl/styles.xml"]!)).toContain('formatCode="mm/dd/yyyy"');
    expect(strFromU8(files["[Content_Types].xml"]!)).toContain("/xl/styles.xml");
    expect(strFromU8(files["xl/_rels/workbook.xml.rels"]!)).toContain("styles.xml");
  });

  it("follow the French day order and accept comma decimals under fr", () => {
    setLocale("fr");
    const wb = readWorkbook(makeXlsx('<row r="1"><c r="A1"><v>1</v></c></row>'));
    const sheet = wb.sheets[0]!;
    setCellInput(sheet, 1, 2, "8/7/2026");
    expect(Number(getCell(sheet, 1, 2)!.value)).toBe(isoToSerial("2026-07-08"));
    expect(getCell(sheet, 1, 2)!.numFmt).toBe("dd/mm/yyyy");
    setCellInput(sheet, 1, 3, "3,5");
    expect(getCell(sheet, 1, 3)).toMatchObject({ kind: "n", value: "3.5" });
  });

  it("comma decimals stay text under en", () => {
    const wb = readWorkbook(makeXlsx('<row r="1"><c r="A1"><v>1</v></c></row>'));
    setCellInput(wb.sheets[0]!, 1, 2, "3,5");
    expect(getCell(wb.sheets[0]!, 1, 2)!.kind).toBe("s");
  });

  it("percent input divides by 100 and formats as percent", () => {
    const wb = readWorkbook(makeXlsx('<row r="1"><c r="A1"><v>1</v></c></row>'));
    const sheet = wb.sheets[0]!;
    setCellInput(sheet, 1, 2, "50%");
    expect(getCell(sheet, 1, 2)).toMatchObject({ kind: "n", value: "0.5", numFmt: "0%" });
    setCellInput(sheet, 1, 3, "12.5%");
    expect(getCell(sheet, 1, 3)).toMatchObject({ value: "0.125", numFmt: "0.00%" });
  });
});

// --- ods fixtures ----------------------------------------------------------

const ODS_CONTENT = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:spreadsheet>
  <table:table table:name="Sheet1">
   <table:table-column table:number-columns-repeated="4"/>
   <table:table-row>
    <table:table-cell office:value-type="date" office:date-value="2026-07-08"><text:p>08/07/2026</text:p></table:table-cell>
    <table:table-cell office:value-type="currency" office:currency="EUR" office:value="10"><text:p>10,00 €</text:p></table:table-cell>
    <table:table-cell office:value-type="time" office:time-value="PT13H30M0S"><text:p>13:30</text:p></table:table-cell>
    <table:table-cell office:value-type="percentage" office:value="0.5"><text:p>50 %</text:p></table:table-cell>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>`;

const makeOds = () =>
  zipSync({
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(ODS_CONTENT),
  } as Record<string, Uint8Array>);

describe("ods typed cells", () => {
  it("read dates/times as serials with the producer's display text", () => {
    const wb = readWorkbook(makeOds());
    const s = wb.sheets[0]!;
    expect(Number(getCell(s, 1, 1)!.value)).toBe(isoToSerial("2026-07-08"));
    expect(getCell(s, 1, 1)!.display).toBe("08/07/2026");
    expect(Number(getCell(s, 1, 3)!.value)).toBeCloseTo(0.5625, 10);
  });

  it("keep their ODF value type when edited", () => {
    const wb = readWorkbook(makeOds());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 1, "2026-12-25");
    setCellInput(s, 1, 2, "12"); // currency cell: numeric edit keeps type + currency
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(xml).toContain('office:date-value="2026-12-25"');
    expect(xml).not.toContain('office:date-value="2026-07-08"');
    expect(xml).toContain('office:currency="EUR"');
    expect(xml).toMatch(/office:value-type="currency"[^>]*office:value="12"|office:value="12"[^>]*office:value-type="currency"/);
  });

  it("a percentage edit stays a percentage", () => {
    const wb = readWorkbook(makeOds());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 4, "75%");
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(xml).toContain('office:value-type="percentage"');
    expect(xml).toContain('office:value="0.75"');
  });

  it("a time edit round-trips as a duration", () => {
    const wb = readWorkbook(makeOds());
    const s = wb.sheets[0]!;
    setCellInput(s, 1, 3, "09:15");
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(xml).toContain('office:time-value="PT9H15M0S"');
  });
});

describe("csv stays literal", () => {
  it("typed dates and percents keep their text", () => {
    const wb = readCsv("a,b\n");
    const s = wb.sheets[0]!;
    setCellInput(s, 2, 1, "2026-07-08");
    setCellInput(s, 2, 2, "50%");
    expect(getCell(s, 2, 1)!.kind).toBe("s");
    expect(getCell(s, 2, 2)!.kind).toBe("s");
    expect(writeCsv(wb)).toBe("a,b\n2026-07-08,50%\n");
  });
});
