import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../index";
import { addSheet, renameSheet, deleteSheet, moveSheet } from "./sheet-ops";

function twoSheetXlsx(): Uint8Array {
  const ws = (a1: string) => strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${a1}</t></is></c></row></sheetData></worksheet>`);
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Alpha" sheetId="1" r:id="rId1"/><sheet name="Beta" sheetId="2" r:id="rId2"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": ws("A"),
    "xl/worksheets/sheet2.xml": ws("B"),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

const names = (bytes: Uint8Array): string[] => readWorkbook(bytes).sheets.map((s) => s.name);

describe("sheet-ops (xlsx)", () => {
  it("adds a sheet that survives write + re-read", () => {
    const wb = readWorkbook(twoSheetXlsx());
    const i = addSheet(wb, "Gamma");
    expect(i).toBe(2);
    expect(names(writeWorkbook(wb))).toEqual(["Alpha", "Beta", "Gamma"]);
  });
  it("renames a sheet and dedupes against an existing name", () => {
    const wb = readWorkbook(twoSheetXlsx());
    expect(renameSheet(wb, 0, "Renamed")).toBe("Renamed");
    expect(names(writeWorkbook(wb))).toEqual(["Renamed", "Beta"]);
    expect(renameSheet(wb, 0, "Beta")).toBe("Beta (2)"); // collision deduped
  });
  it("deletes a sheet and drops its part", () => {
    const wb = readWorkbook(twoSheetXlsx());
    deleteSheet(wb, 0);
    const out = writeWorkbook(wb);
    expect(names(out)).toEqual(["Beta"]);
    expect(Object.keys(unzipSync(out))).not.toContain("xl/worksheets/sheet1.xml");
  });
  it("refuses to delete the last sheet", () => {
    const wb = readWorkbook(twoSheetXlsx());
    deleteSheet(wb, 0);
    expect(() => deleteSheet(wb, 0)).toThrow(/at least one/);
  });
  it("moves a sheet and the order survives re-read", () => {
    const wb = readWorkbook(twoSheetXlsx());
    moveSheet(wb, 1, 0);
    expect(names(writeWorkbook(wb))).toEqual(["Beta", "Alpha"]);
  });
});

const ODS = strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:spreadsheet>
  <table:table table:name="One"><table:table-column/><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table>
  <table:table table:name="Two"><table:table-column/><table:table-row><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row></table:table>
 </office:spreadsheet></office:body>
</office:document-content>`);
function makeOds(): Uint8Array {
  return zipSync({ mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array, "content.xml": ODS });
}

describe("sheet-ops (ods)", () => {
  it("adds, renames, moves and deletes tables", () => {
    let wb = readWorkbook(makeOds());
    addSheet(wb, "Three");
    renameSheet(wb, 0, "First");
    moveSheet(wb, 2, 0);
    expect(readWorkbook(writeWorkbook(wb)).sheets.map((s) => s.name)).toEqual(["Three", "First", "Two"]);
    wb = readWorkbook(writeWorkbook(wb));
    deleteSheet(wb, 0);
    expect(readWorkbook(writeWorkbook(wb)).sheets.map((s) => s.name)).toEqual(["First", "Two"]);
  });
});
