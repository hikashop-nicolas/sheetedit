import { afterEach, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync } from "fflate";
import { getCell } from "../../core/model";
import { readWorkbook, setCellInput, writeWorkbook } from "../../core/workbook";
import { blankWorkbookBytes, createWorkbook } from "../../core/blank";
import { setLocale } from "../../core/i18n";
import { looksLikeXls } from "../xls";
import { createSheetEditor } from "../../core/editor";
import { importedSaveNotice, readHtmlWorkbook, readSpreadsheetMl, saveFormat, setSaveFormat, sniffSheetText } from ".";

afterEach(() => setLocale("en"));

const XML = '<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="S"><Table><Row><Cell><Data ss:Type="Number">5</Data></Cell></Row></Table></Worksheet></Workbook>';
const EXCEL_HTML = '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table><tr><td x:num>5</td></tr></table></body></html>';

describe("sniffSheetText", () => {
  const u8 = (s: string): Uint8Array => strToU8(s);
  it("recognizes Excel 2003 XML under any name", () => {
    for (const ext of [".xls", ".xml", undefined]) expect(sniffSheetText(u8(XML), ext)).toBe("spreadsheetml");
    expect(sniffSheetText(u8('﻿<?xml version="1.0"?><ss:Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"/>'), ".xml")).toBe("spreadsheetml");
  });
  it("recognizes Excel's HTML under any name, and any HTML table under .xls", () => {
    for (const ext of [".xls", ".htm", ".html", undefined]) expect(sniffSheetText(u8(EXCEL_HTML), ext)).toBe("html");
    const plain = "<!DOCTYPE html><html><body><h1>Statement</h1><table><tr><td>1</td></tr></table></body></html>";
    expect(sniffSheetText(u8(plain), ".xls")).toBe("html");
    expect(sniffSheetText(u8("<table><tr><td>1</td></tr></table>"), "xls")).toBe("html");
    expect(sniffSheetText(u8("MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=\"b\"\r\n\r\n--b\r\nContent-Type: text/html\r\n\r\n<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\"></html>"), ".mht")).toBe("html");
  });
  it("leaves ordinary pages, XML files and binaries alone", () => {
    const page = "<!DOCTYPE html><html><head><title>News</title></head><body><table><tr><td>1</td></tr></table></body></html>";
    expect(sniffSheetText(u8(page), ".html")).toBeNull();
    expect(sniffSheetText(u8(page), ".htm")).toBeNull();
    expect(sniffSheetText(u8(page))).toBeNull();
    expect(sniffSheetText(u8("<!DOCTYPE html><html><body><p>No table here</p></body></html>"), ".xls")).toBeNull();
    const word = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><body><table><tr><td>1</td></tr></table></body></html>';
    expect(sniffSheetText(u8(word), ".htm")).toBeNull();
    for (const xml of ['<?xml version="1.0"?><project><modelVersion>4.0.0</modelVersion></project>', '<?xml version="1.0"?><rss version="2.0"><channel/></rss>', '<svg xmlns="http://www.w3.org/2000/svg"/>',
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>']) {
      expect(sniffSheetText(u8(xml), ".xml")).toBeNull();
      expect(sniffSheetText(u8(xml), ".xls")).toBeNull();
    }
    expect(sniffSheetText(new Uint8Array([0x50, 0x4b, 3, 4, 0, 0, 0, 0]), ".xls")).toBeNull();
    expect(sniffSheetText(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]), ".xls")).toBeNull();
    expect(sniffSheetText(u8("a,b,c\n1,2,3\n"), ".xls")).toBeNull();
  });
  it("finds a table past the first 64 KB of a .xls page", () => {
    const big = "<html><body>" + "<p>filler text</p>".repeat(5000) + "<table><tr><td>1</td></tr></table></body></html>";
    expect(sniffSheetText(u8(big), ".xls")).toBe("html");
    expect(sniffSheetText(u8(big), ".html")).toBeNull();
  });
});

describe("saving imported workbooks", () => {
  it("saves as xlsx by default and as a real .xls once asked", () => {
    const wb = readHtmlWorkbook(EXCEL_HTML);
    expect(saveFormat(wb)).toBe("xlsx");
    expect(importedSaveNotice(wb)).toContain("Excel (.xlsx)");
    const xlsx = writeWorkbook(wb);
    expect(unzipSync(xlsx)["xl/workbook.xml"]).toBeTruthy();
    setSaveFormat(wb, "xls");
    expect(saveFormat(wb)).toBe("xls");
    expect(importedSaveNotice(wb)).toContain("(.xls)");
    setCellInput(wb.sheets[0]!, 1, 2, "=A1*2");
    const xls = writeWorkbook(wb);
    expect(looksLikeXls(xls)).toBe(true);
    const back = readWorkbook(xls);
    expect(getCell(back.sheets[0]!, 1, 1)?.value).toBe("5");
    expect(getCell(back.sheets[0]!, 1, 2)?.formula).toBe("A1*2");
    expect(back.importedFrom).toBeUndefined();
    setSaveFormat(wb, "xlsx");
    expect(saveFormat(wb)).toBe("xlsx");
  });

  it("keeps styles, merges and names through a BIFF8 save", () => {
    const wb = readSpreadsheetMl('<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Styles><Style ss:ID="b"><Font ss:Bold="1"/><Interior ss:Color="#FFFF00" ss:Pattern="Solid"/></Style></Styles><Names><NamedRange ss:Name="Two" ss:RefersTo="=S!R1C1:R1C2"/></Names><Worksheet ss:Name="S"><Table><Row><Cell ss:StyleID="b" ss:MergeAcross="1"><Data ss:Type="String">Title</Data></Cell></Row></Table></Worksheet></Workbook>');
    setSaveFormat(wb, "xls");
    const back = readWorkbook(writeWorkbook(wb));
    const c = getCell(back.sheets[0]!, 1, 1)!;
    expect(c.cellStyle?.bold).toBe(true);
    expect(c.cellStyle?.bg?.toLowerCase()).toBe("#ffff00");
    expect(back.sheets[0]!.merges).toContainEqual({ r1: 1, c1: 1, r2: 1, c2: 2 });
    expect(back.definedNames?.get("Two")).toBe("S!$A$1:$B$1");
  });
});

describe("the editor on an imported export", () => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
  it("saves an HTML .xls as a real .xls and other names as .xlsx, even unedited", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8(EXCEL_HTML), { fileName: "report.xls" });
    expect(ed.importedFrom()).toBe("html");
    expect(ed.importNotice()).toContain("(.xls)");
    expect(looksLikeXls(await ed.getBytes())).toBe(true);
    ed.setSaveFormat("xlsx");
    expect(unzipSync(await ed.getBytes())["xl/workbook.xml"]).toBeTruthy();
    ed.destroy();
    const ed2 = createSheetEditor(host, strToU8(XML), { fileName: "export.xml" });
    expect(ed2.importedFrom()).toBe("spreadsheetml");
    expect(unzipSync(await ed2.getBytes())["xl/workbook.xml"]).toBeTruthy();
    ed2.destroy();
    host.remove();
  });
});

describe("blank workbooks", () => {
  for (const format of ["xlsx", "xls", "ods"] as const) {
    it(`writes a ${format} that reopens with one empty sheet`, () => {
      const bytes = blankWorkbookBytes(format);
      if (format === "xls") expect(looksLikeXls(bytes)).toBe(true);
      else {
        const files = unzipSync(bytes);
        if (format === "ods") expect(strFromU8(files["mimetype"]!)).toBe("application/vnd.oasis.opendocument.spreadsheet");
        else expect(files["xl/workbook.xml"]).toBeTruthy();
      }
      const wb = readWorkbook(bytes);
      expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1"]);
      expect([...wb.sheets[0]!.cells.values()].filter((c) => c.value !== "")).toHaveLength(0);
    });
  }

  it("names the sheet for the locale or as asked, and can be edited and saved", () => {
    setLocale("fr");
    expect(createWorkbook("xlsx").sheets[0]!.name).toBe("Feuille1");
    setLocale("de");
    expect(createWorkbook("ods").sheets[0]!.name).toBe("Tabelle1");
    const wb = createWorkbook("ods", { sheetName: "Budget" });
    setCellInput(wb.sheets[0]!, 2, 2, "42");
    const back = readWorkbook(writeWorkbook(wb));
    expect(back.sheets[0]!.name).toBe("Budget");
    expect(getCell(back.sheets[0]!, 2, 2)?.value).toBe("42");
  });
});
