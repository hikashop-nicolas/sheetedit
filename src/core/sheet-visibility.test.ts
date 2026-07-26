import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "./workbook";
import { setSheetVisibility, visibleSheetCount } from "./sheet-ops";

// A hidden sheet is a real part of a workbook's shape: Excel and Calc both draw no tab for one, so
// showing it would misrepresent the file. xlsx states it on <sheet state>, ODF on table:display.

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** A two-sheet xlsx whose second sheet carries `state`, when given one. */
function xlsx(state?: string): Uint8Array {
  const sheet = (n: number) => strToU8(`<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>${n}</v></c></row></sheetData></worksheet>`);
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets>` +
      `<sheet name="One" sheetId="1" r:id="rId1"/>` +
      `<sheet name="Two" sheetId="2" r:id="rId2"${state ? ` state="${state}"` : ""}/>` +
      `</sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${R}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": sheet(1),
    "xl/worksheets/sheet2.xml": sheet(2),
  });
}

const wbXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["xl/workbook.xml"]!);

describe("reading sheet visibility", () => {
  it("reads a hidden sheet, and a very hidden one", () => {
    expect(readWorkbook(xlsx()).sheets[1]!.visibility).toBeUndefined();
    expect(readWorkbook(xlsx("hidden")).sheets[1]!.visibility).toBe("hidden");
    expect(readWorkbook(xlsx("veryHidden")).sheets[1]!.visibility).toBe("veryHidden");
  });

  it("still reads the hidden sheet's cells: hidden is not absent", () => {
    const wb = readWorkbook(xlsx("hidden"));
    expect(wb.sheets).toHaveLength(2);
    expect(wb.sheets[1]!.cells.size).toBe(1);
  });

  it("counts what the user can actually see", () => {
    expect(visibleSheetCount(readWorkbook(xlsx()))).toBe(2);
    expect(visibleSheetCount(readWorkbook(xlsx("hidden")))).toBe(1);
  });
});

describe("writing sheet visibility", () => {
  it("hides and shows again, round-tripping through the reader", () => {
    const wb = readWorkbook(xlsx());
    setSheetVisibility(wb, 1, "hidden");
    expect(wbXml(writeWorkbook(wb))).toMatch(/name="Two"[^>]*state="hidden"/);
    expect(readWorkbook(writeWorkbook(wb)).sheets[1]!.visibility).toBe("hidden");

    setSheetVisibility(wb, 1, undefined);
    expect(wbXml(writeWorkbook(wb))).not.toContain("state=");
    expect(readWorkbook(writeWorkbook(wb)).sheets[1]!.visibility).toBeUndefined();
  });

  it("writes very hidden as its own state", () => {
    const wb = readWorkbook(xlsx());
    setSheetVisibility(wb, 1, "veryHidden");
    expect(readWorkbook(writeWorkbook(wb)).sheets[1]!.visibility).toBe("veryHidden");
  });

  it("refuses to hide the last visible sheet", () => {
    const wb = readWorkbook(xlsx("hidden"));
    expect(() => setSheetVisibility(wb, 0, "hidden")).toThrow(/at least one visible sheet/);
    expect(wb.sheets[0]!.visibility).toBeUndefined();
  });

  it("lets an already-hidden sheet be re-hidden without tripping the guard", () => {
    const wb = readWorkbook(xlsx("hidden"));
    expect(() => setSheetVisibility(wb, 1, "veryHidden")).not.toThrow();
  });
});

// --- ODF, which states this somewhere else entirely ---------------------------
// Not on the table element: ODF keeps sheet visibility in the sheet's TABLE STYLE. LibreOffice
// ignores a table:display attribute written on <table:table>, which is how this was found.

const ODS_NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"` +
  ` xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"`;

/** A two-sheet ods; each sheet names a table style, and `styles` defines them. */
function ods(styles: string, style1: string, style2: string): Uint8Array {
  const table = (name: string, style: string): string =>
    `<table:table table:name="${name}" table:style-name="${style}">` +
    `<table:table-row><table:table-cell office:value-type="string"><text:p>${name}</text:p></table:table-cell></table:table-row></table:table>`;
  return zipSync({
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(
      `<?xml version="1.0"?><office:document-content ${ODS_NS}>` +
      `<office:automatic-styles>${styles}</office:automatic-styles>` +
      `<office:body><office:spreadsheet>${table("One", style1)}${table("Two", style2)}` +
      `</office:spreadsheet></office:body></office:document-content>`),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

const tableStyle = (name: string, display: string): string =>
  `<style:style style:name="${name}" style:family="table"><style:table-properties table:display="${display}"/></style:style>`;

const contentOf = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["content.xml"]!);

describe("ODF sheet visibility", () => {
  const shownAndHidden = (): Uint8Array =>
    ods(tableStyle("ta1", "true") + tableStyle("ta2", "false"), "ta1", "ta2");

  it("reads it out of the table style", () => {
    expect(readWorkbook(shownAndHidden(), "x.ods").sheets.map((s) => s.visibility)).toEqual([undefined, "hidden"]);
  });

  it("ignores a table:display written on the element, as LibreOffice does", () => {
    // The attribute in the wrong place must not be honoured, or we would disagree with every
    // other reader about which sheets a file shows.
    const wrong = zipSync({
      mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
      "content.xml": strToU8(
        `<?xml version="1.0"?><office:document-content ${ODS_NS}><office:automatic-styles/>` +
        `<office:body><office:spreadsheet><table:table table:name="One" table:display="false">` +
        `<table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row>` +
        `</table:table></office:spreadsheet></office:body></office:document-content>`),
      "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
    });
    expect(readWorkbook(wrong, "x.ods").sheets[0]!.visibility).toBeUndefined();
  });

  it("hides a sheet by putting the flag in its style, and reads it back", () => {
    const wb = readWorkbook(ods(tableStyle("ta1", "true"), "ta1", "ta1"), "x.ods");
    setSheetVisibility(wb, 1, "hidden");
    expect(contentOf(writeWorkbook(wb))).toMatch(/<style:table-properties[^>]*table:display="false"/);
    expect(readWorkbook(writeWorkbook(wb), "y.ods").sheets.map((s) => s.visibility)).toEqual([undefined, "hidden"]);
  });

  it("does not drag a sheet sharing the same style into hiding with it", () => {
    // Both sheets start on ta1; hiding one must clone the style rather than edit it in place.
    const wb = readWorkbook(ods(tableStyle("ta1", "true"), "ta1", "ta1"), "x.ods");
    setSheetVisibility(wb, 1, "hidden");
    expect(readWorkbook(writeWorkbook(wb), "y.ods").sheets[0]!.visibility).toBeUndefined();
  });

  it("shows a hidden sheet again", () => {
    const wb = readWorkbook(shownAndHidden(), "x.ods");
    setSheetVisibility(wb, 1, undefined);
    expect(readWorkbook(writeWorkbook(wb), "y.ods").sheets[1]!.visibility).toBeUndefined();
  });

  it("stores very hidden as ordinary hidden, since ODF has no such state", () => {
    const wb = readWorkbook(ods(tableStyle("ta1", "true"), "ta1", "ta1"), "x.ods");
    setSheetVisibility(wb, 1, "veryHidden");
    expect(readWorkbook(writeWorkbook(wb), "y.ods").sheets[1]!.visibility).toBe("hidden");
  });
});
