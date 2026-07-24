import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, setCellInput, writeWorkbook } from "../../index";
import { setOdsComment, setOdsDataValidation, setOdsHyperlink } from "./write";
import { getCell } from "../../core/model";

const contentOf = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["content.xml"]!);

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:xlink="http://www.w3.org/1999/xlink" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/"`;

function ods(cells: string): Uint8Array {
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1"><table:table-row>${cells}</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods hyperlink + comment authoring", () => {
  it("reads an existing text:a link and office:annotation", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.com">site</text:a></text:p></table:table-cell>` +
      `<table:table-cell office:value-type="string"><office:annotation><dc:creator>Ada</dc:creator><text:p>look here</text:p></office:annotation><text:p>x</text:p></table:table-cell>`,
    ));
    expect(getCell(wb.sheets[0], 1, 1)?.link).toEqual({ href: "https://example.com" });
    expect(getCell(wb.sheets[0], 1, 2)?.comments?.[0]).toEqual({ author: "Ada", text: "look here" });
  });

  it("reads an internal link and normalises it to Sheet!A1", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="#Sheet1.B2">go</text:a></text:p></table:table-cell>`,
    ));
    expect(getCell(wb.sheets[0], 1, 1)?.link).toEqual({ href: "Sheet1!B2", internal: true });
  });

  it("authors a hyperlink and a note that round-trip", () => {
    const wb = readWorkbook(ods(`<table:table-cell office:value-type="string"><text:p>hello</text:p></table:table-cell>`));
    setOdsHyperlink(wb.sheets[0], 1, 1, { href: "https://a.test" });
    setOdsComment(wb.sheets[0], 1, 2, "please review", "Bob");
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.link?.href).toBe("https://a.test");
    expect(getCell(re.sheets[0], 1, 2)?.comments?.[0]).toEqual({ author: "Bob", text: "please review" });
  });

  it("reads an existing list content-validation", () => {
    const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
      `<table:content-validations><table:content-validation table:name="v1" table:condition='of:cell-content-is-in-list("Yes";"No")' table:allow-empty-cell="true"/></table:content-validations>` +
      `<table:table table:name="Sheet1"><table:table-row>` +
      `<table:table-cell table:content-validation-name="v1" office:value-type="string"><text:p>Yes</text:p></table:table-cell>` +
      `</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const wb = readWorkbook(zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
      "content.xml": strToU8(content),
      "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
    }));
    const dv = wb.sheets[0].validations?.[0];
    expect(dv?.values).toEqual(["Yes", "No"]);
    expect(dv?.ranges[0]).toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
  });

  it("authors a list validation that round-trips", () => {
    const wb = readWorkbook(ods(`<table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell>`));
    setOdsDataValidation(wb, wb.sheets[0], [{ r1: 1, c1: 1, r2: 5, c2: 1 }], { values: ["A", "B", "C"], allowBlank: true });
    const re = readWorkbook(writeWorkbook(wb));
    const dv = re.sheets[0].validations?.[0];
    expect(dv?.values).toEqual(["A", "B", "C"]);
    // The whole authored range carries the validation.
    expect(dv?.ranges.length).toBe(5);
  });

  it("editing a cell's value preserves its note position, formatting and validation", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell table:content-validation-name="v1" office:value-type="float" office:value="5">` +
      `<office:annotation svg:x="1.2cm" svg:y="3.4cm" office:name="A1"><dc:creator>Ada</dc:creator><dc:date>2020-01-02T03:04:05</dc:date><text:p>keep me</text:p></office:annotation>` +
      `<text:p>5</text:p></table:table-cell>`,
    ));
    setCellInput(wb.sheets[0], 1, 1, "9"); // a plain value edit, not touching the note/validation
    const xml = contentOf(writeWorkbook(wb));
    expect(xml).toContain('office:value="9"'); // new value written
    expect(xml).toContain('svg:x="1.2cm"'); // annotation position preserved
    expect(xml).toContain("keep me"); // annotation text preserved
    expect(xml).toContain('table:content-validation-name="v1"'); // validation reference preserved
  });

  it("authoring a note keeps the cell value and an existing hyperlink", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string" office:string-value="site"><text:p><text:a xlink:href="https://x.test">site</text:a></text:p></table:table-cell>`,
    ));
    setOdsComment(wb.sheets[0], 1, 1, "review this", "Bob");
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.value).toBe("site");
    expect(getCell(re.sheets[0], 1, 1)?.link?.href).toBe("https://x.test"); // link untouched by note authoring
    expect(getCell(re.sheets[0], 1, 1)?.comments?.[0]).toEqual({ author: "Bob", text: "review this" });
  });

  it("editing one note preserves a second annotation only when notes are untouched", () => {
    // Two notes; editing the VALUE (not the notes) keeps both verbatim.
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="float" office:value="1">` +
      `<office:annotation><text:p>first</text:p></office:annotation>` +
      `<office:annotation><text:p>second</text:p></office:annotation>` +
      `<text:p>1</text:p></table:table-cell>`,
    ));
    setCellInput(wb.sheets[0], 1, 1, "2");
    const xml = contentOf(writeWorkbook(wb));
    expect(xml).toContain("first");
    expect(xml).toContain("second");
    expect((xml.match(/office:annotation/g) ?? []).length).toBeGreaterThanOrEqual(2);
  });

  it("preserves a non-list validation (and does not show it as a dropdown) across a value edit", () => {
    const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
      `<table:content-validations><table:content-validation table:name="w1" table:condition="of:cell-content-is-between(1,10)" table:allow-empty-cell="true"><table:error-message table:message-type="stop"><text:p>1-10 only</text:p></table:error-message></table:content-validation></table:content-validations>` +
      `<table:table table:name="Sheet1"><table:table-row>` +
      `<table:table-cell table:content-validation-name="w1" office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>` +
      `</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const wb = readWorkbook(zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
      "content.xml": strToU8(content),
      "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
    }));
    expect(wb.sheets[0].validations).toBeUndefined(); // a between-rule is not a dropdown
    setCellInput(wb.sheets[0], 1, 1, "7");
    const xml = contentOf(writeWorkbook(wb));
    expect(xml).toContain('office:value="7"'); // value edited
    expect(xml).toContain('table:content-validation-name="w1"'); // validation reference preserved
    expect(xml).toContain("cell-content-is-between(1,10)"); // the definition + error message preserved
    expect(xml).toContain("1-10 only");
  });

  it("removes a hyperlink when set to null", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://x.test">x</text:a></text:p></table:table-cell>`,
    ));
    setOdsHyperlink(wb.sheets[0], 1, 1, null);
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.link).toBeUndefined();
  });
});
