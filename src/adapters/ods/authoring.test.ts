import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { setOdsComment, setOdsDataValidation, setOdsHyperlink } from "./write";
import { getCell } from "../../core/model";

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:xlink="http://www.w3.org/1999/xlink" ` +
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

  it("removes a hyperlink when set to null", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://x.test">x</text:a></text:p></table:table-cell>`,
    ));
    setOdsHyperlink(wb.sheets[0], 1, 1, null);
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.link).toBeUndefined();
  });
});
