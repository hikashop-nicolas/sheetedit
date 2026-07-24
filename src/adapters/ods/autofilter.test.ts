import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { setOdsAutoFilter } from "./write";

const NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`;

function ods(dbRanges = ""): Uint8Array {
  const sc = (t: string) => `<table:table-cell office:value-type="string"><text:p>${t}</text:p></table:table-cell>`;
  const rows = `<table:table-row>${sc("h1")}${sc("h2")}</table:table-row><table:table-row>${sc("a")}${sc("b")}</table:table-row>`;
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>${dbRanges}` +
    `<table:table table:name="Sheet1">${rows}</table:table></office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods autofilter", () => {
  it("reads a database-range with filter buttons into sheet.autoFilter", () => {
    const db = `<table:database-ranges><table:database-range table:display-filter-buttons="true" table:target-range-address="Sheet1.A1:Sheet1.B2"/></table:database-ranges>`;
    const wb = readWorkbook(ods(db));
    expect(wb.sheets[0].autoFilter).toEqual({ r1: 1, c1: 1, r2: 2, c2: 2 });
  });

  it("authors and removes an autofilter database-range", () => {
    const wb = readWorkbook(ods());
    setOdsAutoFilter(wb, wb.sheets[0], { r1: 1, c1: 1, r2: 2, c2: 2 });
    let re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].autoFilter).toEqual({ r1: 1, c1: 1, r2: 2, c2: 2 });
    setOdsAutoFilter(wb, wb.sheets[0], null);
    re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].autoFilter).toBeUndefined();
  });

  it("persists filtered (hidden) rows as table:visibility on save", () => {
    const wb = readWorkbook(ods());
    wb.sheets[0].autoFilter = { r1: 1, c1: 1, r2: 2, c2: 2 };
    wb.sheets[0].filterHidden = new Set([2]);
    wb.sheets[0].odsDirty = true;
    const re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].hiddenRows?.has(2)).toBe(true);
  });
});
