import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";

const NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"`;

function ods(): Uint8Array {
  const nc = (v: number) => `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
  const rows = `<table:table-row>${nc(3)}${nc(5)}${nc(2)}${nc(8)}${nc(4)}<table:table-cell/></table:table-row>`;
  const spark = `<calcext:sparkline-groups><calcext:sparkline-group calcext:type="column" calcext:color-series="#0369a3" calcext:color-negative="#c9211e"><calcext:sparklines><calcext:sparkline calcext:cell-address="Sheet1.F1" calcext:data-range="Sheet1.A1:Sheet1.E1"/></calcext:sparklines></calcext:sparkline-group></calcext:sparkline-groups>`;
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet><table:table table:name="Sheet1">${rows}${spark}</table:table></office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods sparklines (calcext)", () => {
  it("reads a calcext sparkline-group into sheet.sparklines", () => {
    const sp = readWorkbook(ods()).sheets[0].sparklines ?? [];
    expect(sp.length).toBe(1);
    expect(sp[0].type).toBe("column");
    expect(sp[0].host).toEqual({ r: 1, c: 6 }); // F1
    expect(sp[0].dataRef).toBe("Sheet1!A1:E1");
    expect(sp[0].color.toLowerCase()).toContain("0369a3");
    expect(sp[0].negColor?.toLowerCase()).toContain("c9211e");
  });
});
