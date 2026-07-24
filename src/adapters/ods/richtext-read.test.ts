import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { getCell } from "../../core/model";

const NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"`;

function ods(cell: string): Uint8Array {
  const styles = `<office:automatic-styles>` +
    `<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:color="#cc0000"/></style:style>` +
    `<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>` +
    `</office:automatic-styles>`;
  const content = `<?xml version="1.0"?><office:document-content ${NS}>${styles}<office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1"><table:table-row>${cell}</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods rich text", () => {
  it("reads per-run styling from text:span into richRuns", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p>Hi <text:span text:style-name="T1">red</text:span> and <text:span text:style-name="T2">it</text:span></text:p></table:table-cell>`,
    ));
    const cell = getCell(wb.sheets[0], 1, 1);
    expect(cell?.value).toBe("Hi red and it");
    const runs = cell?.richRuns;
    expect(runs?.length).toBe(4);
    expect(runs?.[1]).toMatchObject({ text: "red", bold: true });
    expect(runs?.[1].color?.toLowerCase()).toContain("cc0000");
    expect(runs?.[3]).toMatchObject({ text: "it", italic: true });
  });

  it("leaves richRuns undefined for a single unstyled string", () => {
    const wb = readWorkbook(ods(`<table:table-cell office:value-type="string"><text:p>plain</text:p></table:table-cell>`));
    expect(getCell(wb.sheets[0], 1, 1)?.richRuns).toBeUndefined();
  });
});
