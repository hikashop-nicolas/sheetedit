import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { computeCondVisuals } from "../xlsx/condformat";
import { setOdsCondFormat } from "./write";
import { key } from "../../core/model";

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
  `xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"`;

function ods(): Uint8Array {
  const nc = (v: number) => `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
  const rows = [1, 2, 3, 4, 5].map((v) => `<table:table-row>${nc(v)}${nc(v)}</table:table-row>`).join("");
  const styles = `<office:automatic-styles><style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffc7ce"/></style:style></office:automatic-styles>`;
  const cf = `<calcext:conditional-formats>` +
    `<calcext:conditional-format calcext:target-range-address="Sheet1.A1:Sheet1.A5"><calcext:condition calcext:apply-style-name="ce1" calcext:value="&gt;3" calcext:base-cell-address="Sheet1.A1"/></calcext:conditional-format>` +
    `<calcext:conditional-format calcext:target-range-address="Sheet1.B1:Sheet1.B5"><calcext:color-scale><calcext:color-scale-entry calcext:type="minimum" calcext:color="#f8696b"/><calcext:color-scale-entry calcext:type="maximum" calcext:color="#63be7b"/></calcext:color-scale></calcext:conditional-format>` +
    `</calcext:conditional-formats>`;
  const content = `<?xml version="1.0"?><office:document-content ${NS}>${styles}<office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1">${rows}</table:table>${cf}` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods conditional formatting read", () => {
  it("reads a cellIs highlight and a colour scale", () => {
    const wb = readWorkbook(ods());
    const cfs = wb.sheets[0].condFormats ?? [];
    expect(cfs.length).toBe(2);
    const vis = computeCondVisuals(wb.sheets[0]);
    // Column A: value > 3 gets the ce1 fill; <= 3 does not.
    expect(vis.get(key(4, 1))?.bg?.toLowerCase()).toBe("#ffc7ce");
    expect(vis.get(key(5, 1))?.bg?.toLowerCase()).toBe("#ffc7ce");
    expect(vis.get(key(2, 1))?.bg).toBeUndefined();
    // Column B: every cell gets a colour-scale fill.
    expect(vis.get(key(1, 2))?.bg).toBeTruthy();
    expect(vis.get(key(5, 2))?.bg).toBeTruthy();
  });

  it("authors a cellIs highlight that round-trips", () => {
    const wb = readWorkbook(ods());
    setOdsCondFormat(wb, wb.sheets[0], [{ r1: 1, c1: 3, r2: 5, c2: 3 }], { kind: "cellIs", operator: "greaterThan", value: "2", fill: "#c6efce" });
    const re = readWorkbook(writeWorkbook(wb));
    const cf = (re.sheets[0].condFormats ?? []).find((c) => c.ranges[0]?.c1 === 3);
    expect(cf?.rules[0].type).toBe("cellIs");
    expect(cf?.rules[0].operator).toBe("greaterThan");
    expect(cf?.rules[0].dxf?.bg?.toLowerCase()).toBe("#c6efce");
  });
});
