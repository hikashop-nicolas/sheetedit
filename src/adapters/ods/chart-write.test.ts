import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { buildChart, defaultAnchor } from "../../core/chart-build";

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`;

function ods(): Uint8Array {
  const sc = (t: string) => `<table:table-cell office:value-type="string"><text:p>${t}</text:p></table:table-cell>`;
  const nc = (v: number) => `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1"><table:table-column table:number-columns-repeated="3"/>` +
    `<table:table-row>${sc("Qtr")}${sc("North")}${sc("South")}</table:table-row>` +
    `<table:table-row>${sc("Q1")}${nc(10)}${nc(22)}</table:table-row>` +
    `<table:table-row>${sc("Q2")}${nc(30)}${nc(14)}</table:table-row>` +
    `</table:table></office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods chart writer", () => {
  it("a created chart survives write + re-read as an embedded object", () => {
    const wb = readWorkbook(ods());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 }; // A1:C3
    const model = buildChart("Sheet1", "column", rect, { firstRowHeader: true, firstColLabels: true }, "new1", defaultAnchor(rect));
    model.title = "Quarterly";
    model.legend = { show: true, pos: "bottom" };
    (wb.sheets[0].charts ??= []).push(model);

    const out = writeWorkbook(wb);
    const names = Object.keys(unzipSync(out));
    expect(names.some((n) => /^Object \d+\/content\.xml$/.test(n))).toBe(true);
    expect(new TextDecoder().decode(unzipSync(out)["META-INF/manifest.xml"])).toContain("opendocument.chart");

    const re = readWorkbook(out).sheets[0].charts!;
    expect(re).toHaveLength(1);
    expect(re[0].kind).toBe("column");
    expect(re[0].title).toBe("Quarterly");
    expect(re[0].legend.show).toBe(true);
    expect(re[0].series).toHaveLength(2);
    expect(re[0].series[0].values.ref).toBe("Sheet1!B2:B3");
    expect(re[0].categories?.ref).toBe("Sheet1!A2:A3");
  });

  it("orientation (bar) and 100% stacked round-trip through the plot-area style", () => {
    const wb = readWorkbook(ods());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 };
    const m = buildChart("Sheet1", "bar", rect, { firstRowHeader: true, firstColLabels: true }, "b", defaultAnchor(rect));
    m.stacked = true;
    m.percent = true;
    (wb.sheets[0].charts ??= []).push(m);
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].charts![0];
    expect(re.kind).toBe("bar"); // horizontal, from chart:vertical="false"
    expect(re.stacked).toBe(true);
    expect(re.percent).toBe(true);
  });
});
