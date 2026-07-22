import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ` +
  `xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ` +
  `xmlns:xlink="http://www.w3.org/1999/xlink"`;

const CONTENT = `<?xml version="1.0"?>
<office:document-content ${NS}>
 <office:body><office:spreadsheet>
  <table:table table:name="Sheet1">
   <table:table-column table:number-columns-repeated="3"/>
   <table:table-row><table:table-cell office:value-type="string"><text:p>Q1</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell></table:table-row>
   <table:table-row><table:table-cell office:value-type="string"><text:p>Q2</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="20"><text:p>20</text:p></table:table-cell></table:table-row>
   <table:table-row>
    <table:table-cell>
     <draw:frame svg:x="4cm" svg:y="1cm" svg:width="8cm" svg:height="6cm" table:end-cell-address="Sheet1.J17" table:end-x="0cm" table:end-y="0cm">
      <draw:object xlink:href="./Object 1"/>
     </draw:frame>
    </table:table-cell>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>`;

const OBJ = `<?xml version="1.0"?>
<office:document-content ${NS}>
 <office:body><office:chart>
  <chart:chart chart:class="chart:bar">
   <chart:title><text:p>Sales</text:p></chart:title>
   <chart:legend chart:legend-position="bottom"/>
   <chart:plot-area>
    <chart:series chart:values-cell-range-address="Sheet1.B1:Sheet1.B2" chart:label-cell-address="Sheet1.B1"/>
    <chart:categories chart:cell-range-address="Sheet1.A1:Sheet1.A2"/>
   </chart:plot-area>
  </chart:chart>
 </office:chart></office:body>
</office:document-content>`;

function ods(): Uint8Array {
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(CONTENT),
    "Object 1/content.xml": strToU8(OBJ),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/vnd.oasis.opendocument.chart"/></manifest:manifest>`),
  });
}

describe("ods chart reader", () => {
  it("parses an embedded bar chart into the model", () => {
    const wb = readWorkbook(ods());
    const charts = wb.sheets[0].charts!;
    expect(charts).toHaveLength(1);
    const ch = charts[0];
    expect(ch.kind).toBe("column");
    expect(ch.title).toBe("Sales");
    expect(ch.legend).toEqual({ show: true, pos: "bottom" });
    expect(ch.series[0].values).toEqual({ ref: "Sheet1!B1:B2" });
    expect(ch.categories).toEqual({ ref: "Sheet1!A1:A2" });
    // frame in the 3rd row, 1st cell; end at J17.
    expect(ch.anchor.fromRow).toBe(3);
    expect(ch.anchor.fromCol).toBe(1);
    expect(ch.anchor.toCol).toBe(10);
    expect(ch.anchor.toRow).toBe(17);
    expect(ch.original?.objectDir).toBe("Object 1");
  });
});
