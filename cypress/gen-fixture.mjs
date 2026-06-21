// Generate minimal valid sample.xlsx and sample.ods fixtures for the e2e tests.
import { strToU8, zipSync } from "fflate";
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
mkdirSync(join(here, "fixtures"), { recursive: true });

// --- xlsx ---
const sheet1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <sheetData>
  <row r="1"><c r="A1" t="inlineStr"><is><t>item</t></is></c><c r="B1" t="inlineStr"><is><t>qty</t></is></c><c r="C1" t="inlineStr"><is><t>total</t></is></c><c r="D1" t="inlineStr"><is><t>price</t></is></c></row>
  <row r="2"><c r="A2" t="inlineStr"><is><t>apples</t></is></c><c r="B2"><v>3</v></c><c r="C2"><f>B2*2</f><v>6</v></c><c r="D2" s="1"><v>3.5</v></c></row>
  <row r="3"><c r="A3" t="inlineStr"><is><t>pears</t></is></c><c r="B3"><v>4</v></c><c r="C3"><f>B3*2</f><v>8</v></c><c r="D3" s="1"><v>4.25</v></c></row>
  <row r="4"><c r="A4" t="inlineStr"><is><t>sum</t></is></c><c r="C4"><f>SUM(C2:C3)</f><v>14</v></c></row>
 </sheetData>
</worksheet>`;

const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0.00"/></numFmts>
 <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs>
</styleSheet>`;

const xlsx = zipSync({
  "[Content_Types].xml": strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  ),
  "xl/styles.xml": strToU8(styles),
  "_rels/.rels": strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  ),
  "xl/workbook.xml": strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  ),
  "xl/_rels/workbook.xml.rels": strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  ),
  "xl/worksheets/sheet1.xml": strToU8(sheet1),
});
writeFileSync(join(here, "fixtures", "sample.xlsx"), xlsx);

// --- ods ---
const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:spreadsheet>
  <table:table table:name="Budget">
   <table:table-column table:number-columns-repeated="3"/>
   <table:table-row>
    <table:table-cell office:value-type="string" office:string-value="item"><text:p>item</text:p></table:table-cell>
    <table:table-cell office:value-type="string" office:string-value="qty"><text:p>qty</text:p></table:table-cell>
    <table:table-cell office:value-type="string" office:string-value="total"><text:p>total</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell office:value-type="string" office:string-value="apples"><text:p>apples</text:p></table:table-cell>
    <table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>
    <table:table-cell table:formula="of:=[.B2]*2" office:value-type="float" office:value="6"><text:p>6</text:p></table:table-cell>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>`;

const ods = zipSync({
  mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
  "content.xml": strToU8(content),
  "META-INF/manifest.xml": strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`,
  ),
});
writeFileSync(join(here, "fixtures", "sample.ods"), ods);

console.log("wrote sample.xlsx and sample.ods");
