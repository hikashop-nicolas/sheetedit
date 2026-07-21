// Builds demo/pq-sales.xlsx: a minimal workbook with a Sales table (B2:D5), an Output
// table (F2:H3, one stale row) and a Power Query DataMashup payload whose "Output" query
// reads Sales, filters/renames/adds a Total and sorts. Used by the Power Query panel demo
// and imported by tables.test.ts, so the fixture and the test share one builder.
// Synthetic-but-per-spec (MS-QDEFF framing); a real Excel-authored file validates alongside
// when available.
import { strToU8, unzipSync, zipSync } from "fflate";

export const SECTION_M = `section Section1;

shared Output = let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    Typed = Table.TransformColumnTypes(Source, {{"Qty", type number}, {"Price", type number}}),
    Filtered = Table.SelectRows(Typed, each [Qty] > 5),
    Renamed = Table.RenameColumns(Filtered, {{"Qty", "Quantity"}}),
    WithTotal = Table.AddColumn(Renamed, "Total", each [Quantity] * [Price]),
    Slim = Table.RemoveColumns(WithTotal, {"Price"}),
    Sorted = Table.Sort(Slim, {{"Total", 1}})
in
    Sorted;
`;

function buildDataMashup(sectionM) {
  const pkg = zipSync({
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="text/xml"/><Default Extension="m" ContentType="application/x-ms-m"/></Types>',
    ),
    "Config/Package.xml": strToU8(
      '<?xml version="1.0" encoding="utf-8"?><Package xmlns="http://schemas.microsoft.com/DataMashup/Package"><Version>2.0</Version><MinVersion>1.0</MinVersion><Culture>en-US</Culture></Package>',
    ),
    "Formulas/Section1.m": strToU8(sectionM),
  });
  const permissions = strToU8(
    '<?xml version="1.0" encoding="utf-8"?><PermissionList xmlns:xsd="http://www.w3.org/2001/XMLSchema"><CanEvaluateFuturePackages>false</CanEvaluateFuturePackages><FirewallEnabled>true</FirewallEnabled></PermissionList>',
  );
  const metadata = new Uint8Array(0);
  const bindings = new Uint8Array(0);
  const out = new Uint8Array(4 + 4 + pkg.length + 4 + permissions.length + 4 + metadata.length + 4 + bindings.length);
  const view = new DataView(out.buffer);
  let off = 0;
  const u32 = (v) => {
    view.setUint32(off, v, true);
    off += 4;
  };
  const block = (b) => {
    u32(b.length);
    out.set(b, off);
    off += b.length;
  };
  u32(0);
  block(pkg);
  block(permissions);
  block(metadata);
  block(bindings);
  return out;
}

const b64 = (bytes) => {
  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin);
};

const CT =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
  '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
  '<Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
  "</Types>";

const ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
  "</Relationships>";

const WORKBOOK =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>';

const WORKBOOK_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
  "</Relationships>";

const s = (t) => `<is><t>${t}</t></is>`;
const SHEET =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  "<sheetData>" +
  `<row r="2"><c r="B2" t="inlineStr">${s("Product")}</c><c r="C2" t="inlineStr">${s("Qty")}</c><c r="D2" t="inlineStr">${s("Price")}</c>` +
  `<c r="F2" t="inlineStr">${s("Product")}</c><c r="G2" t="inlineStr">${s("Quantity")}</c><c r="H2" t="inlineStr">${s("Total")}</c></row>` +
  `<row r="3"><c r="B3" t="inlineStr">${s("Apples")}</c><c r="C3"><v>10</v></c><c r="D3"><v>2.5</v></c>` +
  `<c r="F3" t="inlineStr">${s("stale")}</c><c r="G3"><v>0</v></c><c r="H3"><v>0</v></c></row>` +
  `<row r="4"><c r="B4" t="inlineStr">${s("Pears")}</c><c r="C4"><v>4</v></c><c r="D4"><v>3</v></c></row>` +
  `<row r="5"><c r="B5" t="inlineStr">${s("Cherries")}</c><c r="C5"><v>20</v></c><c r="D5"><v>5</v></c></row>` +
  "</sheetData>" +
  '<tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts></worksheet>';

const SHEET_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>' +
  "</Relationships>";

const tableXml = (id, name, ref, cols) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${name}" displayName="${name}" ref="${ref}">` +
  `<autoFilter ref="${ref}"/>` +
  `<tableColumns count="${cols.length}">` +
  cols.map((c, i) => `<tableColumn id="${i + 1}" name="${c}"/>`).join("") +
  "</tableColumns>" +
  '<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>';

export function buildPqXlsx() {
  const mashup = buildDataMashup(SECTION_M);
  const itemXml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${b64(mashup)}</DataMashup>`;
  return zipSync({
    "[Content_Types].xml": strToU8(CT),
    "_rels/.rels": strToU8(ROOT_RELS),
    "xl/workbook.xml": strToU8(WORKBOOK),
    "xl/_rels/workbook.xml.rels": strToU8(WORKBOOK_RELS),
    "xl/worksheets/sheet1.xml": strToU8(SHEET),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(SHEET_RELS),
    "xl/tables/table1.xml": strToU8(tableXml(1, "Sales", "B2:D5", ["Product", "Qty", "Price"])),
    "xl/tables/table2.xml": strToU8(tableXml(2, "Output", "F2:H3", ["Product", "Quantity", "Total"])),
    "customXml/item1.xml": strToU8(itemXml),
  });
}

// A variant whose Output query fetches a CSV over the web instead of reading Sales, to
// exercise the browser Web.Contents connector. Output columns match web-sales.csv.
const WEB_SECTION_M = `section Section1;

shared Output = let
    Source = Table.PromoteHeaders(Csv.Document(Web.Contents("/web-sales.csv"))),
    Typed = Table.TransformColumnTypes(Source, {{"Quantity", type number}, {"Total", type number}}),
    Sorted = Table.Sort(Typed, {{"Total", 1}})
in
    Sorted;
`;

export function buildWebPqXlsx() {
  const mashup = buildDataMashup(WEB_SECTION_M);
  const itemXml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${b64(mashup)}</DataMashup>`;
  return zipSync({
    "[Content_Types].xml": strToU8(CT),
    "_rels/.rels": strToU8(ROOT_RELS),
    "xl/workbook.xml": strToU8(WORKBOOK),
    "xl/_rels/workbook.xml.rels": strToU8(WORKBOOK_RELS),
    "xl/worksheets/sheet1.xml": strToU8(SHEET),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(SHEET_RELS),
    "xl/tables/table1.xml": strToU8(tableXml(1, "Sales", "B2:D5", ["Product", "Qty", "Price"])),
    "xl/tables/table2.xml": strToU8(tableXml(2, "Output", "F2:H3", ["Product", "Quantity", "Total"])),
    "customXml/item1.xml": strToU8(itemXml),
  });
}

// An OData.Feed variant: Output = OData.Feed("/odata-sales.json") expanded to a table.
const ODATA_SECTION_M = `section Section1;

shared Output = let
    Source = OData.Feed("/odata-sales.json"),
    Typed = Table.TransformColumnTypes(Source, {{"Quantity", type number}, {"Total", type number}}),
    Sorted = Table.Sort(Typed, {{"Total", 1}})
in
    Sorted;
`;

export function buildODataPqXlsx() {
  const mashup = buildDataMashup(ODATA_SECTION_M);
  const itemXml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${b64(mashup)}</DataMashup>`;
  return zipSync({
    "[Content_Types].xml": strToU8(CT),
    "_rels/.rels": strToU8(ROOT_RELS),
    "xl/workbook.xml": strToU8(WORKBOOK),
    "xl/_rels/workbook.xml.rels": strToU8(WORKBOOK_RELS),
    "xl/worksheets/sheet1.xml": strToU8(SHEET),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(SHEET_RELS),
    "xl/tables/table1.xml": strToU8(tableXml(1, "Sales", "B2:D5", ["Product", "Qty", "Price"])),
    "xl/tables/table2.xml": strToU8(tableXml(2, "Output", "F2:H3", ["Product", "Quantity", "Total"])),
    "customXml/item1.xml": strToU8(itemXml),
  });
}

// Same as pq-sales, but with a connection flagged "Refresh data when opening the file", so
// the Output table auto-populates on open without touching the panel.
const CONNECTIONS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<connection id="1" name="Query - Output" type="5" refreshedVersion="6" background="1" refreshOnLoad="1">' +
  '<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Output" command="SELECT * FROM [Output]"/>' +
  "</connection></connections>";

export function buildAutoRefreshPqXlsx() {
  const base = buildPqXlsx();
  const entries = unzipSync(base);
  entries["xl/connections.xml"] = strToU8(CONNECTIONS);
  return zipSync(entries);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const { writeFileSync } = await import("node:fs");
  writeFileSync(new URL("../demo/pq-sales.xlsx", import.meta.url), buildPqXlsx());
  writeFileSync(new URL("../demo/pq-autorefresh.xlsx", import.meta.url), buildAutoRefreshPqXlsx());
  writeFileSync(new URL("../demo/pq-web.xlsx", import.meta.url), buildWebPqXlsx());
  writeFileSync(new URL("../demo/web-sales.csv", import.meta.url), "Product,Quantity,Total\nCherries,20,100\nApples,10,25\n");
  writeFileSync(new URL("../demo/pq-odata.xlsx", import.meta.url), buildODataPqXlsx());
  writeFileSync(
    new URL("../demo/odata-sales.json", import.meta.url),
    JSON.stringify({ "@odata.context": "$metadata#Sales", value: [{ Product: "Cherries", Quantity: 20, Total: 100 }, { Product: "Apples", Quantity: 10, Total: 25 }] }, null, 2),
  );
  console.log("wrote demo/pq-sales.xlsx, demo/pq-web.xlsx, demo/web-sales.csv, demo/pq-odata.xlsx, demo/odata-sales.json");
}
