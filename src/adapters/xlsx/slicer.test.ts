import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { writeWorkbook } from "../../core/workbook";
import { createXlsxSlicer, createXlsxTableSlicer } from "./slicer-create";
import { listWorkbookTables } from "./tables";

const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SLE = "http://schemas.microsoft.com/office/drawing/2010/slicer";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

/**
 * demo/c-pivot.xlsx has a real pivot (PivotTable1 on the "Pivot" sheet) whose cache field "Region"
 * has shared items North / South. Add the three slicer parts Excel would write:
 * the slicer view, the slicer cache (with the selection), and a drawing graphicFrame for the anchor.
 */
function withSlicer(selected: number[] = [0, 1]): Uint8Array {
  const files = unzipSync(new Uint8Array(readFileSync("demo/c-pivot.xlsx")));
  const items = [0, 1].map((x) => `<x14:i x="${x}"${selected.includes(x) ? ' s="1"' : ""}/>`).join("");
  files["xl/slicerCaches/slicerCache1.xml"] = strToU8(
    `<x14:slicerCacheDefinition xmlns:x14="${X14}" name="Slicer_Region" sourceName="Region">` +
      `<x14:pivotTables><x14:pivotTable tabId="1" name="PivotTable1"/></x14:pivotTables>` +
      `<x14:data><x14:tabular pivotCacheId="1"><x14:items count="2">${items}</x14:items></x14:tabular></x14:data>` +
      `</x14:slicerCacheDefinition>`);
  files["xl/slicers/slicer1.xml"] = strToU8(
    `<x14:slicers xmlns:x14="${X14}">` +
      `<x14:slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="1" rowHeight="234950"/>` +
      `</x14:slicers>`);
  // The pivot lives on sheet2; anchor the slicer there via a drawing graphicFrame.
  files["xl/drawings/drawing9.xml"] = strToU8(
    `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}">` +
      `<xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
      `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="9" name="Region"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="${SLE}"><sle:slicer xmlns:sle="${SLE}" name="Region"/></a:graphicData></a:graphic>` +
      `</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
  // Point sheet2 at both the slicer part and the drawing.
  const rels = strFromU8(files["xl/worksheets/_rels/sheet2.xml.rels"]!)
    .replace("</Relationships>",
      `<Relationship Id="rId50" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>` +
      `<Relationship Id="rId51" Type="${R}/drawing" Target="../drawings/drawing9.xml"/></Relationships>`);
  files["xl/worksheets/_rels/sheet2.xml.rels"] = strToU8(rels);
  return zipSync(files);
}

const cacheOf = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["xl/slicerCaches/slicerCache1.xml"]!);

describe("slicers", () => {
  it("reads the view, the cache selection and the drawing anchor", () => {
    const wb = readWorkbook(withSlicer([0, 1]));
    const sl = wb.sheets.flatMap((s) => s.slicers ?? [])[0];
    expect(sl).toBeTruthy();
    expect(sl!.name).toBe("Region");
    expect(sl!.cache).toBe("Slicer_Region");
    expect(sl!.sourceName).toBe("Region");
    expect(sl!.pivotTables).toEqual(["PivotTable1"]);
    // Item labels come from the pivot cache field's sharedItems.
    expect(sl!.items.map((i) => i.label)).toEqual(["North", "South"]);
    expect(sl!.items.map((i) => i.selected)).toEqual([true, true]);
    // anchored at E2 (col 4 zero-based -> 5) .. H9
    expect(sl!.anchor?.fromCol).toBe(5);
    expect(sl!.anchor?.fromRow).toBe(2);
  });

  it("reads a partial selection", () => {
    const wb = readWorkbook(withSlicer([1])); // only South
    const sl = wb.sheets.flatMap((s) => s.slicers ?? [])[0]!;
    expect(sl.items.map((i) => i.selected)).toEqual([false, true]);
  });

  it("treats a cache with nothing selected as no filter, not everything hidden", () => {
    const wb = readWorkbook(withSlicer([]));
    const sl = wb.sheets.flatMap((s) => s.slicers ?? [])[0]!;
    expect(sl.items.every((i) => i.selected)).toBe(true);
  });

  it("writes a changed selection back into the cache part", () => {
    const wb = readWorkbook(withSlicer([0, 1]));
    const sl = wb.sheets.flatMap((s) => s.slicers ?? [])[0]!;
    sl.items[0]!.selected = false; // deselect North
    sl.dirty = true;
    const out = writeWorkbook(wb);
    const xml = cacheOf(out);
    expect(xml).toMatch(/<x14:i x="0"\s*\/>/);        // North: no s attribute
    expect(xml).toMatch(/<x14:i x="1" s="1"\s*\/>/);  // South: selected
    // and it re-reads
    const re = readWorkbook(out).sheets.flatMap((s) => s.slicers ?? [])[0]!;
    expect(re.items.map((i) => i.selected)).toEqual([false, true]);
  });

  it("leaves an untouched slicer's cache part byte-identical", () => {
    const src = withSlicer([0, 1]);
    const before = cacheOf(src);
    const wb = readWorkbook(src);
    expect(cacheOf(writeWorkbook(wb))).toBe(before);
  });
});

describe("creating a slicer from scratch", () => {
  /** Build a slicer on demo/c-pivot.xlsx's PivotTable1 for the Region field. */
  function created(): { wb: ReturnType<typeof readWorkbook>; out: Uint8Array; files: Record<string, Uint8Array> } {
    const wb = readWorkbook(new Uint8Array(readFileSync("demo/c-pivot.xlsx")));
    const host = wb.sheets.find((s) => (s.pivotTables ?? []).length)!;
    const info = host.pivotTables![0]!;
    const sl = createXlsxSlicer(wb, host, info, "Region", ["North", "South"],
      { fromCol: 6, fromRow: 2, fromColOff: 0, fromRowOff: 0, toCol: 9, toRow: 10, toColOff: 0, toRowOff: 0 });
    expect(sl).toBeTruthy();
    const out = writeWorkbook(wb);
    return { wb, out, files: unzipSync(out) };
  }

  it("writes the cache part with every item selected", () => {
    const { files } = created();
    const cache = strFromU8(files["xl/slicerCaches/slicerCache1.xml"]!);
    expect(cache).toContain('name="Slicer_Region"');
    expect(cache).toContain('sourceName="Region"');
    expect(cache).toContain('<x14:pivotTable tabId="1" name="PivotTable1"/>');
    expect(cache).toMatch(/<x14:i x="0" s="1"\/><x14:i x="1" s="1"\/>/);
    expect(cache).toContain(`xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"`);
  });

  it("writes the view part and registers both content types", () => {
    const { files } = created();
    const view = strFromU8(files["xl/slicers/slicer1.xml"]!);
    expect(view).toContain('<x14:slicer name="Region" cache="Slicer_Region"');
    expect(view).toContain('rowHeight='); // required by CT_Slicer
    const ct = strFromU8(files["[Content_Types].xml"]!);
    expect(ct).toContain("application/vnd.ms-excel.slicer+xml");
    expect(ct).toContain("application/vnd.ms-excel.slicerCache+xml");
  });

  it("registers the cache on the workbook and the view on the worksheet", () => {
    const { files } = created();
    // workbook rel + extLst
    const wbRels = strFromU8(files["xl/_rels/workbook.xml.rels"]!);
    expect(wbRels).toContain("http://schemas.microsoft.com/office/2007/relationships/slicerCache");
    const wbXml = strFromU8(files["xl/workbook.xml"]!);
    expect(wbXml).toContain("{BBE1A952-AA13-448e-AADC-164F8A28A991}");
    expect(wbXml).toMatch(/<x14:slicerCaches[\s\S]*<x14:slicerCache r:id="rId\d+"/);
    // worksheet rel + extLst
    const shRels = Object.keys(files).filter((f) => /worksheets\/_rels\//.test(f)).map((f) => strFromU8(files[f]!)).join("");
    expect(shRels).toContain("http://schemas.microsoft.com/office/2007/relationships/slicer");
    const sheets = Object.keys(files).filter((f) => /worksheets\/sheet\d+\.xml$/.test(f)).map((f) => strFromU8(files[f]!)).join("");
    expect(sheets).toContain("{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
    expect(sheets).toMatch(/<x14:slicerList[\s\S]*<x14:slicer r:id="rId\d+"/);
  });

  it("anchors it in the drawing via an sle:slicer graphicFrame", () => {
    const { files } = created();
    const drawings = Object.keys(files).filter((f) => /drawings\/drawing\d+\.xml$/.test(f)).map((f) => strFromU8(files[f]!)).join("");
    expect(drawings).toContain("http://schemas.microsoft.com/office/drawing/2010/slicer");
    expect(drawings).toMatch(/<sle:slicer[^>]*name="Region"/);
  });

  it("round-trips: the created slicer reads back with its items and anchor", () => {
    const { out } = created();
    const re = readWorkbook(out).sheets.flatMap((s) => s.slicers ?? [])[0];
    expect(re).toBeTruthy();
    expect(re!.name).toBe("Region");
    expect(re!.sourceName).toBe("Region");
    expect(re!.items.map((i) => i.label)).toEqual(["North", "South"]);
    expect(re!.items.every((i) => i.selected)).toBe(true);
    expect(re!.anchor?.fromCol).toBe(6);
  });
});

const X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const EXT_TSC = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}";

describe("slicer kinds", () => {
  /** Swap in a cache of a given flavour on the c-pivot fixture. */
  function withCache(cacheXmlInner: string): Uint8Array {
    const files = unzipSync(withSlicer([0, 1]));
    files["xl/slicerCaches/slicerCache1.xml"] = strToU8(
      `<x14:slicerCacheDefinition xmlns:x14="${X14}" xmlns:x15="${X15}" name="Slicer_Region" sourceName="Region">${cacheXmlInner}</x14:slicerCacheDefinition>`);
    return zipSync(files);
  }
  const first = (b: Uint8Array) => readWorkbook(b).sheets.flatMap((s) => s.slicers ?? [])[0]!;

  it("reads a table slicer (x15:tableSlicerCache) and binds it to the table column", () => {
    // c-pivot.xlsx has no table part, so the binding falls back to no table; assert the KIND and
    // that the cache's tableId/column were parsed rather than mistaken for a pivot cache.
    const sl = first(withCache(
      `<x14:data><x14:tabular pivotCacheId="1"><x14:items count="2"><x14:i x="0" s="1"/><x14:i x="1" s="1"/></x14:items></x14:tabular></x14:data>` +
      `<x14:extLst><x14:ext uri="${EXT_TSC}"><x15:tableSlicerCache tableId="1" column="1"/></x14:ext></x14:extLst>`));
    expect(sl.kind).toBe("table");
  });

  it("reads an OLAP slicer's captions and marks it read-only", () => {
    const sl = first(withCache(
      `<x14:data><x14:olap pivotCacheId="1"><x14:levels count="1"><x14:level uniqueName="[Geo].[Region]" sourceCaption="Region" count="2">` +
      `<x14:ranges><x14:range startItem="0">` +
      `<x14:i n="[Geo].[Region].&amp;[North]" c="North"/><x14:i n="[Geo].[Region].&amp;[South]" c="South"/>` +
      `</x14:range></x14:ranges></x14:level></x14:levels>` +
      `<x14:selections count="1"><x14:selection n="North"/></x14:selections>` +
      `</x14:olap></x14:data>`));
    expect(sl.kind).toBe("olap");
    expect(sl.items.map((i) => i.label)).toEqual(["North", "South"]);
    // only the selected caption is on
    expect(sl.items.map((i) => i.selected)).toEqual([true, false]);
  });

  it("reads the slicer style name", () => {
    const files = unzipSync(withSlicer([0, 1]));
    files["xl/slicers/slicer1.xml"] = strToU8(
      `<x14:slicers xmlns:x14="${X14}"><x14:slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="1" style="SlicerStyleDark3" rowHeight="234950"/></x14:slicers>`);
    expect(first(zipSync(files)).style).toBe("SlicerStyleDark3");
  });
});

describe("slicer style accents", () => {
  it("maps the built-in style families to an accent colour", async () => {
    const { styleAccent } = await import("../../core/ui/slicer-layer");
    expect(styleAccent(undefined)).toBeUndefined();
    expect(styleAccent("NotAStyle")).toBeUndefined();
    expect(styleAccent("SlicerStyleLight1")).toBe("#7f7f7f");   // the neutral one
    expect(styleAccent("SlicerStyleLight2")).toBe("#ed7d31");   // accent2
    // Dark uses a deeper tone of the same accent than Light.
    expect(styleAccent("SlicerStyleDark2")).not.toBe(styleAccent("SlicerStyleLight2"));
    expect(styleAccent("SlicerStyleDark2")).toMatch(/^#[0-9a-f]{6}$/);
  });
});

describe("table slicer bound to a real table", () => {
  const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";
  /** A one-sheet workbook with a ListObject over A1:B5 and a table slicer on its 2nd column. */
  function tableBook(): Uint8Array {
    const rows = [["Item", "Region"], ["a", "North"], ["b", "South"], ["c", "North"], ["d", "West"]]
      .map((cells, i) => `<row r="${i + 1}">${cells.map((v, c) => `<c r="${String.fromCharCode(65 + c)}${i + 1}" t="inlineStr"><is><t>${v}</t></is></c>`).join("")}</row>`).join("");
    return zipSync({
      "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
      "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`),
      "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
      "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`),
      "xl/worksheets/_rels/sheet1.xml.rels": strToU8(`<Relationships xmlns="${RELNS}">` +
        `<Relationship Id="rId1" Type="${R}/table" Target="../tables/table1.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>` +
        `</Relationships>`),
      "xl/tables/table1.xml": strToU8(`<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Tbl" displayName="Tbl" ref="A1:B5" headerRowCount="1"><tableColumns count="2"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Region"/></tableColumns></table>`),
      "xl/slicers/slicer1.xml": strToU8(`<x14:slicers xmlns:x14="${X14}"><x14:slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="1" rowHeight="234950"/></x14:slicers>`),
      "xl/slicerCaches/slicerCache1.xml": strToU8(
        `<x14:slicerCacheDefinition xmlns:x14="${X14}" xmlns:x15="${X15}" name="Slicer_Region" sourceName="Region">` +
        `<x14:data><x14:tabular><x14:items count="3"><x14:i x="0" s="1"/><x14:i x="1" s="1"/><x14:i x="2" s="1"/></x14:items></x14:tabular></x14:data>` +
        `<x14:extLst><x14:ext uri="${EXT_TSC}"><x15:tableSlicerCache tableId="1" column="2"/></x14:ext></x14:extLst>` +
        `</x14:slicerCacheDefinition>`),
      "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
    });
  }

  it("takes its items from the bound table column's distinct values", () => {
    const sl = readWorkbook(tableBook()).sheets.flatMap((s) => s.slicers ?? [])[0]!;
    expect(sl.kind).toBe("table");
    // column id 2 -> the 2nd table column ("Region"), whose body holds North/South/North/West
    expect(sl.table).toMatchObject({ sheetIndex: 0, r1: 1, c1: 1, r2: 5, c2: 2, headerRows: 1, col: 1 });
    expect(sl.items.map((i) => i.label)).toEqual(["North", "South", "West"]); // distinct, first-seen order
    expect(sl.items.every((i) => i.selected)).toBe(true);
  });
});

const RELNS2 = "http://schemas.openxmlformats.org/package/2006/relationships";

describe("creating a table slicer", () => {
  /** A minimal workbook holding one Excel table and no slicers yet. */
  function bareTableBook(): Uint8Array {
    const rows = [["Item", "Region"], ["a", "North"], ["b", "South"], ["c", "North"], ["d", "West"]]
      .map((cells, i) => `<row r="${i + 1}">${cells.map((v, c) => `<c r="${String.fromCharCode(65 + c)}${i + 1}" t="inlineStr"><is><t>${v}</t></is></c>`).join("")}</row>`).join("");
    return zipSync({
      "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
      "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS2}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`),
      "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS2}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
      "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`),
      "xl/worksheets/_rels/sheet1.xml.rels": strToU8(`<Relationships xmlns="${RELNS2}"><Relationship Id="rId1" Type="${R}/table" Target="../tables/table1.xml"/></Relationships>`),
      "xl/tables/table1.xml": strToU8(`<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="7" name="Tbl" displayName="Tbl" ref="A1:B5" headerRowCount="1"><tableColumns count="2"><tableColumn id="1" name="Item"/><tableColumn id="4" name="Region"/></tableColumns></table>`),
      "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
    });
  }

  const ANCHOR = { fromCol: 5, fromRow: 1, fromColOff: 0, fromRowOff: 0, toCol: 8, toRow: 9, toColOff: 0, toRowOff: 0 };

  /** Create a table slicer on the bare fixture and hand back the saved bytes. */
  function created(): { bytes: Uint8Array; file: (p: string) => string } {
    const wb = readWorkbook(bareTableBook());
    const table = listWorkbookTables(wb)[0]!;
    const sheet = wb.sheets[0]!;
    const sl = createXlsxTableSlicer(wb, sheet, table, 7, 4, 1, "Region", ["North", "South", "West"], ANCHOR);
    expect(sl).toBeTruthy();
    const bytes = writeWorkbook(wb);
    const files = unzipSync(bytes);
    return { bytes, file: (p) => strFromU8(files[p]!) };
  }

  it("writes a cache bound to the table column, not to a pivot", () => {
    const { file } = created();
    const cache = file("xl/slicerCaches/slicerCache1.xml");
    expect(cache).toContain(`<x15:tableSlicerCache tableId="7" column="4"`);
    expect(cache).toContain(`uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"`);
    // A table slicer has no pivot to name and no pivot cache to point at.
    expect(cache).not.toContain("pivotTable");
    expect(cache).not.toContain("pivotCacheId");
    expect(cache).toContain(`<x14:items count="3">`);
  });

  it("registers the cache under the TABLE workbook extension, in an x15 container", () => {
    const { file } = created();
    const wbXml = file("xl/workbook.xml");
    expect(wbXml).toContain("{46BE6895-7355-4a93-B00E-2C351335B9C9}");
    expect(wbXml).not.toContain("{BBE1A952-AA13-448e-AADC-164F8A28A991}");
    expect(wbXml).toMatch(/slicerCaches/);
    // The cache name is also a defined name, the way Excel writes it.
    expect(wbXml).toMatch(/<definedName name="Slicer_Region">#N\/A<\/definedName>/);
  });

  it("registers the view under the TABLE worksheet extension", () => {
    const { file } = created();
    const sheetXml = file("xl/worksheets/sheet1.xml");
    expect(sheetXml).toContain("{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");
    expect(sheetXml).not.toContain("{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
  });

  it("writes both content types, both relationships and the drawing anchor", () => {
    const { file } = created();
    const ct = file("[Content_Types].xml");
    expect(ct).toContain("application/vnd.ms-excel.slicer+xml");
    expect(ct).toContain("application/vnd.ms-excel.slicerCache+xml");
    expect(file("xl/_rels/workbook.xml.rels")).toContain("slicerCaches/slicerCache1.xml");
    expect(file("xl/worksheets/_rels/sheet1.xml.rels")).toContain("../slicers/slicer1.xml");
    const drawing = Object.keys(unzipSync(created().bytes)).find((k) => /^xl\/drawings\/drawing\d+\.xml$/.test(k))!;
    expect(file(drawing)).toContain("sle:slicer");
  });

  it("reads back as a table slicer bound to the right column", () => {
    const again = readWorkbook(created().bytes);
    const sl = again.sheets.flatMap((s) => s.slicers ?? [])[0]!;
    expect(sl.kind).toBe("table");
    // tableColumn id 4 is the SECOND column, so the offset the UI filters on is 1.
    expect(sl.table).toMatchObject({ sheetIndex: 0, r1: 1, c1: 1, r2: 5, c2: 2, headerRows: 1, col: 1 });
    expect(sl.items.map((i) => i.label)).toEqual(["North", "South", "West"]);
    expect(sl.items.every((i) => i.selected)).toBe(true);
  });

  it("still writes a pivot slicer under the pivot extensions", () => {
    const wb = readWorkbook(withSlicer([0, 1]));
    const host = wb.sheets.find((s) => s.pivotTables?.length)!;
    const info = host.pivotTables![0]!;
    createXlsxSlicer(wb, host, info, "Product", ["Apple", "Banana"], ANCHOR);
    const wbXml = strFromU8(unzipSync(writeWorkbook(wb))["xl/workbook.xml"]!);
    expect(wbXml).toContain("{BBE1A952-AA13-448e-AADC-164F8A28A991}");
    expect(wbXml).not.toContain("{46BE6895-7355-4a93-B00E-2C351335B9C9}");
  });
});
