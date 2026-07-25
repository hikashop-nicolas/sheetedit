import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { writeWorkbook } from "../../core/workbook";
import { createXlsxSlicer } from "./slicer-create";

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
