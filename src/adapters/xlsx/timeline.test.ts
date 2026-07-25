import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { writeWorkbook } from "../../core/workbook";
import { timelinePeriods } from "../../core/ui/timeline-layer";

const X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const TS = "http://schemas.microsoft.com/office/drawing/2012/timeslicer";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

/**
 * demo/c-pivot.xlsx has PivotTable1 on the "Pivot" sheet. Add the parts Excel writes for a
 * timeline: the view part, the cache part carrying the selected range and the data bounds, and a
 * drawing graphicFrame carrying the timeslicer extension for the anchor.
 */
function withTimeline(sel?: { start: string; end: string }, level = "2"): Uint8Array {
  const files = unzipSync(new Uint8Array(readFileSync("demo/c-pivot.xlsx")));
  const selection = sel ? `<x15:selection startDate="${sel.start}" endDate="${sel.end}"/>` : "";
  files["xl/timelineCaches/timelineCache1.xml"] = strToU8(
    `<x15:timelineCacheDefinition xmlns:x15="${X15}" name="NativeTimeline_Date" sourceName="Date">` +
      `<x15:pivotTables><x15:pivotTable tabId="1" name="PivotTable1"/></x15:pivotTables>` +
      `<x15:state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="1" filterType="dateBetween">` +
      selection +
      `<x15:bounds startDate="2024-01-01T00:00:00" endDate="2024-05-01T00:00:00"/>` +
      `</x15:state></x15:timelineCacheDefinition>`);
  files["xl/timelines/timeline1.xml"] = strToU8(
    `<x15:timelines xmlns:x15="${X15}">` +
      `<x15:timeline name="Date" cache="NativeTimeline_Date" caption="Date" level="${level}" selectionLevel="${level}"/>` +
      `</x15:timelines>`);
  files["xl/drawings/drawing9.xml"] = strToU8(
    `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}">` +
      `<xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
      `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="9" name="Date"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="${TS}"><tsle:timeslicer xmlns:tsle="${TS}" name="Date"/></a:graphicData></a:graphic>` +
      `</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
  const rels = strFromU8(files["xl/worksheets/_rels/sheet2.xml.rels"]!)
    .replace("</Relationships>",
      `<Relationship Id="rId60" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/>` +
      `<Relationship Id="rId61" Type="${R}/drawing" Target="../drawings/drawing9.xml"/></Relationships>`);
  files["xl/worksheets/_rels/sheet2.xml.rels"] = strToU8(rels);
  return zipSync(files);
}

const cacheOf = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["xl/timelineCaches/timelineCache1.xml"]!);

describe("timelines", () => {
  it("reads the view, the cache state and the drawing anchor", () => {
    const wb = readWorkbook(withTimeline({ start: "2024-02-01T00:00:00", end: "2024-03-01T00:00:00" }));
    const tl = wb.sheets.flatMap((s) => s.timelines ?? [])[0];
    expect(tl).toBeTruthy();
    expect(tl!.name).toBe("Date");
    expect(tl!.cache).toBe("NativeTimeline_Date");
    expect(tl!.sourceName).toBe("Date");
    expect(tl!.pivotTables).toEqual(["PivotTable1"]);
    expect(tl!.startDate).toBe("2024-02-01T00:00:00");
    expect(tl!.endDate).toBe("2024-03-01T00:00:00");
    expect(tl!.boundStart).toBe("2024-01-01T00:00:00");
    expect(tl!.boundEnd).toBe("2024-05-01T00:00:00");
    // anchored at E2 (col 4 zero-based -> 5)
    expect(tl!.anchor?.fromCol).toBe(5);
    expect(tl!.anchor?.fromRow).toBe(2);
  });

  it("reads a timeline with no selection as unfiltered", () => {
    const wb = readWorkbook(withTimeline());
    const tl = wb.sheets.flatMap((s) => s.timelines ?? [])[0]!;
    expect(tl.startDate).toBeUndefined();
    expect(tl.endDate).toBeUndefined();
  });

  it("leaves the cache part untouched when nothing changed", () => {
    const src = withTimeline({ start: "2024-02-01T00:00:00", end: "2024-03-01T00:00:00" });
    const wb = readWorkbook(src);
    const out = writeWorkbook(wb);
    expect(cacheOf(out)).toBe(cacheOf(src));
  });

  it("writes a changed range into the cache selection", () => {
    const wb = readWorkbook(withTimeline({ start: "2024-02-01T00:00:00", end: "2024-03-01T00:00:00" }));
    const tl = wb.sheets.flatMap((s) => s.timelines ?? [])[0]!;
    tl.startDate = "2024-03-01T00:00:00";
    tl.endDate = "2024-04-01T00:00:00";
    tl.dirty = true;
    const xml = cacheOf(writeWorkbook(wb));
    expect(xml).toContain('startDate="2024-03-01T00:00:00"');
    expect(xml).toContain('endDate="2024-04-01T00:00:00"');
    expect(xml).not.toContain('startDate="2024-02-01T00:00:00"');
    // The bounds are untouched.
    expect(xml).toContain('endDate="2024-05-01T00:00:00"');
  });

  it("adds a selection element to a timeline that had none", () => {
    const wb = readWorkbook(withTimeline());
    const tl = wb.sheets.flatMap((s) => s.timelines ?? [])[0]!;
    tl.startDate = "2024-02-01T00:00:00";
    tl.endDate = "2024-03-01T00:00:00";
    tl.dirty = true;
    const xml = cacheOf(writeWorkbook(wb));
    expect(xml).toMatch(/selection[^>]*startDate="2024-02-01T00:00:00"/);
    expect(xml).toMatch(/selection[^>]*endDate="2024-03-01T00:00:00"/);
  });

  it("removes the selection when the range is cleared", () => {
    const wb = readWorkbook(withTimeline({ start: "2024-02-01T00:00:00", end: "2024-03-01T00:00:00" }));
    const tl = wb.sheets.flatMap((s) => s.timelines ?? [])[0]!;
    tl.startDate = undefined;
    tl.endDate = undefined;
    tl.dirty = true;
    const xml = cacheOf(writeWorkbook(wb));
    expect(xml).not.toContain("selection");
    expect(xml).toContain("bounds");
  });

  it("round-trips a rewritten range", () => {
    const wb = readWorkbook(withTimeline({ start: "2024-02-01T00:00:00", end: "2024-03-01T00:00:00" }));
    const tl = wb.sheets.flatMap((s) => s.timelines ?? [])[0]!;
    tl.startDate = "2024-04-01T00:00:00";
    tl.endDate = "2024-05-01T00:00:00";
    tl.dirty = true;
    const again = readWorkbook(writeWorkbook(wb));
    const tl2 = again.sheets.flatMap((s) => s.timelines ?? [])[0]!;
    expect(tl2.startDate).toBe("2024-04-01T00:00:00");
    expect(tl2.endDate).toBe("2024-05-01T00:00:00");
  });
});

describe("timeline periods", () => {
  const base = { name: "Date", cache: "c", sourceName: "Date", pivotTables: [], boundStart: "2024-01-01T00:00:00", boundEnd: "2024-05-01T00:00:00" };

  it("splits the bounds into months by default", () => {
    const ps = timelinePeriods({ ...base, level: "2" });
    expect(ps.map((p) => p.label)).toEqual(["Jan 24", "Feb 24", "Mar 24", "Apr 24", "May 24"]);
    expect(ps[0]!.start).toBe("2024-01-01T00:00:00");
    expect(ps[0]!.end).toBe("2024-02-01T00:00:00");
  });

  it("splits into years at level 0", () => {
    const ps = timelinePeriods({ ...base, level: "0", boundEnd: "2026-05-01T00:00:00" });
    expect(ps.map((p) => p.label)).toEqual(["2024", "2025", "2026"]);
    expect(ps[1]!.start).toBe("2025-01-01T00:00:00");
  });

  it("splits into quarters at level 1", () => {
    const ps = timelinePeriods({ ...base, level: "1" });
    expect(ps.map((p) => p.label)).toEqual(["Q1 2024", "Q2 2024"]);
    expect(ps[0]!.end).toBe("2024-04-01T00:00:00");
  });

  it("splits into days at level 3", () => {
    const ps = timelinePeriods({ ...base, level: "3", boundStart: "2024-01-30T00:00:00", boundEnd: "2024-02-02T00:00:00" });
    expect(ps.map((p) => p.label)).toEqual(["30/1", "31/1", "1/2", "2/2"]);
    expect(ps[1]!.end).toBe("2024-02-01T00:00:00");
  });

  it("returns nothing when the bounds are unknown", () => {
    expect(timelinePeriods({ name: "x", cache: "c", sourceName: "d", pivotTables: [] })).toEqual([]);
  });
});
