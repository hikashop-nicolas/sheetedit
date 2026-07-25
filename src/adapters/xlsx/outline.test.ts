import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../core/workbook";
import { groupLines, setGroupCollapsed, ungroupLines } from "../../core/outline";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** A minimal workbook whose sheet body can carry outline attributes. */
function book(sheetInner: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}">${sheetInner}</worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${MAIN}"/>`),
  });
}

const rows = (spec: { r: number; attrs?: string }[]): string =>
  `<sheetData>${spec.map((s) => `<row r="${s.r}"${s.attrs ?? ""}><c r="A${s.r}" t="inlineStr"><is><t>v${s.r}</t></is></c></row>`).join("")}</sheetData>`;

const sheetXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["xl/worksheets/sheet1.xml"]!);

describe("xlsx outline grouping", () => {
  it("reads row outline levels, collapsed markers and the summary side", () => {
    const wb = readWorkbook(book(
      `<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>` +
      rows([{ r: 1 }, { r: 2, attrs: ` outlineLevel="1" hidden="1"` }, { r: 3, attrs: ` outlineLevel="2" hidden="1"` }, { r: 4, attrs: ` collapsed="1"` }])));
    const s = wb.sheets[0]!;
    expect([...s.rowOutline!.entries()]).toEqual([[2, 1], [3, 2]]);
    expect([...s.rowCollapsed!]).toEqual([4]);
    expect([...s.hiddenRows!].sort((a, b) => a - b)).toEqual([2, 3]);
    expect(s.summaryBelow).toBe(false);
    expect(s.summaryRight).toBe(false);
  });

  it("defaults the summary side to below/right when outlinePr is absent", () => {
    const wb = readWorkbook(book(rows([{ r: 1 }])));
    expect(wb.sheets[0]!.summaryBelow).toBeUndefined(); // absent means Excel's default
  });

  it("reads column outline levels from the cols spans", () => {
    const wb = readWorkbook(book(`<cols><col min="2" max="4" outlineLevel="1" hidden="1"/><col min="5" max="5" width="12"/></cols>` + rows([{ r: 1 }])));
    const s = wb.sheets[0]!;
    expect([...s.colOutline!.entries()]).toEqual([[2, 1], [3, 1], [4, 1]]);
    expect([...s.hiddenCols!].sort((a, b) => a - b)).toEqual([2, 3, 4]);
  });

  it("leaves the sheet untouched when nothing was grouped", () => {
    const src = book(rows([{ r: 1 }, { r: 2 }]));
    expect(sheetXml(writeWorkbook(readWorkbook(src)))).toBe(sheetXml(src));
  });

  it("writes outlineLevel onto the grouped rows", () => {
    const wb = readWorkbook(book(rows([{ r: 1 }, { r: 2 }, { r: 3 }, { r: 4 }])));
    groupLines(wb.sheets[0]!, "row", 2, 3);
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/<row r="2"[^>]*outlineLevel="1"/);
    expect(xml).toMatch(/<row r="3"[^>]*outlineLevel="1"/);
    expect(xml).not.toMatch(/<row r="4"[^>]*outlineLevel/);
  });

  it("writes hidden + collapsed when a group is collapsed", () => {
    const wb = readWorkbook(book(rows([{ r: 1 }, { r: 2 }, { r: 3 }, { r: 4 }])));
    const s = wb.sheets[0]!;
    groupLines(s, "row", 2, 3);
    setGroupCollapsed(s, "row", 2, 1, true, 4);
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/<row r="2"[^>]*hidden="1"/);
    expect(xml).toMatch(/<row r="3"[^>]*hidden="1"/);
    // The summary row is the one just below the group.
    expect(xml).toMatch(/<row r="4"[^>]*collapsed="1"/);
  });

  it("clears the attributes again when the group is removed", () => {
    const wb = readWorkbook(book(rows([{ r: 1 }, { r: 2, attrs: ` outlineLevel="1" hidden="1"` }, { r: 3, attrs: ` collapsed="1"` }])));
    const s = wb.sheets[0]!;
    ungroupLines(s, "row", 2, 2);
    s.rowCollapsed?.delete(3);
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).not.toContain("outlineLevel");
    expect(xml).not.toContain("hidden");
    expect(xml).not.toContain("collapsed");
  });

  it("writes column groups as per-column cols spans, keeping widths", () => {
    const wb = readWorkbook(book(`<cols><col min="1" max="5" width="12"/></cols>` + rows([{ r: 1 }])));
    const s = wb.sheets[0]!;
    groupLines(s, "col", 2, 3);
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/<col min="2" max="2"[^>]*width="12"[^>]*outlineLevel="1"/);
    expect(xml).toMatch(/<col min="1" max="1"[^>]*width="12"/);
    expect(xml).not.toMatch(/<col min="1" max="1"[^>]*outlineLevel/);
  });

  it("round-trips a grouped, collapsed sheet", () => {
    const wb = readWorkbook(book(rows([{ r: 1 }, { r: 2 }, { r: 3 }, { r: 4 }, { r: 5 }])));
    const s = wb.sheets[0]!;
    groupLines(s, "row", 2, 4);
    groupLines(s, "row", 3, 3);
    setGroupCollapsed(s, "row", 3, 2, true, 5);
    const again = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect([...again.rowOutline!.entries()]).toEqual([[2, 1], [3, 2], [4, 1]]);
    expect([...again.hiddenRows!]).toEqual([3]);
    expect([...again.rowCollapsed!]).toEqual([4]);
  });

  it("writes outlinePr only when the summary side is not the default", () => {
    const wb = readWorkbook(book(rows([{ r: 1 }, { r: 2 }])));
    const s = wb.sheets[0]!;
    groupLines(s, "row", 2, 2);
    expect(sheetXml(writeWorkbook(wb))).not.toContain("outlinePr");
    s.summaryBelow = false;
    s.outlineDirty = true;
    expect(sheetXml(writeWorkbook(wb))).toMatch(/<outlinePr summaryBelow="0"/);
  });
});
