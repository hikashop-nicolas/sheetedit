import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../core/workbook";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";

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

const DATA = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c></row></sheetData>`;
const sheetXml = (b: Uint8Array): string => strFromU8(unzipSync(b)["xl/worksheets/sheet1.xml"]!);

describe("xlsx frozen panes", () => {
  it("reads a frozen pane's row and column counts", () => {
    const wb = readWorkbook(book(`<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>${DATA}`));
    expect(wb.sheets[0]!.freeze).toEqual({ rows: 1, cols: 2 });
  });

  it("ignores a split (not frozen) pane", () => {
    const wb = readWorkbook(book(`<sheetViews><sheetView workbookViewId="0"><pane xSplit="1200" ySplit="600" state="split"/></sheetView></sheetViews>${DATA}`));
    expect(wb.sheets[0]!.freeze).toBeUndefined();
  });

  it("leaves the sheet untouched when the freeze was not changed", () => {
    const src = book(`<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>${DATA}`);
    expect(sheetXml(writeWorkbook(readWorkbook(src)))).toBe(sheetXml(src));
  });

  it("writes a new pane into a sheet that had no sheetViews", () => {
    const wb = readWorkbook(book(DATA));
    const s = wb.sheets[0]!;
    s.freeze = { rows: 1, cols: 0 };
    s.freezeDirty = true;
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/<sheetViews><sheetView workbookViewId="0"><pane[^>]*ySplit="1"/);
    expect(xml).toMatch(/topLeftCell="A2"/);
    expect(xml).toMatch(/activePane="bottomLeft"/);
    expect(xml).toMatch(/state="frozen"/);
    expect(xml).not.toContain("xSplit");
  });

  it("names the bottom-right pane when both axes are frozen", () => {
    const wb = readWorkbook(book(DATA));
    Object.assign(wb.sheets[0]!, { freeze: { rows: 2, cols: 3 }, freezeDirty: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/xSplit="3"/);
    expect(xml).toMatch(/ySplit="2"/);
    expect(xml).toMatch(/topLeftCell="D3"/);
    expect(xml).toMatch(/activePane="bottomRight"/);
  });

  it("updates an existing pane in place, keeping the rest of the view", () => {
    const wb = readWorkbook(book(`<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" state="frozen"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/></sheetView></sheetViews>${DATA}`));
    Object.assign(wb.sheets[0]!, { freeze: { rows: 3, cols: 0 }, freezeDirty: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/ySplit="3"/);
    expect(xml).toContain(`showGridLines="0"`);
    expect(xml).toContain(`activeCell="A5"`);
  });

  it("removes the pane (and the selection's pane hint) when unfreezing", () => {
    const wb = readWorkbook(book(`<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" state="frozen"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/></sheetView></sheetViews>${DATA}`));
    Object.assign(wb.sheets[0]!, { freeze: undefined, freezeDirty: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).not.toContain("<pane");
    expect(xml).not.toContain(`pane="bottomLeft"`);
    expect(xml).toContain(`activeCell="A5"`);
  });

  it("round-trips a freeze it wrote", () => {
    const wb = readWorkbook(book(DATA));
    Object.assign(wb.sheets[0]!, { freeze: { rows: 1, cols: 2 }, freezeDirty: true });
    expect(readWorkbook(writeWorkbook(wb)).sheets[0]!.freeze).toEqual({ rows: 1, cols: 2 });
  });
});
