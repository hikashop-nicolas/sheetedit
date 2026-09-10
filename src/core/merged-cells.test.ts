import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";

// What a merged range looks like on screen. Both of these come from the same place: the grid draws
// only the top-left cell of a merge and treats it as if it were the whole thing, when in the file
// the covered cells still carry their own borders, and a hidden column inside the range is still
// counted in its width.

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

const frame = () => new Promise<void>((res) => setTimeout(res, 40));
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

// Two borders: id 1 has a top edge only, id 2 a bottom edge only. A vertical merge of a cell
// styled with the first over one styled with the second must show BOTH.
const STYLES = `<styleSheet xmlns="${SS}">
 <fonts count="1"><font><sz val="11"/></font></fonts>
 <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
 <borders count="3">
  <border><left/><right/><top/><bottom/></border>
  <border><left/><right/><top style="thin"><color rgb="FF112233"/></top><bottom/></border>
  <border><left/><right/><top/><bottom style="medium"><color rgb="FF445566"/></bottom></border>
 </borders>
 <cellXfs count="3">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
 </cellXfs>
</styleSheet>`;

function book(sheetBody: string, cols = ""): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}" xmlns:r="${R}">${cols}${sheetBody}</worksheet>`),
    "xl/styles.xml": strToU8(STYLES),
  });
}

async function mount(bytes: Uint8Array): Promise<HTMLElement> {
  const host = document.createElement("div");
  document.body.appendChild(host);
  createSheetEditor(host, bytes, { filename: "s.xlsx" });
  await frame();
  return host;
}

describe("a merged range's borders", () => {
  it("shows the edge borders of the cells it covers, not just the top-left one's", async () => {
    // A2 (covered) carries the bottom border of the A1:A2 merge.
    const body = `<sheetData><row r="1"><c r="A1" s="1" t="str"><v>x</v></c></row><row r="2"><c r="A2" s="2"/></row></sheetData><mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>`;
    const host = await mount(book(body));
    const td = host.querySelector('td[data-rc="1:1"]') as HTMLElement;
    expect(td.rowSpan).toBe(2);
    expect(td.style.boxShadow).toContain("#112233"); // its own top edge
    expect(td.style.boxShadow, "the covered cell's bottom edge").toContain("#445566");
  });
});

describe("a merge containing a hidden column", () => {
  it("spans only the columns that are drawn", async () => {
    // A1:C1 merged with B hidden: two columns are rendered, so the span is 2, not 3. Counting the
    // hidden one pushed the merge a column too wide and shunted the rest of the row along.
    const cols = `<cols><col min="2" max="2" width="9" hidden="1" customWidth="1"/></cols>`;
    const body = `<sheetData><row r="1"><c r="A1" t="str"><v>wide</v></c><c r="B1"/><c r="C1"/><c r="D1" t="str"><v>after</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>`;
    const host = await mount(book(body, cols));
    const td = host.querySelector('td[data-rc="1:1"]') as HTMLElement;
    expect(td.colSpan).toBe(2);
    // D1 still renders, and to the right of the merge rather than pushed off it.
    const after = host.querySelector('td[data-rc="1:4"]') as HTMLElement;
    expect(after).toBeTruthy();
  });

  it("spans every column when none is hidden", async () => {
    const body = `<sheetData><row r="1"><c r="A1" t="str"><v>wide</v></c><c r="B1"/><c r="C1"/></row></sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>`;
    const host = await mount(book(body));
    expect((host.querySelector('td[data-rc="1:1"]') as HTMLElement).colSpan).toBe(3);
  });
});
