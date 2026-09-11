import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";

// Selecting a cell and editing it are two different things. Every display overlay (multi-format
// runs, rotation, wrapping, ruby, spill) gives way to the plain <input> that carries the text, and
// that handover used to happen on focus: clicking a cell once was enough to flatten a half-bold
// title to one unstyled line, and a run style applied mid-edit stayed invisible until the caret
// left the cell. The handover now waits for an actual edit, which is what "editing" marks.

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

// A1 is a two-run string: "Half" in bold, "plain" not. A2 is ordinary text.
const SHARED = `<sst xmlns="${SS}" count="2" uniqueCount="2">
 <si><r><rPr><b/></rPr><t>Half</t></r><r><t>plain</t></r></si>
 <si><t>ordinary</t></si>
</sst>`;

function book(): Uint8Array {
  const body = `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row></sheetData>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}" xmlns:r="${R}">${body}</worksheet>`),
    "xl/sharedStrings.xml": strToU8(SHARED),
  });
}

async function mount(): Promise<HTMLElement> {
  const host = document.createElement("div");
  document.body.appendChild(host);
  createSheetEditor(host, book(), { filename: "s.xlsx" });
  await frame();
  return host;
}

const cell = (host: HTMLElement, rc: string) => host.querySelector(`td[data-rc="${rc}"]`) as HTMLTableCellElement;

describe("a cell that is merely selected", () => {
  it("keeps drawing its runs instead of handing over to the input", async () => {
    const host = await mount();
    const td = cell(host, "1:1");
    expect(td.classList.contains("has-rich"), "the two-run cell has a display overlay").toBe(true);
    td.querySelector("input")!.focus();
    await frame();
    expect(td.classList.contains("editing"), "focus alone is not an edit").toBe(false);
    expect(td.querySelector(".sheetedit-cellrich")!.textContent).toBe("Halfplain");
  });

  it("hands over as soon as the text is typed into", async () => {
    const host = await mount();
    const td = cell(host, "1:1");
    const input = td.querySelector("input")!;
    input.focus();
    input.value = "Halfplain!";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    expect(td.classList.contains("editing")).toBe(true);
  });

  it("hands over on F2, with the caret at the end", async () => {
    const host = await mount();
    const td = cell(host, "1:1");
    const input = td.querySelector("input")!;
    input.focus();
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "F2", bubbles: true }));
    expect(td.classList.contains("editing")).toBe(true);
    expect(input.selectionStart).toBe(input.value.length);
  });

  it("hands over on a second click, not on the one that selected it", async () => {
    const host = await mount();
    const td = cell(host, "1:1");
    const input = td.querySelector("input")!;
    input.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    input.focus();
    input.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(td.classList.contains("editing"), "the click that selects the cell does not edit it").toBe(false);
    input.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    input.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(td.classList.contains("editing"), "a click on the cell the caret is already in").toBe(true);
  });

  it("stops editing when the caret leaves", async () => {
    const host = await mount();
    const td = cell(host, "1:1");
    const input = td.querySelector("input")!;
    input.focus();
    input.dispatchEvent(new Event("input", { bubbles: true }));
    expect(td.classList.contains("editing")).toBe(true);
    input.blur();
    await frame();
    expect(cell(host, "1:1").classList.contains("editing")).toBe(false);
  });

  it("edits one cell at a time", async () => {
    const host = await mount();
    const a1 = cell(host, "1:1"), a2 = cell(host, "2:1");
    a1.querySelector("input")!.focus();
    a1.querySelector("input")!.dispatchEvent(new Event("input", { bubbles: true }));
    a2.querySelector("input")!.focus();
    a2.querySelector("input")!.dispatchEvent(new Event("input", { bubbles: true }));
    await frame();
    expect(cell(host, "1:1").classList.contains("editing")).toBe(false);
    expect(cell(host, "2:1").classList.contains("editing")).toBe(true);
  });
});
