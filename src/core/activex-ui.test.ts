import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";
import { readWorkbook } from "./workbook";

// An ActiveX combo differs from a form-control one in what it writes: the chosen TEXT rather than
// the 1-based position. Excel's two control families diverge there, and the linked cell is what
// every formula downstream reads, so the difference is not cosmetic.

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

const frame = (): Promise<void> => new Promise((res) => setTimeout(res, 30));
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const AX = "http://schemas.microsoft.com/office/2006/activeX";
const AXREL = "http://schemas.microsoft.com/office/2006/relationships";

const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff];
const u32 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];
const padded = (t: string): number[] => {
  const b = [...t].map((c) => c.charCodeAt(0));
  return [...b, ...Array((4 - (b.length % 4)) % 4).fill(0)];
};
const COMBO_CLSID = [0x30, 0x1d, 0xd2, 0x8b, 0x42, 0xec, 0xce, 0x11, 0x9e, 0x0d, 0x00, 0xaa, 0x00, 0x60, 0x02, 0xf3];

/** A combo whose persisted value is `value`, as Excel writes one. */
const comboBin = (value: string): Uint8Array => {
  const data = [...u32(0x2c80481b), 0x03, 0x01, 0x02, 0x00, ...u32(0x80000000 | value.length)];
  const extra = [...u32(6244), ...u32(900), ...padded(value)];
  return new Uint8Array([...COMBO_CLSID, 0, 2, ...u16(8 + data.length + extra.length), ...u32(0x80450141), ...u32(0), ...data, ...extra]);
};

/** A workbook with one ActiveX combo: linked to B1, listing from the name "Days". */
function book(): Uint8Array {
  const anchor =
    `<anchor moveWithCells="1">` +
    `<from><xdr:col xmlns:xdr="${XDR}">3</xdr:col><xdr:colOff xmlns:xdr="${XDR}">0</xdr:colOff><xdr:row xmlns:xdr="${XDR}">0</xdr:row><xdr:rowOff xmlns:xdr="${XDR}">0</xdr:rowOff></from>` +
    `<to><xdr:col xmlns:xdr="${XDR}">5</xdr:col><xdr:colOff xmlns:xdr="${XDR}">0</xdr:colOff><xdr:row xmlns:xdr="${XDR}">2</xdr:row><xdr:rowOff xmlns:xdr="${XDR}">0</xdr:rowOff></to>` +
    `</anchor>`;
  const cell = (ref: string, text: string): string => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>` +
      `<definedNames><definedName name="Days">S!$A$1:$A$3</definedName></definedNames></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(
      `<worksheet xmlns="${MAIN}" xmlns:r="${R}" xmlns:mc="${MC}" mc:Ignorable="x14"><sheetData>` +
      `<row r="1">${cell("A1", "Mon")}</row><row r="2">${cell("A2", "Tue")}</row><row r="3">${cell("A3", "Wed")}</row>` +
      `</sheetData><controls><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="x14">` +
      `<control shapeId="1" r:id="rIdA" name="Combo1"><controlPr defaultSize="0" linkedCell="B1" listFillRange="Days" r:id="rIdImg">${anchor}</controlPr></control>` +
      `</mc:Choice><mc:Fallback><control shapeId="1" r:id="rIdA" name="Combo1"/></mc:Fallback></mc:AlternateContent></controls></worksheet>`),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(
      `<Relationships xmlns="${RELNS}"><Relationship Id="rIdA" Type="${AXREL}/control" Target="../activeX/activeX1.xml"/></Relationships>`),
    "xl/activeX/activeX1.xml": strToU8(`<ax:ocx xmlns:ax="${AX}" xmlns:r="${R}" ax:classid="{8BD21D30-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistStreamInit" r:id="rId1"/>`),
    "xl/activeX/_rels/activeX1.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${AXREL}/activeXControlBinary" Target="activeX1.bin"/></Relationships>`),
    "xl/activeX/activeX1.bin": comboBin("Tue"),
  });
}

describe("an ActiveX combo in the grid", () => {
  it("reads its linked cell and its item source off the worksheet", () => {
    const ctl = readWorkbook(book()).sheets[0]!.controls![0]!;
    expect(ctl.linkedCell).toBe("B1");
    expect(ctl.sourceRange).toBe("Days");
    expect(ctl.activeXValue).toBe("Tue");
  });

  it("lists the items its listFillRange names, resolving it as a defined name", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, book(), { fileName: "a.xlsm" });
    await frame();
    const select = host.querySelector<HTMLSelectElement>(".sheetedit-ctrlbox select")!;
    expect([...select.options].map((o) => o.text)).toEqual(["", "Mon", "Tue", "Wed"]);
    expect(select.value).toBe("Tue");   // the persisted value, selected
    expect(select.disabled).toBe(false); // a linked cell makes it live
    ed.destroy();
    host.remove();
  });

  it("writes the chosen TEXT to the linked cell, not a position", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, book(), { fileName: "a.xlsm" });
    await frame();
    const select = host.querySelector<HTMLSelectElement>(".sheetedit-ctrlbox select")!;
    select.value = "Wed";
    select.dispatchEvent(new Event("change"));
    await frame();
    // B1 is the linked cell: a form control would have put "3" there, an ActiveX one puts "Wed".
    expect(host.querySelector<HTMLInputElement>('td[data-rc="1:2"] input')?.value).toBe("Wed");
    ed.destroy();
    host.remove();
  });
});
