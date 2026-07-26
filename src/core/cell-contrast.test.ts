import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";

// A spreadsheet's default text colour is black, not the UI theme's foreground. A cell the file
// fills but gives no colour to has to be read against that fill: in dark mode the theme's light
// foreground on a pale fill left the value all but invisible, which a screenshot caught.

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

/** A1 pale green with no font colour, B1 dark blue with none, C1 pale with an explicit red. */
function book(): Uint8Array {
  const fill = (rgb: string): string => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/></patternFill></fill>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(
      `<worksheet xmlns="${MAIN}"><sheetData><row r="1">` +
      `<c r="A1" s="1" t="inlineStr"><is><t>pale</t></is></c>` +
      `<c r="B1" s="2" t="inlineStr"><is><t>dark</t></is></c>` +
      `<c r="C1" s="3" t="inlineStr"><is><t>own</t></is></c>` +
      `<c r="D1" t="inlineStr"><is><t>plain</t></is></c>` +
      `</row></sheetData></worksheet>`),
    "xl/styles.xml": strToU8(
      `<styleSheet xmlns="${MAIN}">` +
      `<fonts count="2"><font><sz val="11"/></font><font><color rgb="FFFF0000"/></font></fonts>` +
      `<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
      `${fill("FFEBF1DE")}${fill("FF1F3864")}</fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="4">` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
      `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>` +
      `<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>` +
      `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/>` +
      `</cellXfs></styleSheet>`),
  });
}

describe("text on a filled cell", () => {
  it("goes dark on a pale fill and light on a dark one", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, book(), { fileName: "c.xlsx" });
    await frame();
    const at = (rc: string): HTMLInputElement => host.querySelector<HTMLInputElement>(`td[data-rc="${rc}"] input`)!;
    expect(at("1:1").style.color).toBe("rgb(26, 26, 26)");    // pale green -> near-black
    expect(at("1:2").style.color).toBe("rgb(245, 245, 245)"); // dark blue -> near-white
    ed.destroy();
    host.remove();
  });

  it("never overrides a colour the file states", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, book(), { fileName: "c.xlsx" });
    await frame();
    const own = host.querySelector<HTMLInputElement>('td[data-rc="1:3"] input')!;
    expect(own.style.color).toBe("rgb(255, 0, 0)");
    ed.destroy();
    host.remove();
  });

  it("leaves an unfilled cell to the theme, so dark mode still works", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, book(), { fileName: "c.xlsx" });
    await frame();
    expect(host.querySelector<HTMLInputElement>('td[data-rc="1:4"] input')!.style.color).toBe("");
    ed.destroy();
    host.remove();
  });
});
