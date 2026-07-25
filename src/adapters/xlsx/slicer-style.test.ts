import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { styleAccent } from "../../core/ui/slicer-layer";

const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

/**
 * Add a user-defined slicer style to demo/c-slicer.xlsx: the x14:slicerStyles group and the x14
 * dxfs list it indexes, both in styles.xml's extLst the way Excel writes them.
 */
function withStyle(opts: { name?: string; elements?: string; dxfs?: string } = {}): Uint8Array {
  const files = unzipSync(new Uint8Array(readFileSync("demo/c-slicer.xlsx")));
  const name = opts.name ?? "MySlicerStyle";
  const dxfs = opts.dxfs ??
    `<x14:dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill><font><color rgb="FFFFFFFF"/></font></x14:dxf>` +
    `<x14:dxf><fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill><font><color rgb="FF112233"/></font></x14:dxf>`;
  const elements = opts.elements ??
    `<x14:slicerStyleElement type="selectedItemWithData" dxfId="0"/>` +
    `<x14:slicerStyleElement type="unselectedItemWithData" dxfId="1"/>`;
  const ext =
    `<extLst>` +
    `<ext uri="{46F421CA-312F-682f-3DD2-61675219B42D}" xmlns:x14="${X14}"><x14:dxfs count="2">${dxfs}</x14:dxfs></ext>` +
    `<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="${X14}">` +
      `<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1">` +
        `<x14:slicerStyle name="${name}"><x14:slicerStyleElements count="2">${elements}</x14:slicerStyleElements></x14:slicerStyle>` +
      `</x14:slicerStyles></ext>` +
    `</extLst>`;
  files["xl/styles.xml"] = strToU8(strFromU8(files["xl/styles.xml"]!).replace("</styleSheet>", ext + "</styleSheet>"));
  return zipSync(files);
}

describe("custom slicer styles", () => {
  it("reads a user-defined style's selected and unselected colours", () => {
    const wb = readWorkbook(withStyle());
    const def = wb.slicerStyles?.get("MySlicerStyle");
    expect(def).toBeTruthy();
    expect(def!.selectedFill).toBe("#ff0000");
    expect(def!.selectedText).toBe("#ffffff");
    expect(def!.unselectedFill).toBe("#00ff00");
    expect(def!.unselectedText).toBe("#112233");
  });

  it("takes the dxf's bgColor as the fill, the way a differential format stores it", () => {
    const wb = readWorkbook(withStyle({
      dxfs: `<x14:dxf><fill><patternFill><fgColor rgb="FF123456"/><bgColor rgb="FF654321"/></patternFill></fill></x14:dxf>`,
      elements: `<x14:slicerStyleElement type="selectedItemWithData" dxfId="0"/>`,
    }));
    expect(wb.slicerStyles?.get("MySlicerStyle")?.selectedFill).toBe("#654321");
  });

  it("falls back to fgColor when the dxf has no bgColor", () => {
    const wb = readWorkbook(withStyle({
      dxfs: `<x14:dxf><fill><patternFill><fgColor rgb="FF123456"/></patternFill></fill></x14:dxf>`,
      elements: `<x14:slicerStyleElement type="selectedItemWithData" dxfId="0"/>`,
    }));
    expect(wb.slicerStyles?.get("MySlicerStyle")?.selectedFill).toBe("#123456");
  });

  it("ignores an element whose dxfId points past the list", () => {
    const wb = readWorkbook(withStyle({
      dxfs: `<x14:dxf><fill><patternFill><bgColor rgb="FFAABBCC"/></patternFill></fill></x14:dxf>`,
      elements: `<x14:slicerStyleElement type="selectedItemWithData" dxfId="7"/>`,
    }));
    const def = wb.slicerStyles?.get("MySlicerStyle");
    expect(def).toBeTruthy();
    expect(def!.selectedFill).toBeUndefined();
  });

  it("reads the hovered and no-data variants without applying them", () => {
    const wb = readWorkbook(withStyle({
      elements: `<x14:slicerStyleElement type="hoveredSelectedItemWithData" dxfId="0"/>` +
        `<x14:slicerStyleElement type="selectedItemWithNoData" dxfId="1"/>`,
    }));
    const def = wb.slicerStyles?.get("MySlicerStyle")!;
    expect(def.selectedFill).toBeUndefined();
    expect(def.unselectedFill).toBeUndefined();
  });

  it("leaves slicerStyles unset when the workbook defines none", () => {
    const wb = readWorkbook(new Uint8Array(readFileSync("demo/c-slicer.xlsx")));
    expect(wb.slicerStyles).toBeUndefined();
  });

  it("keeps the style name on the slicer so it round-trips", () => {
    const wb = readWorkbook(withStyle());
    const sl = wb.sheets.flatMap((s) => s.slicers ?? [])[0];
    // c-slicer.xlsx names a built-in style; the custom map is separate from what the slicer stores.
    expect(sl).toBeTruthy();
    expect(wb.slicerStyles!.size).toBe(1);
  });

  it("still maps built-in style families to a theme accent", () => {
    expect(styleAccent("SlicerStyleLight2")).toBe("#ed7d31");
    expect(styleAccent("MySlicerStyle")).toBeUndefined();
  });
});
