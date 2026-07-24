import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { writeXlsx } from "./write";
import { deleteXlsxShape } from "./shape-write";
import { writeWorkbook } from "../../core/workbook";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

function base(drawingBody: string): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheetData/>${drawingBody ? '<drawing r:id="rId1"/>' : ""}</worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  };
  if (drawingBody) {
    files["xl/worksheets/_rels/sheet1.xml.rels"] = strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`);
    files["xl/drawings/drawing1.xml"] = strToU8(`<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">${drawingBody}</xdr:wsDr>`);
  }
  return zipSync(files);
}

const spAnchor = (prst: string) => `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;

describe("drawing shapes", () => {
  it("reads an xdr:sp shape (geometry, fill, outline, text)", () => {
    const sh = readWorkbook(base(spAnchor("ellipse"))).sheets[0].shapes ?? [];
    expect(sh.length).toBe(1);
    expect(sh[0].geom).toBe("ellipse");
    expect(sh[0].fill?.toLowerCase()).toBe("#4472c4");
    expect(sh[0].stroke?.toLowerCase()).toBe("#222222");
    expect(sh[0].strokeWidth).toBe(2); // 19050 EMU / 9525
    expect(sh[0].text).toBe("Hi");
    expect(sh[0].anchor.fromCol).toBe(2); // B2
  });

  it("reads and authors the extended preset shapes (diamond / star / arrow)", () => {
    expect(readWorkbook(base(spAnchor("diamond"))).sheets[0].shapes![0].geom).toBe("diamond");
    expect(readWorkbook(base(spAnchor("star5"))).sheets[0].shapes![0].geom).toBe("star");
    expect(readWorkbook(base(spAnchor("rightArrow"))).sheets[0].shapes![0].geom).toBe("rightArrow");
    const wb = readWorkbook(base(""));
    (wb.sheets[0].shapes ??= []).push({ geom: "star", anchor: { fromCol: 2, fromRow: 2, fromColOff: 0, fromRowOff: 0, toCol: 5, toRow: 6, toColOff: 0, toRowOff: 0 }, fill: "#ffcc00", created: true, dirty: true });
    writeXlsx(wb);
    expect(new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"])).toContain('prst="star5"');
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].shapes![0].geom).toBe("star");
  });

  it("maps an unknown preset to a rect but keeps the original preset name", () => {
    const sh = readWorkbook(base(spAnchor("cloud"))).sheets[0].shapes ?? [];
    expect(sh[0].geom).toBe("rect");
    expect(sh[0].preset).toBe("cloud");
  });

  it("authors a new shape into a drawing that did not exist, and round-trips it", () => {
    const wb = readWorkbook(base("")); // no drawing yet
    (wb.sheets[0].shapes ??= []).push({
      geom: "roundRect",
      anchor: { fromCol: 2, fromRow: 2, fromColOff: 0, fromRowOff: 0, toCol: 5, toRow: 6, toColOff: 0, toRowOff: 0 },
      fill: "#ff0000", stroke: "#000000", strokeWidth: 2, text: "New", created: true, dirty: true,
    });
    writeXlsx(wb);
    const draw = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    expect(draw).toContain('prst="roundRect"');
    expect(draw).toContain("<a:srgbClr val=\"FF0000\"/>");
    expect(draw).toContain("<a:t>New</a:t>");
    // the worksheet now references a drawing part
    expect(new TextDecoder().decode(wb.files["xl/worksheets/sheet1.xml"])).toContain("<drawing");
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].shapes ?? [];
    expect(re).toHaveLength(1);
    expect(re[0]).toMatchObject({ geom: "roundRect", text: "New" });
    expect(re[0].fill?.toLowerCase()).toBe("#ff0000");
  });

  it("patches an existing shape's move + fill in place", () => {
    const wb = readWorkbook(base(spAnchor("rect")));
    const sh = wb.sheets[0].shapes![0];
    sh.anchor = { ...sh.anchor, fromCol: 3, fromRow: 4 };
    sh.fill = "#00ff00";
    sh.dirty = true;
    writeXlsx(wb);
    const draw = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    expect(draw).toMatch(/<xdr:from><xdr:col>2<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>3<\/xdr:row>/);
    expect(draw).toContain("00FF00");
    expect(draw).not.toContain("4472C4"); // old fill replaced
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].shapes![0];
    expect(re.anchor.fromCol).toBe(3);
    expect(re.fill?.toLowerCase()).toBe("#00ff00");
  });

  it("deletes a shape and re-indexes the remaining ones", () => {
    const wb = readWorkbook(base(spAnchor("rect") + spAnchor("ellipse")));
    const sheet = wb.sheets[0];
    expect(sheet.shapes).toHaveLength(2);
    const first = sheet.shapes![0]; // anchorIndex 0
    const second = sheet.shapes![1]; // anchorIndex 1
    sheet.shapes!.splice(0, 1); // model removal (as the editor does)
    deleteXlsxShape(wb, sheet, first);
    expect(second.anchorIndex).toBe(0); // shifted down
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].shapes ?? [];
    expect(re).toHaveLength(1);
    expect(re[0].geom).toBe("ellipse");
  });

  it("leaves an untouched shape's drawing part byte-identical", () => {
    const wb = readWorkbook(base(spAnchor("rect")));
    const before = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    writeXlsx(wb);
    expect(new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"])).toBe(before);
  });
});
