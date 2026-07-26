import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { writeXlsx } from "./write";
import { deleteXlsxShape } from "./shape-write";
import { writeWorkbook } from "../../core/workbook";
import { shapeSvg } from "../../core/ui/shape-layer";

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

  it("paints a gallery shape from its <xdr:style> theme references", () => {
    // What Excel's shape gallery writes: no fill and no line of its own, only refs into the
    // theme's format scheme. Ignoring them leaves the shape unfilled with unreadable text.
    const styled = `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx="1"><a:schemeClr val="accent3"/></a:lnRef><a:fillRef idx="2"><a:schemeClr val="accent3"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent3"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></xdr:style><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Go</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;
    const theme = `<a:theme xmlns:a="${A}"><a:themeElements><a:clrScheme name="t"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="70AD47"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="A5A5A5"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fmtScheme name="f"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:tint val="60000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:gs></a:gsLst></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/></a:schemeClr></a:solidFill></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
    const files = unzipSync(base(styled));
    files["xl/theme/theme1.xml"] = strToU8(theme);
    const sh = readWorkbook(zipSync(files)).sheets[0].shapes![0]!;
    // fillRef idx 2 is the gradient; its first stop is accent3 lightened, so a pale green.
    expect(sh.fill).toMatch(/^#[0-9a-f]{6}$/);
    expect(sh.fill).not.toBe("#70ad47"); // transformed, not the raw accent
    const lum = parseInt(sh.fill!.slice(3, 5), 16);
    expect(lum, "the tint lightens it").toBeGreaterThan(0xad);
    expect(sh.stroke).toBeDefined();
    expect(sh.textColor).toBe("#000000"); // the fontRef's dk1
    // The recipe is a two-stop gradient, and it is kept as one rather than flattened.
    expect(sh.fillGradient?.stops.length).toBe(2);
    expect(sh.fillGradient?.stops[0]!.pos).toBe(0);
    expect(sh.fillGradient?.stops[1]!.pos).toBe(1);
    expect(sh.fillGradient?.stops[0]!.color).toBe(sh.fill); // fill is the first stop
    expect(sh.fillGradient?.angle).toBe(90); // no <a:lin>, so top to bottom
    const svg = shapeSvg(sh, 100, 40);
    expect(svg).toContain("<linearGradient");
    expect(svg).toMatch(/fill="url\(#sheetedit-grad-\d+\)"/);
    expect(svg).toContain(`stop-color="${sh.fillGradient!.stops[1]!.color}"`);
  });

  it("reads a shape's own gradFill, with the angle the file states", () => {
    const own = `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>G</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;
    const sh = readWorkbook(base(own)).sheets[0].shapes![0]!;
    expect(sh.fill).toBe("#ff0000");
    expect(sh.fillGradient?.angle).toBe(0); // left to right
    expect(sh.fillGradient?.stops.map((x) => x.color)).toEqual(["#ff0000", "#0000ff"]);
    // ang 0 runs along x, so the gradient vector spans the box horizontally.
    expect(shapeSvg(sh, 100, 40)).toContain(`x1="0" y1="0.5" x2="1" y2="0.5"`);
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
