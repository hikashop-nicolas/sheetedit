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
    // gear9 is one of the presets with no drawing of its own; the name is still carried so a
    // round-trip re-emits the shape exactly as the file had it.
    const sh = readWorkbook(base(spAnchor("gear9"))).sheets[0].shapes ?? [];
    expect(sh[0].geom).toBe("rect");
    expect(sh[0].preset).toBe("gear9");
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
    sh.styleDirty = true; // the property dialog changed the paint, not just the position
    writeXlsx(wb);
    const draw = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    expect(draw).toMatch(/<xdr:from><xdr:col>2<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>3<\/xdr:row>/);
    expect(draw).toContain("00FF00");
    expect(draw).not.toContain("4472C4"); // old fill replaced
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].shapes![0];
    expect(re.anchor.fromCol).toBe(3);
    expect(re.fill?.toLowerCase()).toBe("#00ff00");
  });

  it("moving a shape rewrites its anchor and nothing else", () => {
    // Dragging a shape must not repaint it. The model carries one colour, so rewriting the paint
    // on a move flattens whatever the file states, and a shape whose look comes from the theme
    // (an <xdr:style> ref, with no fill of its own) loses it outright.
    const wb = readWorkbook(base(spAnchor("rect")));
    const sh = wb.sheets[0].shapes![0];
    const before = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    sh.anchor = { ...sh.anchor, fromCol: 3, fromRow: 4 };
    sh.dirty = true;
    writeXlsx(wb);
    const draw = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    expect(draw).toMatch(/<xdr:from><xdr:col>2<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>3<\/xdr:row>/);
    expect(draw).toContain('<a:srgbClr val="4472C4"/>'); // the file's own fill, untouched
    // Everything outside the anchor is byte-identical (the writer adds the XML declaration).
    const body = (x: string) => x.replace(/^<\?xml[^?]*\?>\s*/, "").replace(/<xdr:from>[\s\S]*?<\/xdr:to>/, "");
    expect(body(draw)).toBe(body(before));
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

// A shape's text is paragraphs, and a connector's ends carry arrowheads. Both were dropped on the
// way in: every paragraph's runs were concatenated into one string, and the line ends were never
// looked at, so a two-line callout came out as one clipped line at the end of an arrow with no
// point on it.
describe("shape text and line ends", () => {
  const textBox = (paras: string, bodyPr = '<a:bodyPr anchor="t"/>') =>
    `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="T"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody>${bodyPr}<a:lstStyle/>${paras}</xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;

  it("keeps each paragraph on its own line", () => {
    // Two paragraphs, the first split across two runs (Excel splits on a formatting boundary).
    const paras = `<a:p><a:pPr algn="l"/><a:r><a:t>Montant a inscrir</a:t></a:r><a:r><a:t>e au contrat.</a:t></a:r></a:p><a:p><a:pPr algn="l"/><a:r><a:t>Second line.</a:t></a:r></a:p>`;
    const sh = readWorkbook(base(textBox(paras))).sheets[0].shapes![0]!;
    expect(sh.text).toBe("Montant a inscrire au contrat.\nSecond line.");
    expect(sh.textAlign).toBe("left");
    expect(sh.textValign).toBe("top");
  });

  it("does not leave a trailing blank line for an empty last paragraph", () => {
    const sh = readWorkbook(base(textBox(`<a:p><a:r><a:t>One</a:t></a:r></a:p><a:p><a:endParaRPr lang="en-GB"/></a:p>`))).sheets[0].shapes![0]!;
    expect(sh.text).toBe("One");
  });

  it("draws the text as wrapping HTML, positioned as the file asks", () => {
    const sh = readWorkbook(base(textBox(`<a:p><a:pPr algn="l"/><a:r><a:t>a\nb</a:t></a:r></a:p>`))).sheets[0].shapes![0]!;
    const svg = shapeSvg(sh, 120, 40);
    expect(svg).toContain("<foreignObject");
    expect(svg).toContain("white-space:pre-wrap");
    expect(svg).toContain("text-align:left");
    expect(svg).toContain("justify-content:flex-start"); // anchor="t"
  });

  const connector = (ln: string, flip = "") =>
    `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id="3" name="C"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm${flip}/><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill>${ln}</a:ln></xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>`;

  it("reads the line ends and draws an arrowhead for them", () => {
    const sh = readWorkbook(base(connector(`<a:headEnd type="none"/><a:tailEnd type="triangle"/>`))).sheets[0].shapes![0]!;
    expect(sh.tailEnd).toBe("triangle");
    const svg = shapeSvg(sh, 100, 50);
    expect(svg).toContain("<marker");
    expect(svg).toMatch(/marker-end="url\(#sheetedit-marker-\d+\)"/);
    expect(svg).not.toContain("marker-start="); // headEnd "none" draws nothing
    expect(svg).toContain('fill="#0070C0"'); // the marker takes the line's colour
  });

  it("runs a flipped line from the other corner of its box", () => {
    const plain = shapeSvg(readWorkbook(base(connector(""))).sheets[0].shapes![0]!, 100, 50);
    expect(plain).toMatch(/x1="0\.5" y1="0\.5" x2="99\.5" y2="49\.5"/);
    const sh = readWorkbook(base(connector("", ' flipH="1"'))).sheets[0].shapes![0]!;
    expect(sh.flipH).toBe(true);
    expect(shapeSvg(sh, 100, 50)).toMatch(/x1="99\.5" y1="0\.5" x2="0\.5" y2="49\.5"/);
  });
});

// Braces and brackets are open outlines, not closed shapes. An unrecognised preset falls back to a
// rectangle, and a brace drawn as the rectangle around the columns it groups says nothing at all.
describe("braces and brackets", () => {
  const braceAnchor = (prst: string, xfrm = "<a:xfrm/>") =>
    `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="B"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>${xfrm}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;

  it("reads each brace and bracket preset as itself", () => {
    for (const prst of ["leftBrace", "rightBrace", "bracePair", "leftBracket", "rightBracket", "bracketPair"]) {
      expect(readWorkbook(base(braceAnchor(prst))).sheets[0].shapes![0]!.geom, prst).toBe(prst);
    }
  });

  it("draws one as a stroked open path, never a filled box", () => {
    const sh = readWorkbook(base(braceAnchor("rightBrace"))).sheets[0].shapes![0]!;
    const svg = shapeSvg(sh, 20, 200);
    expect(svg).toContain("<path");
    expect(svg).not.toContain("<rect");
    expect(svg).toContain('fill="none"');
    expect(svg).toContain('stroke="#ED7D31"');
    expect(svg).not.toMatch(/[Zz]"/); // open: the path is never closed
  });

  // A "{" is a "}" mirrored, so the two must not come out identical, which is what a shared
  // fallback would give.
  it("mirrors the left and right forms", () => {
    const left = shapeSvg(readWorkbook(base(braceAnchor("leftBrace"))).sheets[0].shapes![0]!, 20, 200);
    const right = shapeSvg(readWorkbook(base(braceAnchor("rightBrace"))).sheets[0].shapes![0]!, 20, 200);
    expect(left).not.toBe(right);
  });

  it("draws a pair as two outlines", () => {
    const pair = shapeSvg(readWorkbook(base(braceAnchor("bracketPair"))).sheets[0].shapes![0]!, 60, 200);
    expect((pair.match(/M /g) ?? []).length).toBe(2);
  });

  // A brace bracketing a row of columns is a tall one turned on its side. Without the angle it is
  // drawn upright, down the rows it was meant to sit under.
  it("reads the rotation angle, in degrees", () => {
    const sh = readWorkbook(base(braceAnchor("rightBrace", '<a:xfrm rot="16200000"/>'))).sheets[0].shapes![0]!;
    expect(sh.rotation).toBe(270); // 16200000 / 60000
    expect(readWorkbook(base(braceAnchor("rightBrace"))).sheets[0].shapes![0]!.rotation).toBeUndefined();
  });
});

// A connector says something with more than its position: a dashed shaft means provisional, an
// elbow points around what is in the way, and a see-through fill is a watermark rather than a slab
// over the cells. All of it was dropped on the way in.
describe("how a shape is painted", () => {
  const shape = (spPr: string, txBody = "") =>
    `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>${spPr}</xdr:spPr>${txBody}</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;
  const read = (spPr: string, txBody = "") => readWorkbook(base(shape(spPr, txBody))).sheets[0].shapes![0]!;

  it("keeps a dashed outline dashed", () => {
    const sh = read(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="0F9ED5"/></a:solidFill><a:prstDash val="dash"/></a:ln>`);
    expect(sh.dash).toBe("dash");
    expect(shapeSvg(sh, 100, 50)).toContain('stroke-dasharray="8 6"'); // 4x and 3x the 2px stroke
  });

  it("leaves a solid outline undashed", () => {
    const sh = read(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln>`);
    expect(shapeSvg(sh, 100, 50)).not.toContain("stroke-dasharray");
  });

  it("reads a see-through fill and paints it see-through", () => {
    const sh = read(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="D9D9D9"><a:alpha val="68000"/></a:srgbClr></a:solidFill>`);
    expect(sh.fillOpacity).toBeCloseTo(0.68, 5);
    expect(sh.fill?.toLowerCase()).toBe("#d9d9d9"); // the colour itself stays a plain hex
    expect(shapeSvg(sh, 100, 50)).toContain('fill-opacity="0.68"');
  });

  it("leaves an opaque fill alone", () => {
    const sh = read(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill>`);
    expect(sh.fillOpacity).toBeUndefined();
    expect(shapeSvg(sh, 100, 50)).not.toContain("fill-opacity");
  });

  it("draws an elbow connector as a bent line, not a box", () => {
    const sh = read(`<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 25000"/></a:avLst></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>`);
    expect(sh.geom).toBe("elbow");
    expect(sh.adjust).toBeCloseTo(0.25, 5);
    const svg = shapeSvg(sh, 100, 40);
    expect(svg).not.toContain("<rect");
    expect(svg).toContain('fill="none"');
    expect(svg).toMatch(/d="M 0 0 L 24\.75 0 L 24\.75 39 L 99 39"/); // out, turn at a quarter, across
    expect(svg).toMatch(/marker-end="url\(#sheetedit-marker-\d+\)"/); // the arrow survives the bend
  });

  it("sizes rotated text as the file does", () => {
    const sh = read(
      `<a:xfrm rot="1442775"><a:ext cx="10648950" cy="1543050"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`,
      `<xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="fr-FR" sz="4000"/><a:t>WATERMARK</a:t></a:r></a:p></xdr:txBody>`,
    );
    expect(sh.textSize).toBe(40);
    expect(sh.rotation).toBeCloseTo(24.046, 2);
    // The <a:ext> is the shape's own size; the anchor is where it lands once turned.
    expect(sh.extent!.w).toBeCloseTo(1118, 0);
    expect(sh.extent!.h).toBeCloseTo(162, 0);
    expect(shapeSvg(sh, 100, 50)).toContain("font:40pt sans-serif");
  });
});

// OOXML names 187 preset geometries and an unrecognised one falls back to a rectangle. That is
// fair for a decorative shape and wrong for one that means something by its outline: an arrow
// pointing nowhere, a callout with no tail to say what it is about, a connector drawn as the box
// around it. Those presets have to arrive as themselves.
describe("preset geometries", () => {
  const geomOfPreset = (prst: string) =>
    readWorkbook(base(`<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`))
      .sheets[0].shapes![0]!;

  // An arrow's direction IS its message. All four used to share the right-pointing drawing, so a
  // left arrow pointed right, which reads as deliberate rather than as an approximation.
  it("points each arrow the way the file points it", () => {
    const at = (prst: string) => shapeSvg(geomOfPreset(prst), 100, 100);
    const right = at("rightArrow"), left = at("leftArrow"), up = at("upArrow"), down = at("downArrow");
    expect(new Set([right, left, up, down]).size, "four directions, four drawings").toBe(4);
    expect(geomOfPreset("leftArrow").geom).toBe("leftArrow");
    expect(geomOfPreset("upArrow").geom).toBe("upArrow");
    expect(geomOfPreset("downArrow").geom).toBe("downArrow");
  });

  it("gives a callout its tail", () => {
    for (const prst of ["wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "borderCallout1", "accentCallout2"]) {
      const sh = geomOfPreset(prst);
      expect(sh.geom, prst).toMatch(/allout$|^callout$/i);
      const svg = shapeSvg(sh, 100, 60);
      expect(svg, prst).toContain("<path");
      // The tail reaches below the box, which is how it points at the cell it is about.
      const below = [...svg.matchAll(/[ML] [\d.]+ ([\d.]+)/g)].some((m) => Number(m[1]) > 60);
      expect(below, prst).toBe(true);
    }
  });

  it("draws every connector as a line of some kind, never as a box", () => {
    for (const prst of ["straightConnector1", "lineInv", "bentConnector2", "bentConnector3", "curvedConnector3", "curvedConnector5", "arc"]) {
      const svg = shapeSvg(geomOfPreset(prst), 100, 60);
      expect(svg, prst).not.toContain("<rect");
      expect(svg, prst).toMatch(/<line|<path/);
    }
  });

  it("folds the presets that really are one of our shapes", () => {
    const same: [string, string][] = [
      ["flowChartDecision", "diamond"], ["flowChartConnector", "ellipse"],
      ["flowChartInputOutput", "parallelogram"], ["flowChartPreparation", "hexagon"],
      ["flowChartTerminator", "roundRect"], ["flowChartManualOperation", "trapezoidDown"],
      ["star32", "star"], ["irregularSeal1", "star"], ["octagon", "octagon"],
      ["snip2DiagRect", "snipRect"], ["cloudCallout", "ovalCallout"], ["pieWedge", "pie"],
    ];
    for (const [prst, geom] of same) expect(geomOfPreset(prst).geom, prst).toBe(geom);
  });

  it("still falls back to a rectangle for one it does not know", () => {
    expect(geomOfPreset("gear9").geom).toBe("rect");
    expect(geomOfPreset("flowChartProcess").geom).toBe("rect"); // which a process box actually is
  });
});
