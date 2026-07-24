import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../../index";
import { writeXlsx } from "./write";
import { writeWorkbook } from "../../core/workbook";
import { serializeXml } from "../../core/model";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const png = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3, 4]);

function base(anchorXml: string): Uint8Array {
  const drawing = `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}">${anchorXml}</xdr:wsDr>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheetData/><drawing r:id="rId1"/></worksheet>`),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`),
    "xl/drawings/drawing1.xml": strToU8(drawing),
    "xl/drawings/_rels/drawing1.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/image" Target="../media/image1.png"/></Relationships>`),
    "xl/media/image1.png": png,
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

const twoCell = `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Pic"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
const oneCell = `<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="914400" cy="914400"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Pic"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;

describe("embedded image writer", () => {
  it("writes a moved/resized twoCellAnchor picture back into the drawing", () => {
    const wb = readWorkbook(base(twoCell));
    const im = wb.sheets[0].images![0];
    im.anchor = { fromCol: 3, fromRow: 4, fromColOff: 5, fromRowOff: 6, toCol: 7, toRow: 10, toColOff: 8, toRowOff: 9 };
    im.dirty = true;
    writeXlsx(wb);
    const draw = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    // from cell C5 (col 2,row 4 zero-based), to cell G11 (col 6,row 10 zero-based).
    expect(draw).toMatch(/<xdr:from><xdr:col>2<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>3<\/xdr:row>/);
    expect(draw).toMatch(/<xdr:to><xdr:col>6<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>9<\/xdr:row>/);
    // round-trips back to the new anchor
    const re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].images![0].anchor).toMatchObject({ fromCol: 3, fromRow: 4, toCol: 7, toRow: 10 });
  });

  it("converts a oneCellAnchor picture to a twoCellAnchor when moved/resized", () => {
    const wb = readWorkbook(base(oneCell));
    const im = wb.sheets[0].images![0];
    expect(im.anchor.toCol).toBe(2); // read as same cell (col 1 -> 1-based 2)
    im.anchor = { fromCol: 2, fromRow: 2, fromColOff: 0, fromRowOff: 0, toCol: 6, toRow: 9, toColOff: 0, toRowOff: 0 };
    im.dirty = true;
    writeXlsx(wb);
    const draw = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    expect(draw).toContain("<xdr:twoCellAnchor");
    expect(draw).not.toContain("oneCellAnchor");
    expect(draw).toContain('r:embed="rId1"'); // the picture content survived the conversion
    const re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].images![0].anchor).toMatchObject({ fromCol: 2, toCol: 6, toRow: 9 });
  });

  it("replaces an image's bytes in place when the extension is unchanged", () => {
    const wb = readWorkbook(base(twoCell));
    const im = wb.sheets[0].images![0];
    const fresh = Uint8Array.from([137, 80, 78, 71, 9, 9, 9, 9]);
    im.replaceBytes = fresh; im.replaceExt = "png"; im.dirty = true;
    writeXlsx(wb);
    expect(Array.from(wb.files["xl/media/image1.png"])).toEqual(Array.from(fresh));
  });

  it("replaces with a different extension: new part, retargeted rel, content type", () => {
    const wb = readWorkbook(base(twoCell));
    const im = wb.sheets[0].images![0];
    im.replaceBytes = Uint8Array.from([255, 216, 255, 224]); im.replaceExt = "jpg"; im.dirty = true;
    writeXlsx(wb);
    expect(wb.files["xl/media/image1.jpg"]).toBeTruthy();
    expect(new TextDecoder().decode(wb.files["xl/drawings/_rels/drawing1.xml.rels"])).toContain("image1.jpg");
    expect(new TextDecoder().decode(wb.files["[Content_Types].xml"])).toMatch(/Extension="jpg"/);
    // read back: the image resolves to the new jpeg part
    const re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].images![0].dataUri.startsWith("data:image/jpeg")).toBe(true);
  });

  it("leaves an untouched image's drawing part byte-identical", () => {
    const wb = readWorkbook(base(twoCell));
    const before = new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"]);
    writeXlsx(wb); // nothing dirty
    expect(new TextDecoder().decode(wb.files["xl/drawings/drawing1.xml"])).toBe(before);
    void serializeXml;
  });
});
