import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";

const NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"`;

function ods(body: string, styles = ""): Uint8Array {
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:automatic-styles>${styles}</office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell/></table:table-row>${body}</table:table></office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods drawing shapes", () => {
  it("reads a draw:rect with its graphic style fill/outline into sheet.shapes", () => {
    const style = `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#4472c4" draw:stroke="solid" svg:stroke-color="#222222" svg:stroke-width="0.053cm"/></style:style>`;
    const rect = `<table:shapes><draw:rect draw:style-name="gr1" svg:x="1cm" svg:y="1cm" svg:width="3cm" svg:height="2cm"><text:p>Hi</text:p></draw:rect></table:shapes>`;
    const sh = readWorkbook(ods(rect, style)).sheets[0].shapes ?? [];
    expect(sh.length).toBe(1);
    expect(sh[0].geom).toBe("rect");
    expect(sh[0].fill?.toLowerCase()).toBe("#4472c4");
    expect(sh[0].stroke?.toLowerCase()).toBe("#222222");
    expect(sh[0].text).toBe("Hi");
  });

  it("reads a draw:ellipse and a draw:line", () => {
    const body = `<table:shapes><draw:ellipse svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm"/><draw:line svg:x1="0cm" svg:y1="0cm" svg:x2="3cm" svg:y2="1cm"/></table:shapes>`;
    const sh = readWorkbook(ods(body)).sheets[0].shapes ?? [];
    expect(sh.map((s) => s.geom).sort()).toEqual(["ellipse", "line"]);
  });

  it("authors a new shape into table:shapes and round-trips it", () => {
    const wb = readWorkbook(ods(""));
    (wb.sheets[0].shapes ??= []).push({
      geom: "ellipse",
      anchor: { fromCol: 2, fromRow: 2, fromColOff: 0, fromRowOff: 0, toCol: 5, toRow: 6, toColOff: 0, toRowOff: 0 },
      fill: "#00aa00", stroke: "#003300", strokeWidth: 2, created: true, dirty: true,
    });
    const out = writeWorkbook(wb);
    const content = strFromU8(unzipSync(out)["content.xml"]!);
    expect(content).toContain("<table:shapes>");
    expect(content).toContain("<draw:ellipse");
    expect(content).toContain('draw:fill-color="#00aa00"');
    const re = readWorkbook(out).sheets[0].shapes ?? [];
    expect(re).toHaveLength(1);
    expect(re[0].geom).toBe("ellipse");
    expect(re[0].fill?.toLowerCase()).toBe("#00aa00");
  });

  it("authors a polygon shape as a custom-shape and round-trips its geometry", () => {
    const wb = readWorkbook(ods(""));
    (wb.sheets[0].shapes ??= []).push({
      geom: "diamond",
      anchor: { fromCol: 2, fromRow: 2, fromColOff: 0, fromRowOff: 0, toCol: 5, toRow: 6, toColOff: 0, toRowOff: 0 },
      fill: "#c00000", created: true, dirty: true,
    });
    const out = writeWorkbook(wb);
    const content = strFromU8(unzipSync(out)["content.xml"]!);
    expect(content).toContain("<draw:custom-shape");
    expect(content).toContain('draw:type="diamond"');
    expect(content).toContain("draw:enhanced-path");
    const re = readWorkbook(out).sheets[0].shapes ?? [];
    expect(re[0].geom).toBe("diamond");
  });

  it("patches an existing shape's move + restyle", () => {
    const style = `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#4472c4"/></style:style>`;
    const rect = `<table:shapes><draw:rect draw:style-name="gr1" svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="1cm"/></table:shapes>`;
    const wb = readWorkbook(ods(rect, style));
    const sh = wb.sheets[0].shapes![0];
    sh.odsFrame = { x: 96, y: 96, w: 192, h: 96 }; // a move/resize (96px = 2.54cm)
    sh.fill = "#ff0000";
    sh.dirty = true;
    const content = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(content).toMatch(/svg:x="2\.540cm"/);
    expect(content).toContain('draw:fill-color="#ff0000"');
  });
});
