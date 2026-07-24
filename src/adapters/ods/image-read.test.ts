import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ` +
  `xmlns:xlink="http://www.w3.org/1999/xlink" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"`;

function ods(frameCell: string, pics: Record<string, Uint8Array> = {}): Uint8Array {
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1"><table:table-row>${frameCell}</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    ...pics,
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods embedded images", () => {
  it("reads a draw:frame > draw:image referencing a Pictures/ part into sheet.images", () => {
    const frame = `<table:table-cell><draw:frame draw:name="Image1" svg:width="2cm" svg:height="1cm" svg:x="0.1cm" svg:y="0.1cm"><draw:image xlink:href="Pictures/a.png" xlink:type="simple"/></draw:frame></table:table-cell>`;
    const wb = readWorkbook(ods(frame, { "Pictures/a.png": new Uint8Array([137, 80, 78, 71, 1, 2, 3]) }));
    const imgs = wb.sheets[0].images ?? [];
    expect(imgs.length).toBe(1);
    expect(imgs[0].dataUri.startsWith("data:image/png;base64,")).toBe(true);
    expect(imgs[0].anchor.fromCol).toBe(1); // anchored to A1
  });

  it("reads an inline office:binary-data image", () => {
    const frame = `<table:table-cell><draw:frame svg:width="1cm" svg:height="1cm"><draw:image><office:binary-data>QUJD</office:binary-data></draw:image></draw:frame></table:table-cell>`;
    const wb = readWorkbook(ods(frame));
    expect(wb.sheets[0].images?.[0].dataUri).toContain("QUJD");
  });

  it("preserves the frame + picture part on save (image untouched)", () => {
    const frame = `<table:table-cell><draw:frame draw:name="Image1" svg:width="2cm" svg:height="1cm"><draw:image xlink:href="Pictures/a.png"/></draw:frame></table:table-cell>`;
    const wb = readWorkbook(ods(frame, { "Pictures/a.png": new Uint8Array([137, 80, 78, 71, 9]) }));
    const out = writeWorkbook(wb);
    const files = unzipSync(out);
    expect(files["Pictures/a.png"]).toBeTruthy(); // media preserved
    expect(strFromU8(files["content.xml"]!)).toContain("Pictures/a.png"); // frame preserved
  });

  it("writes a moved/resized frame back to svg:x/y/width/height", () => {
    const frame = `<table:table-cell><draw:frame draw:name="Image1" svg:width="2cm" svg:height="1cm" svg:x="0.1cm" svg:y="0.1cm" table:end-cell-address="Sheet1.C3"><draw:image xlink:href="Pictures/a.png"/></draw:frame></table:table-cell>`;
    const wb = readWorkbook(ods(frame, { "Pictures/a.png": new Uint8Array([137, 80, 78, 71, 9]) }));
    const im = wb.sheets[0].images![0];
    // a move/resize (as the overlay would commit): 96px = 1in = 2.54cm.
    im.odsFrame = { x: 96, y: 192, w: 288, h: 96 };
    im.dirty = true;
    const content = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(content).toMatch(/svg:x="2\.540cm"/);
    expect(content).toMatch(/svg:y="5\.080cm"/);
    expect(content).toMatch(/svg:width="7\.620cm"/);
    expect(content).toMatch(/svg:height="2\.540cm"/);
    expect(content).not.toContain("end-cell-address"); // explicit size supersedes the two-cell end
  });

  it("replaces an ods image's bytes in place (same extension)", () => {
    const frame = `<table:table-cell><draw:frame draw:name="Image1" svg:width="2cm" svg:height="1cm"><draw:image xlink:href="Pictures/a.png"/></draw:frame></table:table-cell>`;
    const wb = readWorkbook(ods(frame, { "Pictures/a.png": new Uint8Array([1, 2, 3]) }));
    const im = wb.sheets[0].images![0];
    const fresh = new Uint8Array([9, 8, 7, 6]);
    im.replaceBytes = fresh; im.replaceExt = "png"; im.dirty = true;
    const files = unzipSync(writeWorkbook(wb));
    expect(Array.from(files["Pictures/a.png"]!)).toEqual(Array.from(fresh));
  });

  it("replaces an ods image with a different extension: new part + href + manifest", () => {
    const frame = `<table:table-cell><draw:frame draw:name="Image1" svg:width="2cm" svg:height="1cm"><draw:image xlink:href="Pictures/a.png"/></draw:frame></table:table-cell>`;
    const wb = readWorkbook(ods(frame, { "Pictures/a.png": new Uint8Array([1, 2, 3]) }));
    const im = wb.sheets[0].images![0];
    im.replaceBytes = new Uint8Array([255, 216, 255, 224]); im.replaceExt = "jpg"; im.dirty = true;
    const files = unzipSync(writeWorkbook(wb));
    expect(files["Pictures/a.jpg"]).toBeTruthy();
    expect(strFromU8(files["content.xml"]!)).toContain("Pictures/a.jpg");
    expect(strFromU8(files["META-INF/manifest.xml"]!)).toMatch(/Pictures\/a\.jpg/);
  });

  it("re-reads the moved frame's offset back as the anchor position", () => {
    const frame = `<table:table-cell><draw:frame draw:name="Image1" svg:width="2cm" svg:height="1cm" svg:x="0cm" svg:y="0cm"><draw:image xlink:href="Pictures/a.png"/></draw:frame></table:table-cell>`;
    const wb = readWorkbook(ods(frame, { "Pictures/a.png": new Uint8Array([137, 80, 78, 71, 9]) }));
    const im = wb.sheets[0].images![0];
    im.odsFrame = { x: 96, y: 96, w: 192, h: 96 };
    im.dirty = true;
    const re = readWorkbook(writeWorkbook(wb));
    const a = re.sheets[0].images![0].anchor;
    // anchored to A1 (col 1,row 1) + a 96px offset -> the frame renders 96px right/down of A1.
    expect(a.fromCol).toBe(1);
    expect(Math.round(a.fromColOff)).toBe(96);
    expect(Math.round(a.fromRowOff)).toBe(96);
  });
});
