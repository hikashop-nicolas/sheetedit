import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { groupLines, setGroupCollapsed } from "../../core/outline";

const NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`;

const cell = (t: string): string => `<table:table-cell office:value-type="string"><text:p>${t}</text:p></table:table-cell>`;
const row = (t: string, attrs = ""): string => `<table:table-row${attrs}>${cell(t)}</table:table-row>`;

function ods(body: string, cols = ""): Uint8Array {
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1">${cols}${body}</table:table></office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

const content = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["content.xml"]!);
/** The document body, without the XML declaration the serializer normalises. */
const body = (bytes: Uint8Array): string => content(bytes).replace(/^<\?xml[^>]*\?>\s*/, "");

describe("ods outline grouping", () => {
  it("reads the nesting depth of table-row-group as the outline level", () => {
    const wb = readWorkbook(ods(
      row("a") +
      `<table:table-row-group>${row("b")}<table:table-row-group>${row("c")}</table:table-row-group>${row("d")}</table:table-row-group>` +
      row("e")));
    const s = wb.sheets[0]!;
    expect([...s.rowOutline!.entries()]).toEqual([[2, 1], [3, 2], [4, 1]]);
    expect(s.hiddenRows).toBeUndefined();
  });

  it("treats a group with display=false as collapsed, hiding its rows", () => {
    const wb = readWorkbook(ods(row("a") + `<table:table-row-group table:display="false">${row("b")}${row("c")}</table:table-row-group>` + row("d")));
    const s = wb.sheets[0]!;
    expect([...s.hiddenRows!].sort((a, b) => a - b)).toEqual([2, 3]);
    expect([...s.rowOutline!.keys()].sort((a, b) => a - b)).toEqual([2, 3]);
  });

  it("reads column groups from table-column-group nesting", () => {
    const cols = `<table:table-column/><table:table-column-group><table:table-column table:number-columns-repeated="2"/></table:table-column-group>`;
    const wb = readWorkbook(ods(row("a"), cols));
    expect([...wb.sheets[0]!.colOutline!.entries()]).toEqual([[2, 1], [3, 1]]);
  });

  it("writes grouped rows back into nested row groups", () => {
    const wb = readWorkbook(ods(row("a") + row("b") + row("c") + row("d")));
    const s = wb.sheets[0]!;
    groupLines(s, "row", 2, 3);
    groupLines(s, "row", 3, 3);
    s.odsDirty = true;
    const xml = content(writeWorkbook(wb));
    expect(xml).toContain("table:table-row-group");
    // Row 3 is one level deeper, so its group nests inside the outer one.
    expect(xml).toMatch(/<table:table-row-group>[\s\S]*<table:table-row-group>/);
    const again = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect([...again.rowOutline!.entries()]).toEqual([[2, 1], [3, 2]]);
  });

  it("marks a collapsed group with display=false and hides its rows", () => {
    const wb = readWorkbook(ods(row("a") + row("b") + row("c") + row("d")));
    const s = wb.sheets[0]!;
    groupLines(s, "row", 2, 3);
    setGroupCollapsed(s, "row", 2, 1, true, 4);
    s.odsDirty = true;
    const xml = content(writeWorkbook(wb));
    expect(xml).toMatch(/<table:table-row-group[^>]*table:display="false"/);
    const again = readWorkbook(writeWorkbook(wb)).sheets[0]!;
    expect([...again.hiddenRows!].sort((a, b) => a - b)).toEqual([2, 3]);
  });

  it("leaves an untouched sheet's XML alone", () => {
    const src = ods(row("a") + `<table:table-row-group>${row("b")}</table:table-row-group>` + row("c"));
    expect(body(writeWorkbook(readWorkbook(src)))).toBe(body(src));
  });
});
