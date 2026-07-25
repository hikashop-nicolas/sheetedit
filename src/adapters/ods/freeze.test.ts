import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";

const NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`;
const CFG = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

const cell = (t: string): string => `<table:table-cell office:value-type="string"><text:p>${t}</text:p></table:table-cell>`;

function ods(settings?: string): Record<string, Uint8Array> {
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1"><table:table-row>${cell("a")}</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const files: Record<string, Uint8Array> = {
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  };
  if (settings) files["settings.xml"] = strToU8(settings);
  return files;
}

/** A settings.xml with one sheet entry. */
const settingsWith = (items: string): string =>
  `<?xml version="1.0"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="${CFG}">` +
  `<office:settings><config:config-item-set config:name="ooo:view-settings">` +
  `<config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>` +
  `<config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="Sheet1">${items}</config:config-item-map-entry></config:config-item-map-named>` +
  `</config:config-item-map-entry></config:config-item-map-indexed>` +
  `</config:config-item-set></office:settings></office:document-settings>`;

const int = (name: string, v: number): string => `<config:config-item config:name="${name}" config:type="int">${v}</config:config-item>`;
const settings = (b: Uint8Array): string => strFromU8(unzipSync(b)["settings.xml"]!);

describe("ods frozen panes", () => {
  it("reads the split mode and position from settings.xml", () => {
    const wb = readWorkbook(zipSync(ods(settingsWith(
      int("HorizontalSplitMode", 2) + int("VerticalSplitMode", 2) + int("HorizontalSplitPosition", 2) + int("VerticalSplitPosition", 1)))));
    expect(wb.sheets[0]!.freeze).toEqual({ rows: 1, cols: 2 });
  });

  it("reads a draggable split through PositionBottom, not its pixel offset", () => {
    // Mode 1's position is a view-pixel offset; the trailing pane's first line is the reliable part.
    const wb = readWorkbook(zipSync(ods(settingsWith(
      int("VerticalSplitMode", 1) + int("VerticalSplitPosition", 900) + int("PositionBottom", 3)))));
    expect(wb.sheets[0]!.freeze).toEqual({ rows: 3, cols: 0 });
    expect(wb.sheets[0]!.paneSplit).toBe(true);
  });

  it("ignores a mode-1 split that names no trailing pane", () => {
    const wb = readWorkbook(zipSync(ods(settingsWith(int("VerticalSplitMode", 1) + int("VerticalSplitPosition", 900)))));
    expect(wb.sheets[0]!.freeze).toBeUndefined();
  });

  it("leaves an untouched split's settings byte-identical", () => {
    const src = ods(settingsWith(int("VerticalSplitMode", 1) + int("VerticalSplitPosition", 900) + int("PositionBottom", 3)));
    expect(settings(writeWorkbook(readWorkbook(zipSync(src))))).toBe(strFromU8(src["settings.xml"]!));
  });

  it("turns a MOVED split into a frozen boundary", () => {
    const wb = readWorkbook(zipSync(ods(settingsWith(
      int("VerticalSplitMode", 1) + int("VerticalSplitPosition", 900) + int("PositionBottom", 3)))));
    const s = wb.sheets[0]!;
    expect(s.paneSplit).toBe(true);
    Object.assign(s, { freeze: { rows: 2, cols: 0 }, freezeDirty: true });
    const xml = settings(writeWorkbook(wb));
    expect(xml).toMatch(/VerticalSplitMode"[^>]*>2</);
    expect(xml).toMatch(/VerticalSplitPosition"[^>]*>2</);
    expect(s.paneSplit).toBeUndefined();
    expect(readWorkbook(writeWorkbook(wb)).sheets[0]!.freeze).toEqual({ rows: 2, cols: 0 });
  });

  it("leaves settings.xml alone when the freeze was not changed", () => {
    const src = ods(settingsWith(int("HorizontalSplitMode", 2) + int("HorizontalSplitPosition", 1)));
    const bytes = zipSync(src);
    expect(settings(writeWorkbook(readWorkbook(bytes)))).toBe(strFromU8(src["settings.xml"]!));
  });

  it("writes a freeze into an existing sheet entry", () => {
    const wb = readWorkbook(zipSync(ods(settingsWith(int("HorizontalSplitMode", 0)))));
    Object.assign(wb.sheets[0]!, { freeze: { rows: 1, cols: 2 }, freezeDirty: true });
    const xml = settings(writeWorkbook(wb));
    expect(xml).toMatch(/HorizontalSplitMode"[^>]*>2</);
    expect(xml).toMatch(/VerticalSplitMode"[^>]*>2</);
    expect(xml).toMatch(/HorizontalSplitPosition"[^>]*>2</);
    expect(xml).toMatch(/VerticalSplitPosition"[^>]*>1</);
    expect(xml).toMatch(/PositionBottom"[^>]*>1</);
  });

  it("creates settings.xml (and its manifest entry) when the file had none", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { freeze: { rows: 1, cols: 0 }, freezeDirty: true });
    const out = unzipSync(writeWorkbook(wb));
    expect(out["settings.xml"]).toBeTruthy();
    expect(strFromU8(out["settings.xml"]!)).toContain(`config:name="Sheet1"`);
    expect(strFromU8(out["META-INF/manifest.xml"]!)).toContain(`manifest:full-path="settings.xml"`);
  });

  it("clears the modes when unfreezing", () => {
    const wb = readWorkbook(zipSync(ods(settingsWith(
      int("HorizontalSplitMode", 2) + int("VerticalSplitMode", 2) + int("HorizontalSplitPosition", 2) + int("VerticalSplitPosition", 1)))));
    Object.assign(wb.sheets[0]!, { freeze: undefined, freezeDirty: true });
    const xml = settings(writeWorkbook(wb));
    expect(xml).toMatch(/HorizontalSplitMode"[^>]*>0</);
    expect(xml).toMatch(/VerticalSplitMode"[^>]*>0</);
  });

  it("round-trips a freeze it wrote", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { freeze: { rows: 2, cols: 1 }, freezeDirty: true });
    expect(readWorkbook(writeWorkbook(wb)).sheets[0]!.freeze).toEqual({ rows: 2, cols: 1 });
  });
});
