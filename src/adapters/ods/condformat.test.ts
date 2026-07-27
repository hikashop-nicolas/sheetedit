import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { computeCondVisuals } from "../xlsx/condformat";
import { setOdsCondFormat } from "./write";
import { key } from "../../core/model";
import { makeFormulaEvaluator } from "../../core/recalc";

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
  `xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"`;

function ods(): Uint8Array {
  // Column A uses ce1, whose <style:map> applies ce2 (fill) when value > 3 (standard ODF CF).
  const nc = (v: number, style?: string) => `<table:table-cell ${style ? `table:style-name="${style}" ` : ""}office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
  const rows = [1, 2, 3, 4, 5].map((v) => `<table:table-row>${nc(v, "ce1")}${nc(v)}</table:table-row>`).join("");
  const styles = `<office:automatic-styles>` +
    `<style:style style:name="ce1" style:family="table-cell"><style:map style:condition="cell-content()&gt;3" style:apply-style-name="ce2" style:base-cell-address="Sheet1.A1"/></style:style>` +
    `<style:style style:name="ce2" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffc7ce"/></style:style>` +
    `</office:automatic-styles>`;
  // The colour scale on column B has no style:map form, so it lives in calcext (LibreOffice).
  const cf = `<calcext:conditional-formats>` +
    `<calcext:conditional-format calcext:target-range-address="Sheet1.B1:Sheet1.B5"><calcext:color-scale><calcext:color-scale-entry calcext:type="minimum" calcext:color="#f8696b"/><calcext:color-scale-entry calcext:type="maximum" calcext:color="#63be7b"/></calcext:color-scale></calcext:conditional-format>` +
    `</calcext:conditional-formats>`;
  const content = `<?xml version="1.0"?><office:document-content ${NS}>${styles}<office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1">${rows}</table:table>${cf}` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods conditional formatting read", () => {
  it("reads a cellIs highlight and a colour scale", () => {
    const wb = readWorkbook(ods());
    const cfs = wb.sheets[0].condFormats ?? [];
    expect(cfs.length).toBe(2);
    const vis = computeCondVisuals(wb.sheets[0]);
    // Column A: value > 3 gets the ce1 fill; <= 3 does not.
    expect(vis.get(key(4, 1))?.bg?.toLowerCase()).toBe("#ffc7ce");
    expect(vis.get(key(5, 1))?.bg?.toLowerCase()).toBe("#ffc7ce");
    expect(vis.get(key(2, 1))?.bg).toBeUndefined();
    // Column B: every cell gets a colour-scale fill.
    expect(vis.get(key(1, 2))?.bg).toBeTruthy();
    expect(vis.get(key(5, 2))?.bg).toBeTruthy();
  });

  it("authors a cellIs highlight that round-trips", () => {
    const wb = readWorkbook(ods());
    setOdsCondFormat(wb, wb.sheets[0], [{ r1: 1, c1: 3, r2: 5, c2: 3 }], { kind: "cellIs", operator: "greaterThan", value: "2", fill: "#c6efce" });
    const re = readWorkbook(writeWorkbook(wb));
    const cf = (re.sheets[0].condFormats ?? []).find((c) => c.ranges[0]?.c1 === 3);
    expect(cf?.rules[0].type).toBe("cellIs");
    expect(cf?.rules[0].operator).toBe("greaterThan");
    expect(cf?.rules[0].dxf?.bg?.toLowerCase()).toBe("#c6efce");
  });
});

// The two rule families ODF cannot state in <style:map>, and how LibreOffice really stores them.
// Every string below was taken from LibreOffice's own xlsx -> ods conversion of a workbook holding
// one Excel rule of each kind, rather than invented: text rules exist ONLY in calcext, a formula
// rule is written to BOTH (so it must be counted once), and the two halves name the applied style
// differently - style:map by the escaped style:name, calcext by the display name.
function odsTextCf(): Uint8Array {
  const cell = (v: string, style?: string) =>
    `<table:table-cell ${style ? `table:style-name="${style}" ` : ""}office:value-type="string"><text:p>${v}</text:p></table:table-cell>`;
  const words = ["apple", "banana", "cherry", "apricot"];
  const rows = words.map((w) => `<table:table-row>${cell(w)}${cell(w)}${cell(w, "ce9")}</table:table-row>`).join("");
  const styles = `<office:automatic-styles>` +
    // The applied style, declared the way LibreOffice declares it: an escaped name plus a display name.
    `<style:style style:name="ConditionalStyle_5f_1" style:display-name="ConditionalStyle_1" style:family="table-cell">` +
    `<style:table-cell-properties fo:background-color="#ffff00"/></style:style>` +
    // Column C carries a standard style:map formula rule, mirrored into calcext below.
    `<style:style style:name="ce9" style:family="table-cell">` +
    `<style:map style:condition="is-true-formula(LEN([.C1])&gt;5)" style:apply-style-name="ConditionalStyle_5f_1" style:base-cell-address="Sheet1.C1"/></style:style>` +
    `</office:automatic-styles>`;
  const cf = `<calcext:conditional-formats>` +
    `<calcext:conditional-format calcext:target-range-address="Sheet1.A1:Sheet1.A4">` +
    `<calcext:condition calcext:apply-style-name="ConditionalStyle_1" calcext:value="begins-with(&quot;ap&quot;)" calcext:base-cell-address="Sheet1.A1"/></calcext:conditional-format>` +
    `<calcext:conditional-format calcext:target-range-address="Sheet1.B1:Sheet1.B4">` +
    `<calcext:condition calcext:apply-style-name="ConditionalStyle_1" calcext:value="contains-text(&quot;err&quot;)" calcext:base-cell-address="Sheet1.B1"/></calcext:conditional-format>` +
    `<calcext:conditional-format calcext:target-range-address="Sheet1.C1:Sheet1.C4">` +
    `<calcext:condition calcext:apply-style-name="ConditionalStyle_1" calcext:value="formula-is(LEN([.C1])&gt;5)" calcext:base-cell-address="Sheet1.C1"/></calcext:conditional-format>` +
    `</calcext:conditional-formats>`;
  const content = `<?xml version="1.0"?><office:document-content ${NS}>${styles}<office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1">${rows}</table:table>${cf}` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods conditional formatting: text and formula rules", () => {
  it("renders a begins-with rule, which lives only in calcext", () => {
    const wb = readWorkbook(odsTextCf());
    const vis = computeCondVisuals(wb.sheets[0]);
    expect(vis.get(key(1, 1))?.bg?.toLowerCase()).toBe("#ffff00"); // apple
    expect(vis.get(key(4, 1))?.bg?.toLowerCase()).toBe("#ffff00"); // apricot
    expect(vis.get(key(2, 1))?.bg).toBeUndefined(); // banana
    expect(vis.get(key(3, 1))?.bg).toBeUndefined(); // cherry
  });

  it("renders a contains-text rule", () => {
    const wb = readWorkbook(odsTextCf());
    const vis = computeCondVisuals(wb.sheets[0]);
    expect(vis.get(key(3, 2))?.bg?.toLowerCase()).toBe("#ffff00"); // cherry contains "err"
    expect(vis.get(key(1, 2))?.bg).toBeUndefined();
  });

  it("counts a formula rule once, though it is written to both halves", () => {
    const wb = readWorkbook(odsTextCf());
    const sheet = wb.sheets[0];
    // One rule for column C, from style:map; the calcext formula-is mirror must not add a second.
    const forC = (sheet.condFormats ?? []).filter((c) => c.ranges.some((r) => r.c1 === 3));
    expect(forC.flatMap((c) => c.rules).length).toBe(1);
    expect(forC[0]!.rules[0]!.type).toBe("expression");
    // An expression rule is evaluated by the workbook's formula engine, so it needs one.
    const vis = computeCondVisuals(sheet, { evaluator: makeFormulaEvaluator(wb), sheetName: sheet.name });
    expect(vis.get(key(2, 3))?.bg?.toLowerCase()).toBe("#ffff00"); // banana, 6 chars
    expect(vis.get(key(1, 3))?.bg).toBeUndefined(); // apple, 5
  });
});

// How LibreOffice actually writes a value rule, which is not how the interoperable path assumes.
// It declares the conditional style, puts a <style:map> on an automatic style, and then applies
// that style to NO CELL: the only statement of which cells the rule covers is calcext's
// target-range-address. Reading the standard half alone found nothing and dropped the rule.
function odsLibreOfficeValueCf(): Uint8Array {
  const num = (v: number) => `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
  const rows = [10, 20, 30, 40, 50].map((v) => `<table:table-row>${num(v)}</table:table-row>`).join("");
  const styles = `<office:automatic-styles>` +
    `<style:style style:name="ConditionalStyle_5f_1" style:display-name="ConditionalStyle_1" style:family="table-cell">` +
    `<style:table-cell-properties fo:background-color="#ffff00"/></style:style>` +
    // ce1 carries the standard rule but is applied to nothing, exactly as LibreOffice leaves it.
    `<style:style style:name="ce1" style:family="table-cell">` +
    `<style:map style:condition="cell-content()&gt;20" style:apply-style-name="ConditionalStyle_5f_1" style:base-cell-address="Sheet1.A1"/></style:style>` +
    `</office:automatic-styles>`;
  const cf = `<calcext:conditional-formats>` +
    `<calcext:conditional-format calcext:target-range-address="Sheet1.A1:Sheet1.A5">` +
    `<calcext:condition calcext:apply-style-name="ConditionalStyle_1" calcext:value="&gt;20" calcext:base-cell-address="Sheet1.A1"/>` +
    `</calcext:conditional-format></calcext:conditional-formats>`;
  const content = `<?xml version="1.0"?><office:document-content ${NS}>${styles}<office:body><office:spreadsheet>` +
    `<table:table table:name="Sheet1">${rows}</table:table>${cf}` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods conditional formatting from a LibreOffice-written file", () => {
  it("reads a value rule whose style is applied to no cell", () => {
    const wb = readWorkbook(odsLibreOfficeValueCf());
    const vis = computeCondVisuals(wb.sheets[0]);
    expect(vis.get(key(3, 1))?.bg?.toLowerCase()).toBe("#ffff00"); // 30 > 20
    expect(vis.get(key(5, 1))?.bg?.toLowerCase()).toBe("#ffff00"); // 50 > 20
    expect(vis.get(key(1, 1))?.bg).toBeUndefined(); // 10
    expect(vis.get(key(2, 1))?.bg).toBeUndefined(); // 20, not >
  });

  it("counts the rule once, though the file states it in both halves", () => {
    const wb = readWorkbook(odsLibreOfficeValueCf());
    const rules = (wb.sheets[0].condFormats ?? []).flatMap((c) => c.rules);
    expect(rules.length).toBe(1);
    expect(rules[0]!.operator).toBe("greaterThan");
  });
});
