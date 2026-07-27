import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, setCellInput, writeWorkbook } from "../../index";
import { setOdsComment, setOdsDataValidation, setOdsHyperlink } from "./write";
import { getCell } from "../../core/model";

const contentOf = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["content.xml"]!);

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:xlink="http://www.w3.org/1999/xlink" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/"`;

function ods(cells: string, extra = ""): Uint8Array {
  const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
    `${extra}<table:table table:name="Sheet1"><table:table-row>${cells}</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
    "content.xml": strToU8(content),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  });
}

describe("ods hyperlink + comment authoring", () => {
  it("reads an existing text:a link and office:annotation", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.com">site</text:a></text:p></table:table-cell>` +
      `<table:table-cell office:value-type="string"><office:annotation><dc:creator>Ada</dc:creator><text:p>look here</text:p></office:annotation><text:p>x</text:p></table:table-cell>`,
    ));
    expect(getCell(wb.sheets[0], 1, 1)?.link).toEqual({ href: "https://example.com" });
    expect(getCell(wb.sheets[0], 1, 2)?.comments?.[0]).toEqual({ author: "Ada", text: "look here" });
  });

  it("reads an internal link and normalises it to Sheet!A1", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="#Sheet1.B2">go</text:a></text:p></table:table-cell>`,
    ));
    expect(getCell(wb.sheets[0], 1, 1)?.link).toEqual({ href: "Sheet1!B2", internal: true });
  });

  it("authors a hyperlink and a note that round-trip", () => {
    const wb = readWorkbook(ods(`<table:table-cell office:value-type="string"><text:p>hello</text:p></table:table-cell>`));
    setOdsHyperlink(wb.sheets[0], 1, 1, { href: "https://a.test" });
    setOdsComment(wb.sheets[0], 1, 2, "please review", "Bob");
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.link?.href).toBe("https://a.test");
    expect(getCell(re.sheets[0], 1, 2)?.comments?.[0]).toEqual({ author: "Bob", text: "please review" });
  });

  it("reads an existing list content-validation", () => {
    const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
      `<table:content-validations><table:content-validation table:name="v1" table:condition='of:cell-content-is-in-list("Yes";"No")' table:allow-empty-cell="true"/></table:content-validations>` +
      `<table:table table:name="Sheet1"><table:table-row>` +
      `<table:table-cell table:content-validation-name="v1" office:value-type="string"><text:p>Yes</text:p></table:table-cell>` +
      `</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const wb = readWorkbook(zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
      "content.xml": strToU8(content),
      "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
    }));
    const dv = wb.sheets[0].validations?.[0];
    expect(dv?.values).toEqual(["Yes", "No"]);
    expect(dv?.ranges[0]).toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
  });

  it("authors a list validation that round-trips", () => {
    const wb = readWorkbook(ods(`<table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell>`));
    setOdsDataValidation(wb, wb.sheets[0], [{ r1: 1, c1: 1, r2: 5, c2: 1 }], { values: ["A", "B", "C"], allowBlank: true });
    const re = readWorkbook(writeWorkbook(wb));
    const dv = re.sheets[0].validations?.[0];
    expect(dv?.values).toEqual(["A", "B", "C"]);
    // The whole authored range carries the validation.
    expect(dv?.ranges.length).toBe(5);
  });

  it("authors typed validations (whole between, text-length, custom) that round-trip", () => {
    const mk = () => readWorkbook(ods(`<table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>`));
    let wb = mk();
    setOdsDataValidation(wb, wb.sheets[0], [{ r1: 1, c1: 1, r2: 3, c2: 1 }], { type: "whole", operator: "between", formula1: "1", formula2: "10", allowBlank: false });
    let content = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(content).toContain("of:cell-content-is-whole-number() and cell-content-is-between(1,10)");
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].validations?.[0]).toMatchObject({ type: "whole", operator: "between", formula1: "1", formula2: "10" });

    wb = mk();
    setOdsDataValidation(wb, wb.sheets[0], [{ r1: 1, c1: 1, r2: 3, c2: 1 }], { type: "textLength", operator: "lessThanOrEqual", formula1: "5" });
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].validations?.[0]).toMatchObject({ type: "textLength", operator: "lessThanOrEqual", formula1: "5" });

    wb = mk();
    setOdsDataValidation(wb, wb.sheets[0], [{ r1: 1, c1: 1, r2: 3, c2: 1 }], { type: "custom", formula1: "[.A1]>0" });
    content = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(content).toContain("of:is-true-formula([.A1]&gt;0)"); // > is XML-escaped in the attribute
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].validations?.[0]).toMatchObject({ type: "custom", formula1: "[.A1]>0" });
  });

  it("editing a cell's value preserves its note position, formatting and validation", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell table:content-validation-name="v1" office:value-type="float" office:value="5">` +
      `<office:annotation svg:x="1.2cm" svg:y="3.4cm" office:name="A1"><dc:creator>Ada</dc:creator><dc:date>2020-01-02T03:04:05</dc:date><text:p>keep me</text:p></office:annotation>` +
      `<text:p>5</text:p></table:table-cell>`,
    ));
    setCellInput(wb.sheets[0], 1, 1, "9"); // a plain value edit, not touching the note/validation
    const xml = contentOf(writeWorkbook(wb));
    expect(xml).toContain('office:value="9"'); // new value written
    expect(xml).toContain('svg:x="1.2cm"'); // annotation position preserved
    expect(xml).toContain("keep me"); // annotation text preserved
    expect(xml).toContain('table:content-validation-name="v1"'); // validation reference preserved
  });

  it("authoring a note keeps the cell value and an existing hyperlink", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string" office:string-value="site"><text:p><text:a xlink:href="https://x.test">site</text:a></text:p></table:table-cell>`,
    ));
    setOdsComment(wb.sheets[0], 1, 1, "review this", "Bob");
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.value).toBe("site");
    expect(getCell(re.sheets[0], 1, 1)?.link?.href).toBe("https://x.test"); // link untouched by note authoring
    expect(getCell(re.sheets[0], 1, 1)?.comments?.[0]).toEqual({ author: "Bob", text: "review this" });
  });

  it("editing one note preserves a second annotation only when notes are untouched", () => {
    // Two notes; editing the VALUE (not the notes) keeps both verbatim.
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="float" office:value="1">` +
      `<office:annotation><text:p>first</text:p></office:annotation>` +
      `<office:annotation><text:p>second</text:p></office:annotation>` +
      `<text:p>1</text:p></table:table-cell>`,
    ));
    setCellInput(wb.sheets[0], 1, 1, "2");
    const xml = contentOf(writeWorkbook(wb));
    expect(xml).toContain("first");
    expect(xml).toContain("second");
    expect((xml.match(/office:annotation/g) ?? []).length).toBeGreaterThanOrEqual(2);
  });

  it("preserves a non-list validation (and does not show it as a dropdown) across a value edit", () => {
    const content = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet>` +
      `<table:content-validations><table:content-validation table:name="w1" table:condition="of:cell-content-is-between(1,10)" table:allow-empty-cell="true"><table:error-message table:message-type="stop"><text:p>1-10 only</text:p></table:error-message></table:content-validation></table:content-validations>` +
      `<table:table table:name="Sheet1"><table:table-row>` +
      `<table:table-cell table:content-validation-name="w1" office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>` +
      `</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const wb = readWorkbook(zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }] as unknown as Uint8Array,
      "content.xml": strToU8(content),
      "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
    }));
    expect(wb.sheets[0].validations).toBeUndefined(); // a between-rule is not a dropdown
    setCellInput(wb.sheets[0], 1, 1, "7");
    const xml = contentOf(writeWorkbook(wb));
    expect(xml).toContain('office:value="7"'); // value edited
    expect(xml).toContain('table:content-validation-name="w1"'); // validation reference preserved
    expect(xml).toContain("cell-content-is-between(1,10)"); // the definition + error message preserved
    expect(xml).toContain("1-10 only");
  });

  it("removes a hyperlink when set to null", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://x.test">x</text:a></text:p></table:table-cell>`,
    ));
    setOdsHyperlink(wb.sheets[0], 1, 1, null);
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.link).toBeUndefined();
  });
});

describe("ods notes and links: what an edit must not throw away", () => {
  it("keeps a cell's other annotations when its note is edited", () => {
    // A cell may carry several annotations; the grid shows the first. Editing that one used to
    // delete every other, losing content the editor never even displayed.
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string">` +
      `<office:annotation svg:x="1cm" svg:y="2cm"><dc:creator>Ada</dc:creator><text:p>first</text:p></office:annotation>` +
      `<office:annotation><dc:creator>Grace</dc:creator><text:p>second</text:p></office:annotation>` +
      `<text:p>x</text:p></table:table-cell>`,
    ));
    expect(getCell(wb.sheets[0], 1, 1)?.comments?.length).toBe(2);
    setOdsComment(wb.sheets[0], 1, 1, "edited", "Ada");
    const out = contentOf(writeWorkbook(wb));
    expect(out).toContain("edited");
    expect(out).toContain("second"); // the annotation nobody touched
    expect(out).toContain("Grace");
    expect(out).toContain('svg:x="1cm"'); // the first one's position is still its own
    const re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0] && getCell(re.sheets[0], 1, 1)?.comments?.map((c) => c.text)).toEqual(["edited", "second"]);
  });

  it("removes only the note it was asked to remove", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string">` +
      `<office:annotation><text:p>first</text:p></office:annotation>` +
      `<office:annotation><text:p>second</text:p></office:annotation>` +
      `<text:p>x</text:p></table:table-cell>`,
    ));
    setOdsComment(wb.sheets[0], 1, 1, null);
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.comments?.map((c) => c.text)).toEqual(["second"]);
  });

  it("keeps how a link opens and looks when its target is changed", () => {
    // Only the href is ours to decide. target-frame-name and the style names say how the link
    // opens and how it is drawn, and rebuilding the anchor from scratch dropped them.
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p>` +
      `<text:a xlink:href="https://old.test" xlink:type="simple" office:target-frame-name="_blank" ` +
      `xlink:show="new" text:style-name="Internet_20_link" text:visited-style-name="Visited_20_link">site</text:a>` +
      `</text:p></table:table-cell>`,
    ));
    setOdsHyperlink(wb.sheets[0], 1, 1, { href: "https://new.test" });
    const out = contentOf(writeWorkbook(wb));
    expect(out).toContain("https://new.test");
    expect(out).not.toContain("https://old.test");
    expect(out).toContain('office:target-frame-name="_blank"');
    expect(out).toContain('xlink:show="new"');
    expect(out).toContain('text:style-name="Internet_20_link"');
    expect(out).toContain('text:visited-style-name="Visited_20_link"');
  });
});

describe("ods part-of-cell links", () => {
  const partial =
    `<table:table-cell office:value-type="string"><text:p>see <text:a xlink:href="https://docs.test">docs</text:a> and ` +
    `<text:a xlink:href="https://spec.test">spec</text:a></text:p></table:table-cell>`;

  it("does not turn the whole cell into a link when the value is edited", () => {
    // ODF can anchor a link to part of a cell; xlsx cannot, so the grid shows the first. Rebuilding
    // the cell around that one used to make ALL of its text a link to it, inventing a link the
    // file never had. The anchors belong to the text that is gone; the value is what survives.
    const wb = readWorkbook(ods(partial));
    const cell = getCell(wb.sheets[0], 1, 1)!;
    expect(cell.link).toEqual({ href: "https://docs.test" }); // still shown in the grid
    expect(cell.linkPartial).toBe(true);
    setCellInput(wb.sheets[0], 1, 1, "plain text now");
    const out = contentOf(writeWorkbook(wb));
    expect(out).toContain("plain text now");
    expect(out).not.toContain("https://docs.test");
    expect(out).not.toContain("<text:a");
  });

  it("keeps every anchor when the text is unchanged", () => {
    // A recalc or a re-entry of the same value must not cost the cell its links.
    const wb = readWorkbook(ods(partial));
    const sheet = wb.sheets[0];
    setCellInput(sheet, 1, 1, "see docs and spec"); // the text it already has
    const out = contentOf(writeWorkbook(wb));
    expect(out).toContain("https://docs.test");
    expect(out).toContain("https://spec.test");
  });

  it("still writes a whole-cell link the user authors on such a cell", () => {
    const wb = readWorkbook(ods(partial));
    setOdsHyperlink(wb.sheets[0], 1, 1, { href: "https://chosen.test" });
    const out = contentOf(writeWorkbook(wb));
    expect(out).toContain("https://chosen.test");
    const re = readWorkbook(writeWorkbook(wb));
    expect(getCell(re.sheets[0], 1, 1)?.link).toEqual({ href: "https://chosen.test" });
    expect(getCell(re.sheets[0], 1, 1)?.linkPartial).toBeFalsy(); // now it does cover the cell
  });

  it("keeps treating a genuine whole-cell link as one", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://whole.test">site</text:a></text:p></table:table-cell>`,
    ));
    expect(getCell(wb.sheets[0], 1, 1)?.linkPartial).toBe(false);
    setCellInput(wb.sheets[0], 1, 1, "renamed");
    const out = contentOf(writeWorkbook(wb));
    expect(out).toContain("https://whole.test"); // the link covered everything, so it follows the text
    expect(out).toContain("renamed");
  });
});

describe("data-validation messages", () => {
  it("reads the help and error messages a file carries", () => {
    const wb = readWorkbook(ods(
      `<table:table-cell table:content-validation-name="v1" office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>`,
      `<table:content-validations><table:content-validation table:name="v1" table:condition="of:cell-content-is-whole-number() and cell-content-is-between(1,10)" table:allow-empty-cell="true">` +
      `<table:help-message table:title="Quantity" table:display="true"><text:p>How many units?</text:p></table:help-message>` +
      `<table:error-message table:message-type="stop" table:title="Out of range" table:display="true"><text:p>Enter 1 to 10.</text:p></table:error-message>` +
      `</table:content-validation></table:content-validations>`,
    ));
    const dv = wb.sheets[0].validations?.[0];
    expect(dv?.promptTitle).toBe("Quantity");
    expect(dv?.promptMessage).toBe("How many units?");
    expect(dv?.errorTitle).toBe("Out of range");
    expect(dv?.errorMessage).toBe("Enter 1 to 10.");
    expect(dv?.errorStyle).toBe("stop");
  });

  it("writes them when a rule is authored, so it can explain itself", () => {
    const wb = readWorkbook(ods(`<table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>`));
    setOdsDataValidation(wb, wb.sheets[0], [{ r1: 1, c1: 1, r2: 1, c2: 1 }], {
      type: "whole", operator: "between", formula1: "1", formula2: "10", allowBlank: true,
      promptTitle: "Quantity", promptMessage: "How many units?",
      errorTitle: "Out of range", errorMessage: "Enter 1 to 10.",
    });
    const out = contentOf(writeWorkbook(wb));
    expect(out).toContain("table:help-message");
    expect(out).toContain('table:title="Quantity"');
    expect(out).toContain("How many units?");
    expect(out).toContain("table:error-message");
    expect(out).toContain('table:message-type="stop"');
    expect(out).toContain("Enter 1 to 10.");
    // And they survive the round-trip.
    const re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].validations?.[0]?.errorMessage).toBe("Enter 1 to 10.");
    expect(re.sheets[0].validations?.[0]?.promptMessage).toBe("How many units?");
  });
});
