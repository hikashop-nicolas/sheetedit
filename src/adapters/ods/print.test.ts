import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { odsLenToMm } from "./print-read";

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"` +
  ` xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
  ` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"`;

/** A one-sheet ods; `styles` goes into styles.xml, `tableAttrs` onto the table. */
function ods(opts: { tableAttrs?: string; contentStyles?: string; stylesXml?: string; rows?: number } = {}): Record<string, Uint8Array> {
  const row = `<table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row>`;
  const content =
    `<?xml version="1.0"?><office:document-content ${NS}><office:automatic-styles>${opts.contentStyles ?? ""}</office:automatic-styles>` +
    `<office:body><office:spreadsheet><table:table table:name="Sheet1"${opts.tableAttrs ? " " + opts.tableAttrs : ""}>` +
    `<table:table-column/>${row.repeat(opts.rows ?? 3)}</table:table></office:spreadsheet></office:body></office:document-content>`;
  return {
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(content),
    "styles.xml": strToU8(opts.stylesXml ?? `<?xml version="1.0"?><office:document-styles ${NS}><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>`),
    "META-INF/manifest.xml": strToU8(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/></manifest:manifest>`),
  };
}

/** A styles.xml with one page layout + master page, and a table style pointing at it. */
const stylesWith = (layoutProps: string, header = "", footer = "", headerStyle = ""): string =>
  `<?xml version="1.0"?><office:document-styles ${NS}><office:automatic-styles>` +
  `<style:page-layout style:name="pm1"><style:page-layout-properties ${layoutProps}/>${headerStyle}</style:page-layout>` +
  `</office:automatic-styles><office:master-styles>` +
  `<style:master-page style:name="mp1" style:page-layout-name="pm1">${header}${footer}</style:master-page>` +
  `</office:master-styles></office:document-styles>`;
const TABLE_STYLE = `<style:style style:name="ta1" style:family="table" style:master-page-name="mp1"/>`;

const content = (b: Uint8Array): string => strFromU8(unzipSync(b)["content.xml"]!);
const styles = (b: Uint8Array): string => strFromU8(unzipSync(b)["styles.xml"]!);

describe("ods lengths", () => {
  it("converts every unit ODF allows", () => {
    expect(odsLenToMm("10mm")).toBeCloseTo(10);
    expect(odsLenToMm("1cm")).toBeCloseTo(10);
    expect(odsLenToMm("1in")).toBeCloseTo(25.4);
    expect(odsLenToMm("72pt")).toBeCloseTo(25.4);
    expect(odsLenToMm("nonsense")).toBeUndefined();
  });
});

describe("ods print setup", () => {
  it("reads the page layout through the table style", () => {
    const wb = readWorkbook(zipSync(ods({
      tableAttrs: `table:style-name="ta1"`,
      contentStyles: TABLE_STYLE,
      stylesXml: stylesWith(`fo:page-width="297mm" fo:page-height="210mm" style:print-orientation="landscape" style:table-centering="horizontal" style:print="grid headers objects"`),
    })));
    const p = wb.sheets[0]!.printSetup!;
    expect(p.orientation).toBe("landscape");
    expect(p.paperSize).toBe(9); // A4, recognised from the dimensions
    expect(p.horizontalCentered).toBe(true);
    expect(p.verticalCentered).toBe(false);
    expect(p.gridLines).toBe(true);
    expect(p.headings).toBe(true);
  });

  it("adds the header block back into the margin, which ODF states separately", () => {
    // ODF: page margin 0.3in + a 0.45in header block. xlsx: top 0.75in with the header at 0.3in.
    const wb = readWorkbook(zipSync(ods({
      tableAttrs: `table:style-name="ta1"`,
      contentStyles: TABLE_STYLE,
      stylesXml: stylesWith(
        `fo:margin-top="0.3in" fo:margin-bottom="0.3in" fo:margin-left="0.5in" fo:margin-right="0.5in"`,
        "", "",
        `<style:header-style><style:header-footer-properties fo:min-height="0.45in"/></style:header-style><style:footer-style><style:header-footer-properties fo:min-height="0.45in"/></style:footer-style>`),
    })));
    expect(wb.sheets[0]!.printSetup!.margins).toEqual({ left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });
  });

  it("turns the header regions back into field codes", () => {
    const header = `<style:header><style:region-left><text:p>Left</text:p></style:region-left><style:region-center><text:p>Page <text:page-number>1</text:page-number> of <text:page-count>9</text:page-count></text:p></style:region-center></style:header>`;
    const wb = readWorkbook(zipSync(ods({ tableAttrs: `table:style-name="ta1"`, contentStyles: TABLE_STYLE, stylesXml: stylesWith("", header) })));
    expect(wb.sheets[0]!.printSetup!.header).toEqual({ left: "Left", center: "Page &P of &N" });
  });

  it("reads the print ranges off the table", () => {
    const wb = readWorkbook(zipSync(ods({ tableAttrs: `table:print-ranges="Sheet1.A1:Sheet1.C5"` })));
    expect(wb.sheets[0]!.printSetup!.printArea).toEqual([{ r1: 1, c1: 1, r2: 5, c2: 3 }]);
  });

  it("reads a page break from the row style", () => {
    const cs = `<style:style style:name="ro9" style:family="table-row"><style:table-row-properties fo:break-before="page"/></style:style>`;
    const files = ods({ contentStyles: cs });
    // Put the break style on the second row.
    const c = strFromU8(files["content.xml"]!).replace("<table:table-row>", "<table:table-row>").split("<table:table-row>");
    const rebuilt = c[0] + "<table:table-row>" + c[1] + `<table:table-row table:style-name="ro9">` + c.slice(2).join("<table:table-row>");
    files["content.xml"] = strToU8(rebuilt);
    expect(readWorkbook(zipSync(files)).sheets[0]!.printSetup!.rowBreaks).toEqual([2]);
  });

  it("a sheet with no page setup has none", () => {
    expect(readWorkbook(zipSync(ods())).sheets[0]!.printSetup).toBeUndefined();
  });

  it("writes a page layout and points the table style at its master page", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { printSetup: { orientation: "landscape", paperSize: 9, gridLines: true }, printDirty: true });
    const out = writeWorkbook(wb);
    expect(styles(out)).toMatch(/style:print-orientation="landscape"/);
    expect(styles(out)).toMatch(/fo:page-width="297mm"/); // A4 turned round
    expect(styles(out)).toMatch(/style:print="[^"]*grid[^"]*"/);
    // The table has to reach the master page through its own style, or nothing applies.
    const masterName = /<style:master-page[^>]*style:name="([^"]*)"/.exec(styles(out))![1]!;
    const styleName = /<table:table [^>]*table:style-name="([^"]*)"/.exec(content(out))![1]!;
    expect(content(out)).toMatch(new RegExp(`style:name="${styleName}"[^>]*style:master-page-name="${masterName}"`));
  });

  it("splits the margins the way ODF states them", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { printSetup: { margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } }, printDirty: true });
    const out = styles(writeWorkbook(wb));
    expect(out).toMatch(/fo:margin-top="7.62mm"/); // the header distance, not the whole top margin
    expect(out).toMatch(/fo:min-height="11.43mm"/); // the block the header sits in (0.75 - 0.3)
    // LibreOffice derives an OOXML margin from content height + this spacing, ignoring min-height.
    expect(out).toMatch(/fo:margin-bottom="7.9mm"/);
  });

  it("writes the print ranges and a page break", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { printSetup: { printArea: [{ r1: 1, c1: 1, r2: 5, c2: 3 }], rowBreaks: [2] }, printDirty: true });
    const out = content(writeWorkbook(wb));
    expect(out).toMatch(/table:print-ranges="Sheet1.A1:Sheet1.C5"/);
    expect(out).toMatch(/fo:break-before="page"/);
  });

  it("clears the print ranges when the area is dropped", () => {
    const wb = readWorkbook(zipSync(ods({ tableAttrs: `table:print-ranges="Sheet1.A1:Sheet1.C5"` })));
    Object.assign(wb.sheets[0]!, { printSetup: { orientation: "portrait" }, printDirty: true });
    expect(content(writeWorkbook(wb))).not.toContain("print-ranges");
  });

  it("round-trips a page setup it wrote", () => {
    const wb = readWorkbook(zipSync(ods()));
    const setup = {
      orientation: "landscape" as const, paperSize: 9,
      margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
      gridLines: true, headings: true, horizontalCentered: true, verticalCentered: false,
      header: { left: "L", center: "C" }, footer: { center: "Page &P" },
      printArea: [{ r1: 1, c1: 1, r2: 5, c2: 3 }], rowBreaks: [2], firstPageNumber: 3,
    };
    Object.assign(wb.sheets[0]!, { printSetup: setup, printDirty: true });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.printSetup!;
    expect(back).toMatchObject(setup);
  });

  it("keeps a scale when no fit-to is in force, and drops it when one is", () => {
    const wb = readWorkbook(zipSync(ods()));
    Object.assign(wb.sheets[0]!, { printSetup: { scale: 80 }, printDirty: true });
    expect(styles(writeWorkbook(wb))).toMatch(/style:scale-to="80%"/);

    const wb2 = readWorkbook(zipSync(ods()));
    Object.assign(wb2.sheets[0]!, { printSetup: { scale: 80, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }, printDirty: true });
    const out = styles(writeWorkbook(wb2));
    expect(out).toMatch(/style:scale-to-X="1"/);
    expect(out).not.toContain("scale-to=");
  });
});
