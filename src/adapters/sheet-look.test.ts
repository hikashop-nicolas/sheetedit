import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { getCell, readWorkbook, writeWorkbook } from "../index";
import { setXlsxCellStyle } from "./xlsx/styles";

// How a sheet asks to be looked at, as opposed to what it holds: the gridlines turned off for a
// sheet laid out as a document, a coloured tab, and text set sideways to fit a narrow column.
// None of it was read, so a form-like sheet arrived covered in a grid its author had removed and
// its rotated headings were clipped to the first letter or two.

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function book(sheetXml: string, styles = `<styleSheet xmlns="${SS}"/>`): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(sheetXml),
    "xl/styles.xml": strToU8(styles),
  });
}

const sheet = (head: string, rows = "") =>
  `<worksheet xmlns="${SS}" xmlns:r="${R}">${head}<sheetData>${rows}</sheetData></worksheet>`;

describe("the sheet's own view settings", () => {
  it("reads showGridLines=0", () => {
    const wb = readWorkbook(book(sheet(`<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>`)));
    expect(wb.sheets[0].hideGridLines).toBe(true);
  });

  it("leaves the flag off when the view says nothing, or says to show them", () => {
    expect(readWorkbook(book(sheet(`<sheetViews><sheetView workbookViewId="0"/></sheetViews>`))).sheets[0].hideGridLines).toBeUndefined();
    expect(readWorkbook(book(sheet(`<sheetViews><sheetView showGridLines="1" workbookViewId="0"/></sheetViews>`))).sheets[0].hideGridLines).toBeUndefined();
  });

  it("reads the tab colour, literal or from the theme", () => {
    expect(readWorkbook(book(sheet(`<sheetPr><tabColor rgb="FFFFFF00"/></sheetPr>`))).sheets[0].tabColor).toBe("#ffff00");
    // A theme-indexed tab colour resolves through the palette rather than being dropped.
    const themed = readWorkbook(book(sheet(`<sheetPr><tabColor theme="4"/></sheetPr>`))).sheets[0].tabColor;
    expect(themed).toMatch(/^#[0-9a-f]{6}$/);
  });
});

describe("rotated cell text", () => {
  const styles = (alignment: string) =>
    `<styleSheet xmlns="${SS}"><fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment ${alignment}/></xf></cellXfs></styleSheet>`;
  const withStyle = (alignment: string) =>
    readWorkbook(book(sheet("", `<row r="1"><c r="A1" s="1" t="str"><v>cadre</v></c></row>`), styles(alignment)));

  it("reads the rotation angle", () => {
    expect(getCell(withStyle(`textRotation="90"`).sheets[0], 1, 1)?.cellStyle?.rot).toBe(90);
    expect(getCell(withStyle(`textRotation="255"`).sheets[0], 1, 1)?.cellStyle?.rot).toBe(255); // stacked
  });

  it("treats no rotation as none", () => {
    expect(getCell(withStyle(`horizontal="center"`).sheets[0], 1, 1)?.cellStyle?.rot).toBeUndefined();
    expect(getCell(withStyle(`textRotation="0"`).sheets[0], 1, 1)?.cellStyle?.rot).toBeUndefined();
  });

  // The style writer rebuilds the xf from scratch, so anything it does not carry is lost the
  // moment the user bolds the cell. Rotation has no toolbar control, which is exactly why it has
  // to survive changes made through the ones that do.
  it("survives an unrelated style change", () => {
    const wb = withStyle(`horizontal="center" vertical="center" textRotation="90"`);
    setXlsxCellStyle(wb, wb.sheets[0], getCell(wb.sheets[0], 1, 1)!, { bold: true });
    const re = readWorkbook(writeWorkbook(wb));
    const cs = getCell(re.sheets[0], 1, 1)?.cellStyle;
    expect(cs?.bold).toBe(true);
    expect(cs?.rot).toBe(90);
  });
});

// The same three settings on the ODF side, so a model field the grid honours does not quietly
// mean "xlsx only". ODF keeps the view settings in settings.xml and the rotation in the style.
describe("the ODF equivalents", () => {
  const OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
  const CFG = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
  const TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
  const STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
  const TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

  function ods(opts: { showGrid?: string; rotation?: string } = {}): Uint8Array {
    const settings = opts.showGrid == null ? {} : {
      "settings.xml": strToU8(`<office:document-settings xmlns:office="${OFFICE}" xmlns:config="${CFG}"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="Sheet1"><config:config-item config:name="ShowGrid" config:type="boolean">${opts.showGrid}</config:config-item></config:config-item-map-entry></config:config-item-map-named></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`),
    };
    const cellStyle = opts.rotation == null ? "" :
      `<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties style:rotation-angle="${opts.rotation}"/></style:style>`;
    const cellAttr = opts.rotation == null ? "" : ` table:style-name="ce1"`;
    return zipSync({
      "mimetype": strToU8("application/vnd.oasis.opendocument.spreadsheet"),
      "content.xml": strToU8(`<office:document-content xmlns:office="${OFFICE}" xmlns:table="${TABLE}" xmlns:style="${STYLE}" xmlns:text="${TEXT}"><office:automatic-styles>${cellStyle}</office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell${cellAttr} office:value-type="string"><text:p>cadre</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`),
      ...settings,
    });
  }

  it("reads ShowGrid=false as gridlines off", () => {
    expect(readWorkbook(ods({ showGrid: "false" })).sheets[0].hideGridLines).toBe(true);
    expect(readWorkbook(ods({ showGrid: "true" })).sheets[0].hideGridLines).toBeUndefined();
    expect(readWorkbook(ods()).sheets[0].hideGridLines).toBeUndefined();
  });

  // ODF counts anticlockwise all the way round; OOXML flips to clockwise past 90. The model uses
  // OOXML's terms, so 270 (a quarter turn clockwise) has to arrive as 180, not as 270.
  it("folds ODF's anticlockwise angle into the model's OOXML one", () => {
    expect(getCell(readWorkbook(ods({ rotation: "90" })).sheets[0], 1, 1)?.cellStyle?.rot).toBe(90);
    expect(getCell(readWorkbook(ods({ rotation: "270" })).sheets[0], 1, 1)?.cellStyle?.rot).toBe(180);
    expect(getCell(readWorkbook(ods({ rotation: "0" })).sheets[0], 1, 1)?.cellStyle?.rot).toBeUndefined();
  });
});
