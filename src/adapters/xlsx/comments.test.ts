import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, getCell } from "../../index";

function xlsx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>` +
        `</Relationships>`,
    ),
    "xl/comments1.xml": strToU8(`<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ada</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>Check this</t></r></text></comment></commentList></comments>`),
    "xl/threadedComments/threadedComment1.xml": strToU8(`<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="B2" dT="2020-01-01T00:00:00" personId="{p1}" id="{c1}"><text>Looks good</text></threadedComment></ThreadedComments>`),
    "xl/persons/person1.xml": strToU8(`<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="Grace" id="{p1}"/></personList>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

describe("xlsx comments", () => {
  it("reads legacy and threaded comments onto their cells", () => {
    const wb = readWorkbook(xlsx());
    expect(getCell(wb.sheets[0], 1, 1)?.comments).toEqual([{ author: "Ada", text: "Check this" }]);
    expect(getCell(wb.sheets[0], 2, 2)?.comments).toEqual([{ author: "Grace", text: "Looks good" }]);
  });
});
