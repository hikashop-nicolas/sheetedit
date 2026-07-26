import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { readWorkbook } from "../index";
import { getCell } from "./model";
import { recalc } from "./recalc";

// Structured references are what Excel's own UI writes the moment a range becomes a table, so a
// workbook full of them evaluated to nothing at all before. These check the rewrite AND the
// arithmetic that comes out the other side.

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** A workbook with Table1 over A1:C4 on Sheet1 (header + 3 rows), plus the given formula cells. */
function tableWorkbook(formulas: Record<string, string>): Uint8Array {
  const data = [
    ["Day", "Units", "Total"],
    ["Mon", 10, 100],
    ["Tue", 20, 200],
    ["Mon", 5, 50],
  ];
  const rows = data
    .map((row, r) => {
      const cells = row
        .map((v, c) => {
          const ref = `${String.fromCharCode(65 + c)}${r + 1}`;
          return typeof v === "number" ? `<c r="${ref}"><v>${v}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
        })
        .join("");
      return `<row r="${r + 1}">${cells}</row>`;
    })
    .join("");
  const extra = Object.entries(formulas)
    .map(([ref, f], i) => `<row r="${10 + i}"><c r="${ref}"><f>${f.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</f><v>0</v></c></row>`)
    .join("");
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}" xmlns:r="${R}"><sheetData>${rows}${extra}</sheetData><tableParts count="1"><tablePart r:id="rIdT"/></tableParts></worksheet>`),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdT" Type="${R}/table" Target="../tables/table1.xml"/></Relationships>`),
    "xl/tables/table1.xml": strToU8(`<table xmlns="${MAIN}" id="1" name="Table1" displayName="Table1" ref="A1:C4" headerRowCount="1"><tableColumns count="3"><tableColumn id="1" name="Day"/><tableColumn id="2" name="Units"/><tableColumn id="3" name="Total"/></tableColumns></table>`),
  });
}

const valueAt = (bytes: Uint8Array, ref: string): { value?: string; failed?: string } => {
  const wb = readWorkbook(bytes);
  recalc(wb); // reading does not evaluate; the editor recalculates once the workbook is in hand
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const col = m[1]!.split("").reduce((a, ch) => a * 26 + (ch.charCodeAt(0) - 64), 0);
  const cell = getCell(wb.sheets[0]!, Number(m[2]), col);
  return { value: cell?.value, failed: cell?.calcFailed };
};

describe("structured table references", () => {
  it("evaluates a column reference", () => {
    expect(valueAt(tableWorkbook({ E10: "SUM(Table1[Units])" }), "E10").value).toBe("35");
  });

  it("evaluates the criteria form a real workbook uses", () => {
    // SUMIF over two of the table's columns: the shape Excel writes for a lookup total.
    const r = valueAt(tableWorkbook({ E10: 'SUMIF(Table1[Day],"Mon",Table1[Total])' }), "E10");
    expect(r.value).toBe("150");
    expect(r.failed).toBeUndefined();
  });

  it("handles the whole table, the headers, and a column span", () => {
    expect(valueAt(tableWorkbook({ E10: "SUM(Table1)" }), "E10").value).toBe("385"); // data body only
    expect(valueAt(tableWorkbook({ E10: "COUNTA(Table1[#All])" }), "E10").value).toBe("12");
    expect(valueAt(tableWorkbook({ E10: "COUNTA(Table1[#Headers])" }), "E10").value).toBe("3");
    expect(valueAt(tableWorkbook({ E10: "SUM(Table1[[Units]:[Total]])" }), "E10").value).toBe("385");
    expect(valueAt(tableWorkbook({ E10: "SUM(Table1[[#Data],[Total]])" }), "E10").value).toBe("350");
  });

  it("resolves @ against the row the formula is on", () => {
    // D2 sits beside the table's first data row, so [@Units] is B2.
    expect(valueAt(tableWorkbook({ D2: "Table1[@Units]*2" }), "D2").value).toBe("20");
    expect(valueAt(tableWorkbook({ D3: "Table1[[#This Row],[Units]]*2" }), "D3").value).toBe("40");
    // A formula outside the table has no "this row" to mean, so it stays unresolved.
    expect(valueAt(tableWorkbook({ E10: "Table1[@Units]" }), "E10").failed).toBeDefined();
  });

  it("leaves a reference it cannot honestly resolve alone", () => {
    // No totals row is modelled, and an unknown column is not a guess to make: the cell keeps the
    // file's stored value and says it could not be evaluated.
    expect(valueAt(tableWorkbook({ E10: "SUM(Table1[#Totals])" }), "E10").failed).toBeDefined();
    expect(valueAt(tableWorkbook({ E10: "SUM(Table1[Nope])" }), "E10").failed).toBeDefined();
    expect(valueAt(tableWorkbook({ E10: "SUM(Other[Units])" }), "E10").failed).toBeDefined();
  });

  it("does not touch a bracket inside a string literal", () => {
    expect(valueAt(tableWorkbook({ E10: '"Table1[Units]"' }), "E10").value).toBe("Table1[Units]");
  });
});
