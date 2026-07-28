import { describe, expect, it } from "vitest";
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { setCellInput } from "../../core/workbook";
import { setXlsxComment, setXlsxCondFormat, setXlsxDataValidation, setXlsxHyperlink, setXlsxMerge } from "./write";
import { createTable } from "./tables";

// Not a test of behaviour: this AUTHORS a workbook exercising the writers that have no outside
// judge, and records what it meant to write. `npm run check:openpyxl` then reads the file with
// openpyxl - a separate implementation, in another language, that shares no code with this one -
// and checks it sees the same thing.
//
// Round-trips prove we agree with ourselves. LibreOffice judges what it supports, and drops
// slicers entirely. The schemas judge structure but not meaning. This asks a different question:
// does an independent reader understand the file the way we intended it?

const OUT = join(process.cwd(), ".cache", "openpyxl-corpus");

/** A minimal but complete workbook, so the authoring writes into a real package. */
function makeXlsx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
    // A real stylesheet skeleton, not an empty element: every workbook Excel writes has these,
    // and a strict reader needs the cell styles before it will register anything else (openpyxl
    // ignores the differential formats entirely without them, then trips over a rule using one).
    "xl/styles.xml": strToU8(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
      `</styleSheet>`,
    ),
  });
}

describe("openpyxl cross-check corpus", () => {
  it("authors a workbook and records what it means", () => {
    mkdirSync(OUT, { recursive: true });
    const wb = readWorkbook(makeXlsx());
    const sheet = wb.sheets[0]!;

    // A header row and some data, which the table and the rules then describe.
    const rows = [
      ["Product", "Units", "Note"],
      ["apple", "12", "keep"],
      ["banana", "3", "keep"],
      ["cherry", "40", "drop"],
    ];
    rows.forEach((row, r) => row.forEach((v, c) => setCellInput(sheet, r + 1, c + 1, v)));

    createTable(wb, 0, { r1: 1, c1: 1, r2: 4, c2: 3 }, { name: "Sales", hasHeaders: true, style: "TableStyleMedium2" });
    setXlsxCondFormat(wb, sheet, [{ r1: 2, c1: 2, r2: 4, c2: 2 }], { kind: "cellIs", operator: "greaterThan", value: "10", fill: "#ffc7ce" });
    setXlsxCondFormat(wb, sheet, [{ r1: 2, c1: 3, r2: 4, c2: 3 }], { kind: "text", operator: "containsText", text: "keep", fill: "#c6efce" });
    setXlsxDataValidation(sheet, [{ r1: 2, c1: 2, r2: 4, c2: 2 }], {
      type: "whole", operator: "between", formula1: "1", formula2: "100", allowBlank: true,
      promptTitle: "Units", promptMessage: "How many units?",
      errorTitle: "Out of range", errorMessage: "Enter 1 to 100.",
    });
    setXlsxHyperlink(wb, sheet, 2, 1, { href: "https://example.test/apple", tip: "the apple page" });
    setXlsxComment(wb, sheet, 3, 1, "check this one", "Ada");
    // Through the authoring APIs, not by poking the model: a freeze is only written when it is
    // flagged, and a merge is written by the function that also records it. Setting the fields
    // directly produced a file with neither, which is what the cross-check reported first.
    sheet.freeze = { rows: 1, cols: 0 };
    sheet.freezeDirty = true;
    setCellInput(sheet, 6, 1, "merged title");
    setXlsxMerge(sheet, 6, 1, 6, 3, true);

    writeFileSync(join(OUT, "authored.xlsx"), writeWorkbook(wb));

    // What an independent reader should find. Kept beside the file so the checker asserts what
    // this test MEANT, not what it happened to produce.
    const expected = {
      sheet: "Sheet1",
      table: { name: "Sales", ref: "A1:C4" },
      condFormats: [
        { range: "B2:B4", type: "cellIs", operator: "greaterThan", formula: "10" },
        { range: "C2:C4", type: "containsText", text: "keep" },
      ],
      validation: {
        range: "B2:B4", type: "whole", operator: "between", formula1: "1", formula2: "100",
        promptTitle: "Units", prompt: "How many units?",
        errorTitle: "Out of range", error: "Enter 1 to 100.",
      },
      hyperlink: { cell: "A2", target: "https://example.test/apple" },
      comment: { cell: "A3", text: "check this one" },
      freezePanes: "A2",
      merged: "A6:C6",
      values: { A1: "Product", B2: 12, C4: "drop" },
    };
    writeFileSync(join(OUT, "expected.json"), JSON.stringify(expected, null, 2));
    expect(expected.table.name).toBe("Sales");
  });
});
