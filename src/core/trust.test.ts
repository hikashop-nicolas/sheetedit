import { beforeAll, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readCsv } from "../adapters/csv/read";
import { createSheetEditor } from "./editor";
import { getCell } from "./model";
import { recalc } from "./recalc";
import { readWorkbook, writeWorkbook } from "./workbook";

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

describe("unknown sheet names", () => {
  it("give #REF! instead of silently reading sheet 1", () => {
    const wb = readCsv("42,=Missing!A1,=SUM(Nope!A1:A9)\n");
    recalc(wb);
    const s = wb.sheets[0]!;
    expect(getCell(s, 1, 2)?.value).toBe("#REF!");
    expect(getCell(s, 1, 3)?.value).toBe("#REF!");
  });

  it("existing sheet names keep working across sheets", () => {
    const wb = readCsv("7,=Sheet1!A1*2\n"); // csv single sheet is named Sheet1
    recalc(wb);
    expect(getCell(wb.sheets[0]!, 1, 2)?.value).toBe("14");
  });
});

describe("circular references", () => {
  it("are flagged on every cell of the cycle", () => {
    const wb = readCsv("=B1+1,=A1+1,=A1+B1\n");
    recalc(wb);
    const s = wb.sheets[0]!;
    expect(getCell(s, 1, 1)?.calcFailed).toBe("circular");
    expect(getCell(s, 1, 2)?.calcFailed).toBe("circular");
    // C1 depends on the cycle but is not part of it.
    expect(getCell(s, 1, 3)?.calcFailed).toBeUndefined();
  });
});

describe("unevaluable formulas keep the cached value but say so", () => {
  const SHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <sheetData>
  <row r="1"><c r="A1"><v>5</v></c><c r="B1"><f>NOTAREALFN(A1)</f><v>42</v></c></row>
 </sheetData>
</worksheet>`;
  const makeXlsx = () =>
    zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "xl/workbook.xml": strToU8(
        `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      ),
      "xl/_rels/workbook.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      ),
      "xl/worksheets/sheet1.xml": strToU8(SHEET),
    });

  it("unknown function: cached value stays, calcFailed set, file unchanged on save", () => {
    const wb = readWorkbook(makeXlsx());
    recalc(wb);
    const cell = getCell(wb.sheets[0]!, 1, 2)!;
    expect(cell.value).toBe("42"); // the cached result is what other apps show too
    expect(cell.calcFailed).toBe("name");
    // Saving keeps the formula and its cached value; calcFailed never serializes.
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]!);
    expect(xml).toContain("NOTAREALFN(A1)");
    expect(xml).toContain("<v>42</v>");
    expect(xml).not.toContain("calcFailed");
  });

  it("the badge clears once the formula becomes computable", () => {
    const wb = readCsv("=NOTAREALFN(1),2\n");
    recalc(wb);
    const s = wb.sheets[0]!;
    const cell = getCell(s, 1, 1)!;
    expect(cell.calcFailed).toBeTruthy();
    cell.formula = "B1*3";
    cell.fDirty = true;
    recalc(wb);
    expect(cell.calcFailed).toBeUndefined();
    expect(cell.value).toBe("6");
  });
});

describe("grid badge", () => {
  it("marks the cell and explains in the tooltip", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("=B1+1,=A1+1\n"), { formatHint: "csv" });
    const td = host.querySelector('td[data-rc="1:1"]') as HTMLElement;
    expect(td.classList.contains("sheetedit-calcerr")).toBe(true);
    const input = td.querySelector("input") as HTMLInputElement;
    expect(input.title.toLowerCase()).toContain("circular");
    ed.destroy();
    host.remove();
  });
});
