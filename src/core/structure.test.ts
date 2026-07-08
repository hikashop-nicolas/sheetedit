import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { applyLineOp, mapPoint, mapSpan, rewriteFormula, type LineOp } from "./structure";
import { createSheetEditor } from "./editor";
import { getCell } from "./model";
import { readWorkbook, writeWorkbook } from "./workbook";

const insRow = (at: number, count = 1): LineOp => ({ axis: "row", kind: "insert", at, count });
const delRow = (at: number, count = 1): LineOp => ({ axis: "row", kind: "delete", at, count });
const insCol = (at: number, count = 1): LineOp => ({ axis: "col", kind: "insert", at, count });
const delCol = (at: number, count = 1): LineOp => ({ axis: "col", kind: "delete", at, count });
const rw = (f: string, op: LineOp, formulaSheet = "Sheet1", targetSheet = "Sheet1") =>
  rewriteFormula(f, formulaSheet, targetSheet, op);

describe("mapPoint / mapSpan", () => {
  it("shifts points across inserts and deletes", () => {
    expect(mapPoint(5, insRow(3))).toBe(6);
    expect(mapPoint(2, insRow(3))).toBe(2);
    expect(mapPoint(3, insRow(3))).toBe(4);
    expect(mapPoint(5, delRow(3))).toBe(4);
    expect(mapPoint(2, delRow(3))).toBe(2);
    expect(mapPoint(3, delRow(3))).toBeNull();
    expect(mapPoint(4, delRow(3, 2))).toBeNull();
    expect(mapPoint(5, delRow(3, 2))).toBe(3);
  });
  it("shrinks, shifts and kills spans on delete", () => {
    expect(mapSpan(1, 10, delRow(5))).toEqual({ a: 1, b: 9 });
    expect(mapSpan(5, 5, delRow(5))).toBeNull();
    expect(mapSpan(4, 6, delRow(5, 2))).toEqual({ a: 4, b: 4 });
    expect(mapSpan(6, 9, delRow(2, 3))).toEqual({ a: 3, b: 6 });
  });
  it("extends spans crossed by an insert", () => {
    expect(mapSpan(2, 5, insRow(3))).toEqual({ a: 2, b: 6 });
    expect(mapSpan(2, 5, insRow(2))).toEqual({ a: 3, b: 6 });
    expect(mapSpan(2, 5, insRow(6))).toEqual({ a: 2, b: 5 });
  });
});

describe("rewriteFormula: rows", () => {
  it("shifts refs at or below an inserted row", () => {
    expect(rw("A1+A5", insRow(3))).toBe("A1+A6");
    expect(rw("SUM(A2:A5)", insRow(3))).toBe("SUM(A2:A6)");
  });
  it("shifts absolute refs too (structural ops ignore anchors)", () => {
    expect(rw("$A$5+B$4", insRow(3))).toBe("$A$6+B$5");
  });
  it("turns refs to a deleted row into #REF! and shrinks ranges", () => {
    expect(rw("A3*2", delRow(3))).toBe("#REF!*2");
    expect(rw("SUM(A1:A10)", delRow(3))).toBe("SUM(A1:A9)");
    expect(rw("SUM(A3:B4)", delRow(3, 2))).toBe("SUM(#REF!)");
  });
  it("handles whole-row ranges", () => {
    expect(rw("SUM(2:5)", insRow(3))).toBe("SUM(2:6)");
    expect(rw("SUM(2:5)", delRow(2, 4))).toBe("SUM(#REF!)");
  });
  it("leaves string literals and function names alone", () => {
    expect(rw('CONCAT("A5 is ", A5)', insRow(3))).toBe('CONCAT("A5 is ", A6)');
    expect(rw("LOG10(A5)", insRow(3))).toBe("LOG10(A6)");
  });
  it("only rewrites refs that resolve to the target sheet", () => {
    expect(rw("Sheet2!A5+A5", insRow(3), "Sheet1", "Sheet1")).toBe("Sheet2!A5+A6");
    expect(rw("Sheet2!A5+A5", insRow(3), "Sheet1", "Sheet2")).toBe("Sheet2!A6+A5");
    expect(rw("'My Sheet'!A5", insRow(3), "Sheet1", "My Sheet")).toBe("'My Sheet'!A6");
    expect(rw("'My Sheet'!A5:B9", delRow(9), "Sheet1", "My Sheet")).toBe("'My Sheet'!A5:B8");
  });
  it("keeps the sheet prefix on a #REF!", () => {
    expect(rw("Sheet2!A3", delRow(3), "Sheet1", "Sheet2")).toBe("Sheet2!#REF!");
  });
});

describe("rewriteFormula: columns", () => {
  it("shifts refs and ranges across column ops", () => {
    expect(rw("B1+D1", insCol(3))).toBe("B1+E1");
    expect(rw("SUM(B1:D9)", delCol(3))).toBe("SUM(B1:C9)");
    expect(rw("C7", delCol(3))).toBe("#REF!");
  });
  it("handles whole-column ranges", () => {
    expect(rw("SUM(B:D)", insCol(3))).toBe("SUM(B:E)");
    expect(rw("SUM(B:D)", delCol(2, 3))).toBe("SUM(#REF!)");
    expect(rw("SUM(Sheet2!B:D)", delCol(3), "Sheet1", "Sheet2")).toBe("SUM(Sheet2!B:C)");
  });
});

// ---------------------------------------------------------------------------
// End to end on real workbooks
// ---------------------------------------------------------------------------

const SHEET1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <dimension ref="A1:C3"/>
 <sheetData>
  <row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>one</t></is></c></row>
  <row r="2"><c r="A2"><v>2</v></c><c r="B2" t="inlineStr"><is><t>two</t></is></c></row>
  <row r="3"><c r="A3"><f>SUM(A1:A2)</f><v>3</v></c><c r="C3" ht="1"><v>9</v></c></row>
 </sheetData>
 <mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells>
</worksheet>`;

function makeXlsx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "xl/workbook.xml": strToU8(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(SHEET1),
  });
}

describe("applyLineOp on xlsx", () => {
  it("insert row: values shift, formula grows, merge moves, XML round-trips", () => {
    const wb = readWorkbook(makeXlsx());
    applyLineOp(wb, 0, insRow(2));
    const s = wb.sheets[0]!;
    expect(getCell(s, 1, 1)?.value).toBe("1");
    expect(getCell(s, 2, 1)).toBeUndefined();
    expect(getCell(s, 3, 1)?.value).toBe("2");
    expect(getCell(s, 4, 1)?.formula).toBe("SUM(A1:A3)");
    expect(s.merges).toEqual([{ r1: 3, c1: 2, r2: 3, c2: 3 }]);

    const out = unzipSync(writeWorkbook(wb));
    const xml = strFromU8(out["xl/worksheets/sheet1.xml"]!);
    expect(xml).toContain('r="A3"'); // the shifted value 2
    expect(xml).toContain("SUM(A1:A3)");
    expect(xml).toContain('ref="B3:C3"');
    expect(xml).not.toContain('ref="B2:C2"');

    // Reparse: the written workbook must read back consistently.
    const wb2 = readWorkbook(writeWorkbook(wb));
    expect(getCell(wb2.sheets[0]!, 4, 1)?.formula).toBe("SUM(A1:A3)");
    expect(getCell(wb2.sheets[0]!, 3, 1)?.value).toBe("2");
  });

  it("delete row: cells drop, refs shrink, dependent recalc runs at save", () => {
    const wb = readWorkbook(makeXlsx());
    applyLineOp(wb, 0, delRow(2));
    const s = wb.sheets[0]!;
    expect(getCell(s, 2, 1)?.formula).toBe("SUM(A1:A1)");
    const wb2 = readWorkbook(writeWorkbook(wb));
    const f = getCell(wb2.sheets[0]!, 2, 1);
    expect(f?.formula).toBe("SUM(A1:A1)");
    expect(f?.value).toBe("1"); // recalculated: only A1 remains
  });

  it("insert column: refs, merges and cell refs shift", () => {
    const wb = readWorkbook(makeXlsx());
    applyLineOp(wb, 0, insCol(2));
    const s = wb.sheets[0]!;
    expect(getCell(s, 1, 3)?.value).toBe("one");
    expect(s.merges).toEqual([{ r1: 2, c1: 3, r2: 2, c2: 4 }]);
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]!);
    expect(xml).toContain('r="C1"');
    expect(xml).toContain('ref="C2:D2"');
  });

  it("delete column drops its cells (including a formula living in it)", () => {
    const wb = readWorkbook(makeXlsx());
    applyLineOp(wb, 0, delCol(1)); // the SUM formula sits in column A
    const s = wb.sheets[0]!;
    expect(getCell(s, 1, 1)?.value).toBe("one"); // B1 became A1
    expect([...s.cells.values()].some((c) => c.formula != null)).toBe(false);
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]!);
    expect(xml).not.toContain("SUM(");
  });

  it("delete columns a surviving formula ranges over yields #REF!", () => {
    const wb = readWorkbook(makeXlsx());
    // Move the formula out of harm's way: reference A1:A2 from D1.
    const s = wb.sheets[0]!;
    const f = getCell(s, 3, 1)!;
    s.cells.delete("3:1");
    f.row = 1;
    f.col = 4;
    f.edited = true;
    s.cells.set("1:4", f);
    applyLineOp(wb, 0, delCol(1)); // shrink: SUM(A1:A2) stays on the deleted-into range
    expect(getCell(s, 1, 3)?.formula).toBe("SUM(#REF!)");
  });
});

describe("editor: header context menu and structural undo", () => {
  // jsdom lacks ResizeObserver (the toolbar overflow logic uses it).
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
  it("deletes a row from the row-header menu, undo restores content and formula", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, makeXlsx());
    const grid = () => (r: number, c: number) =>
      (host.querySelector(`td[data-rc="${r}:${c}"] input`) as HTMLInputElement | null)?.value;
    expect(grid()(2, 1)).toBe("2");
    expect(grid()(3, 1)).toBe("3"); // SUM(A1:A2)

    const rn = [...host.querySelectorAll("th.rownum")].find((th) => th.textContent === "2")!;
    rn.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    const items = [...document.querySelectorAll(".sheetedit-pop-item")] as HTMLButtonElement[];
    expect(items.length).toBe(3);
    items[2]!.click(); // delete row 2

    expect(grid()(2, 1)).toBe("1"); // the formula moved up and now sums A1 only
    const undoBtn = [...host.querySelectorAll("button")].find((b) => (b as HTMLButtonElement).title.includes("Ctrl+Z")) as HTMLButtonElement;
    undoBtn.click();
    expect(grid()(2, 1)).toBe("2");
    expect(grid()(3, 1)).toBe("3");

    // The saved file after undo still has the original formula and both values.
    const xml = strFromU8(unzipSync(await ed.getBytes())["xl/worksheets/sheet1.xml"]!);
    expect(xml).toContain("SUM(A1:A2)");
    ed.destroy();
    host.remove();
  });
});

const ODS_CONTENT = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:spreadsheet>
  <table:table table:name="Sheet1">
   <table:table-column table:number-columns-repeated="3"/>
   <table:table-row>
    <table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>
    <table:table-cell office:value-type="string" office:string-value="one"><text:p>one</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell table:formula="of:=SUM([.A1:.A2])" office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>`;

function makeOds(): Uint8Array {
  const packed: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
    "content.xml": strToU8(ODS_CONTENT),
  };
  return zipSync(packed as Record<string, Uint8Array>);
}

describe("applyLineOp on ods", () => {
  it("insert row shifts values, grows the formula, and round-trips", () => {
    const wb = readWorkbook(makeOds());
    applyLineOp(wb, 0, insRow(2));
    const s = wb.sheets[0]!;
    expect(getCell(s, 3, 1)?.value).toBe("2");
    expect(getCell(s, 4, 1)?.formula).toBe("SUM(A1:A3)");
    const wb2 = readWorkbook(writeWorkbook(wb));
    expect(getCell(wb2.sheets[0]!, 4, 1)?.formula).toBe("SUM(A1:A3)");
    expect(getCell(wb2.sheets[0]!, 2, 1)).toBeUndefined();
  });

  it("delete column updates the declared column count", () => {
    const wb = readWorkbook(makeOds());
    applyLineOp(wb, 0, delCol(2));
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(xml).toContain('table:number-columns-repeated="2"');
  });

  it("insert column grows the declared column count", () => {
    const wb = readWorkbook(makeOds());
    applyLineOp(wb, 0, insCol(2));
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["content.xml"]!);
    expect(xml).toContain('table:number-columns-repeated="4"');
  });
});
