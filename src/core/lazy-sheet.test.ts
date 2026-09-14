import { afterEach, describe, expect, it, vi } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createSheetEditor, getCell, readWorkbook, writeWorkbook } from "../index";
import { setCellInput } from "./workbook";
import { needsCalcOnLoad, recalc } from "./recalc";
import { isSheetLoaded, loadSheet, onSheetLoaded } from "./lazy-sheet";

// Opening a workbook parses only what is used: a sheet nobody looks at costs nothing, and every
// caller still reads plain fields, which parse the sheet the first time they are touched.

const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const sheetXml = (inner: string): string => `<?xml version="1.0"?><worksheet ${NS}>${inner}</worksheet>`;

function makeXlsx(sheets: string[]): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "xl/workbook.xml": strToU8(
      `<workbook ${NS}><sheets>${sheets.map((_, i) => `<sheet name="S${i + 1}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join("")}</Relationships>`,
    ),
  };
  sheets.forEach((s, i) => (files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml(s))));
  return zipSync(files);
}

const FIRST = `<sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>S2!A1*3</f><v>30</v></c></row></sheetData>`;
const SECOND = `<sheetPr><tabColor rgb="FFFF0000"/></sheetPr><sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>`;

describe("parsing sheets on first use", () => {
  it("parses no sheet on open, yet knows each tab's colour", () => {
    const wb = readWorkbook(makeXlsx([FIRST, SECOND]), { lazySheets: true });
    expect(wb.sheets.map(isSheetLoaded)).toEqual([false, false]);
    expect(wb.sheets[1]!.tabColor).toBe("#ff0000");
    expect(wb.sheets.map(isSheetLoaded)).toEqual([false, false]);
  });

  it("parses one sheet when its cells are read, and only that one", () => {
    const wb = readWorkbook(makeXlsx([FIRST, SECOND]), { lazySheets: true });
    expect(getCell(wb.sheets[1]!, 1, 1)?.value).toBe("10");
    expect(wb.sheets[1]!.merges).toEqual([{ r1: 2, c1: 1, r2: 2, c2: 2 }]);
    expect(wb.sheets.map(isSheetLoaded)).toEqual([false, true]);
  });

  it("reads a sheet the same whether parsed up front or on first use", () => {
    const eager = readWorkbook(makeXlsx([FIRST, SECOND]));
    const lazy = readWorkbook(makeXlsx([FIRST, SECOND]), { lazySheets: true });
    for (let i = 0; i < 2; i++) {
      expect([...lazy.sheets[i]!.cells.values()].map((c) => [c.row, c.col, c.value, c.formula])).toEqual(
        [...eager.sheets[i]!.cells.values()].map((c) => [c.row, c.col, c.value, c.formula]),
      );
      expect(lazy.sheets[i]!.maxRow).toBe(eager.sheets[i]!.maxRow);
    }
  });

  // A write before any read must land on the parsed sheet, not be wiped by the parse after it.
  it("keeps a field written before the sheet was parsed", () => {
    const wb = readWorkbook(makeXlsx([FIRST, SECOND]), { lazySheets: true });
    wb.sheets[1]!.freeze = { rows: 1, cols: 0 };
    expect(wb.sheets[1]!.freeze).toEqual({ rows: 1, cols: 0 });
    expect(getCell(wb.sheets[1]!, 1, 1)?.value).toBe("10");
  });

  it("recomputes a formula that reads a sheet not parsed yet", () => {
    const wb = readWorkbook(makeXlsx([FIRST, SECOND]), { lazySheets: true });
    setCellInput(wb.sheets[0]!, 1, 2, "=S2!A1*4");
    recalc(wb);
    expect(getCell(wb.sheets[0]!, 1, 2)?.value).toBe("40");
  });

  it("saves an untouched workbook with every worksheet byte for byte", () => {
    const bytes = makeXlsx([FIRST, SECOND]);
    const wb = readWorkbook(bytes, { lazySheets: true });
    const before = unzipSync(bytes);
    const after = unzipSync(writeWorkbook(wb));
    for (const p of ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]) expect(strFromU8(after[p]!)).toBe(strFromU8(before[p]!));
  });

  it("tells whoever listens about each sheet as it gets parsed", () => {
    const wb = readWorkbook(makeXlsx([FIRST, SECOND]), { lazySheets: true });
    const seen: string[] = [];
    onSheetLoaded(wb, (s) => seen.push(s.name));
    loadSheet(wb.sheets[1]!);
    loadSheet(wb.sheets[1]!);
    void wb.sheets[0]!.cells;
    expect(seen).toEqual(["S2", "S1"]);
  });
});

describe("deciding on a load-time recalc without parsing", () => {
  const withRows = (rows: string): Uint8Array => makeXlsx([FIRST, `<sheetData>${rows}</sheetData>`]);

  it("says yes for a formula with no result in a sheet not parsed", () => {
    const wb = readWorkbook(withRows(`<row r="1"><c r="A1"><f>1+1</f></c></row>`), { lazySheets: true });
    expect(needsCalcOnLoad(wb)).toBe(true);
  });

  it("says yes for the empty placeholder openpyxl writes", () => {
    const wb = readWorkbook(withRows(`<row r="1"><c r="A1"><f>1+1</f><v /></c></row>`), { lazySheets: true });
    expect(needsCalcOnLoad(wb)).toBe(true);
  });

  it("says no for an empty string result, and parses nothing to find out", () => {
    const wb = readWorkbook(withRows(`<row r="1"><c r="A1" t="str"><f>IF(1,"","x")</f><v/></c></row>`), { lazySheets: true });
    expect(needsCalcOnLoad(wb)).toBe(false);
    expect(wb.sheets.map(isSheetLoaded)).toEqual([false, false]);
  });
});

// An openpyxl-style sheet that is not the one shown first: its results are filled in once it is
// parsed, whether the idle loader gets there first or the reader opens its tab first.
describe("results left out on a sheet parsed after opening", () => {
  const UNCOMPUTED = `<sheetData><row r="1"><c r="A1"><v>4</v></c><c r="B1"><f>A1*5</f></c></row></sheetData>`;
  const mount = () => {
    (globalThis as { ResizeObserver?: unknown }).ResizeObserver ??= class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
    const container = document.createElement("div");
    document.body.appendChild(container);
    const ed = createSheetEditor(container, makeXlsx([FIRST, UNCOMPUTED]));
    const openSecond = () => (container.querySelectorAll(".sheetedit-tab")[1] as HTMLElement).click();
    return { ed, openSecond, done: () => { ed.destroy(); container.remove(); } };
  };
  afterEach(() => vi.useRealTimers());

  it("computes them when the idle loader parses the sheet", () => {
    vi.useFakeTimers();
    const { ed, openSecond, done } = mount();
    vi.advanceTimersByTime(3000);
    openSecond();
    expect(ed.getCellValue("B1")).toBe("20");
    done();
  });

  it("computes them when the tab is opened before the idle loader gets there", () => {
    vi.useFakeTimers();
    const { ed, openSecond, done } = mount();
    openSecond();
    vi.advanceTimersByTime(3000);
    expect(ed.getCellValue("B1")).toBe("20");
    done();
  });
});
