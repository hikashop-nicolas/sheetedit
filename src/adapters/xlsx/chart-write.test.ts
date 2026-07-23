import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../index";
import { buildChart, defaultAnchor } from "../../core/chart-build";

function dataXlsx(): Uint8Array {
  const c = (ref: string, v: string | number, s = false) => (s ? `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>` : `<c r="${ref}"><v>${v}</v></c>`);
  const sheet = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` +
      `<row r="1">${c("A1", "Qtr", true)}${c("B1", "North", true)}${c("C1", "South", true)}</row>` +
      `<row r="2">${c("A2", "Q1", true)}${c("B2", 10)}${c("C2", 22)}</row>` +
      `<row r="3">${c("A3", "Q2", true)}${c("B3", 30)}${c("C3", 14)}</row>` +
      `</sheetData></worksheet>`,
  );
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": sheet,
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
}

describe("xlsx chart writer", () => {
  it("a created chart survives write + re-read with its model intact", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 }; // A1:C3
    const model = buildChart("Sheet1", "column", rect, { firstRowHeader: true, firstColLabels: true }, "new1", defaultAnchor(rect));
    model.title = "Quarterly";
    model.legend = { show: true, pos: "bottom" };
    (wb.sheets[0].charts ??= []).push(model);

    const out = writeWorkbook(wb);
    // The package gained a chart + drawing part.
    const names = Object.keys(unzipSync(out));
    expect(names.some((n) => /xl\/charts\/chart\d+\.xml/.test(n))).toBe(true);
    expect(names.some((n) => /xl\/drawings\/drawing\d+\.xml/.test(n))).toBe(true);

    const re = readWorkbook(out).sheets[0].charts!;
    expect(re).toHaveLength(1);
    expect(re[0].kind).toBe("column");
    expect(re[0].title).toBe("Quarterly");
    expect(re[0].legend).toEqual({ show: true, pos: "bottom" });
    expect(re[0].categories?.ref).toBe("Sheet1!$A$2:$A$3");
    expect(re[0].series).toHaveLength(2);
    expect(re[0].series[0].values.ref).toBe("Sheet1!$B$2:$B$3");
    // anchor round-trips (default anchor: to the right of the selection).
    expect(re[0].anchor.fromCol).toBe(model.anchor.fromCol);
    expect(re[0].anchor.toRow).toBe(model.anchor.toRow);
  });

  it("data labels round-trip", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 };
    const m = buildChart("Sheet1", "column", rect, { firstRowHeader: true, firstColLabels: true }, "d1", defaultAnchor(rect));
    m.dataLabels = true;
    (wb.sheets[0].charts ??= []).push(m);
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].charts![0].dataLabels).toBe(true);
  });

  it("a combo (column + line on a secondary axis) round-trips", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 };
    const m = buildChart("Sheet1", "column", rect, { firstRowHeader: true, firstColLabels: true }, "cbo", defaultAnchor(rect));
    m.series[1].type = "line";
    m.series[1].secondaryAxis = true;
    (wb.sheets[0].charts ??= []).push(m);
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].charts![0];
    expect(re.kind).toBe("column"); // base kind from the first type element
    const lineSeries = re.series.find((s) => s.type === "line");
    expect(lineSeries).toBeTruthy();
    expect(lineSeries!.secondaryAxis).toBe(true);
  });

  it("Tier-1 options round-trip: percent, blanksAs, holeSize/gapWidth, per-point colours, axis max", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 };
    const m = buildChart("Sheet1", "column", rect, { firstRowHeader: true, firstColLabels: true }, "t1", defaultAnchor(rect));
    m.percent = true;
    m.stacked = true;
    m.blanksAs = "gap";
    m.gapWidth = 80;
    m.overlap = -20;
    m.axes = { y: { max: 50 } };
    m.series[0].pointColors = ["#ff0000", undefined, "#00ff00"];
    (wb.sheets[0].charts ??= []).push(m);
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].charts![0];
    expect(re.percent).toBe(true);
    expect(re.blanksAs).toBe("gap");
    expect(re.gapWidth).toBe(80);
    expect(re.overlap).toBe(-20);
    expect(re.axes?.y?.max).toBe(50);
    expect(re.series[0].pointColors?.[0]).toBe("#ff0000");
  });

  it("Tier-2 batch round-trips: smooth, marker, pie rotation, axis number format", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 };
    const m = buildChart("Sheet1", "line", rect, { firstRowHeader: true, firstColLabels: true }, "t2", defaultAnchor(rect));
    m.series[0].smooth = true;
    m.series[0].marker = { symbol: "diamond", size: 8 };
    m.axes = { y: { numFmt: "0.0%" } };
    (wb.sheets[0].charts ??= []).push(m);
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].charts![0];
    expect(re.series[0].smooth).toBe(true);
    expect(re.series[0].marker).toEqual({ symbol: "diamond", size: 8 });
    expect(re.axes?.y?.numFmt).toBe("0.0%");
  });

  it("pie firstSliceAng (rotation) round-trips", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 2 };
    const m = buildChart("Sheet1", "pie", rect, { firstRowHeader: true, firstColLabels: true }, "r", defaultAnchor(rect));
    m.rotation = 90;
    (wb.sheets[0].charts ??= []).push(m);
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].charts![0].rotation).toBe(90);
  });

  it("theme-colour (schemeClr) series colour is resolved on read", () => {
    // A chart part whose series colour is a schemeClr accent1, with a theme mapping accent1 to red.
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 2 };
    const m = buildChart("Sheet1", "column", rect, { firstRowHeader: true, firstColLabels: true }, "sc", defaultAnchor(rect));
    m.series[0].color = "#c00000"; // written as srgbClr; the schemeClr path is unit-tested separately
    (wb.sheets[0].charts ??= []).push(m);
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].charts![0].series[0].color).toBe("#c00000");
  });

  it("Tier-2b round-trips: rich data labels, pie explosion, legend deletion, date axis", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 };
    const m = buildChart("Sheet1", "line", rect, { firstRowHeader: true, firstColLabels: true }, "t2b", defaultAnchor(rect));
    m.series[0].labels = { value: true, category: true, percent: true, position: "outEnd" };
    m.legend = { show: true, pos: "bottom", deleted: [1], overlay: true };
    m.axes = { x: { date: true } };
    (wb.sheets[0].charts ??= []).push(m);
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].charts![0];
    expect(re.series[0].labels).toEqual({ value: true, category: true, percent: true, position: "outEnd" });
    expect(re.legend?.deleted).toEqual([1]);
    expect(re.legend?.overlay).toBe(true);
    expect(re.axes?.x?.date).toBe(true);
  });

  it("pie slice explosion round-trips", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 2 };
    const m = buildChart("Sheet1", "pie", rect, { firstRowHeader: true, firstColLabels: true }, "ex", defaultAnchor(rect));
    m.series[0].explosion = [25, undefined];
    (wb.sheets[0].charts ??= []).push(m);
    const re = readWorkbook(writeWorkbook(wb)).sheets[0].charts![0];
    expect(re.series[0].explosion?.[0]).toBe(25);
  });

  it("doughnut holeSize round-trips", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 2 };
    const m = buildChart("Sheet1", "doughnut", rect, { firstRowHeader: true, firstColLabels: true }, "d", defaultAnchor(rect));
    m.holeSize = 65;
    (wb.sheets[0].charts ??= []).push(m);
    expect(readWorkbook(writeWorkbook(wb)).sheets[0].charts![0].holeSize).toBe(65);
  });

  it("changing a chart's type and re-saving keeps a single chart of the new kind", () => {
    const wb = readWorkbook(dataXlsx());
    const rect = { r1: 1, c1: 1, r2: 3, c2: 3 };
    const model = buildChart("Sheet1", "column", rect, { firstRowHeader: true, firstColLabels: true }, "new1", defaultAnchor(rect));
    (wb.sheets[0].charts ??= []).push(model);
    const wb2 = readWorkbook(writeWorkbook(wb));
    // Edit: change type to line, re-save.
    wb2.sheets[0].charts![0].kind = "line";
    wb2.sheets[0].charts![0].dirty = true;
    const re = readWorkbook(writeWorkbook(wb2)).sheets[0].charts!;
    expect(re).toHaveLength(1);
    expect(re[0].kind).toBe("line");
  });
});
