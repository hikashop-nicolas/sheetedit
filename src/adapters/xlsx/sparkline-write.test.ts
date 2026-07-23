import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../core/workbook";
import { setXlsxSparkline, setXlsxSparklineGroup } from "./write";

function base(): ReturnType<typeof readWorkbook> {
  const xlsx = zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B1"><v>3</v></c><c r="C1"><v>5</v></c><c r="D1"><v>2</v></c></row></sheetData></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
  });
  return readWorkbook(xlsx);
}

describe("sparkline authoring", () => {
  it("writes a sparkline that round-trips into the model", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxSparkline(sheet, { r: 1, c: 7 }, { type: "column", color: "#00b050", dataRef: "Sheet1!B1:D1" });
    expect(sheet.sparklines?.[0]).toMatchObject({ type: "column", host: { r: 1, c: 7 }, dataRef: "Sheet1!B1:D1" });
    const re = readWorkbook(writeWorkbook(wb));
    const sp = re.sheets[0].sparklines?.[0];
    expect(sp).toBeTruthy();
    expect(sp?.type).toBe("column");
    expect(sp?.host).toEqual({ r: 1, c: 7 });
    expect(sp?.dataRef).toBe("Sheet1!B1:D1");
    expect(sp?.color.toLowerCase()).toContain("00b050");
  });

  it("round-trips a custom negative colour for win/loss", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxSparkline(sheet, { r: 1, c: 7 }, { type: "stacked", color: "#376092", negColor: "#ff8800", dataRef: "Sheet1!B1:D1" });
    const re = readWorkbook(writeWorkbook(wb));
    const sp = re.sheets[0].sparklines?.[0];
    expect(sp?.type).toBe("stacked");
    expect(sp?.negColor?.toLowerCase()).toContain("ff8800");
  });

  it("removes a sparkline and cleans up the extLst", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxSparkline(sheet, { r: 1, c: 7 }, { type: "line", color: "#376092", dataRef: "Sheet1!B1:D1" });
    setXlsxSparkline(sheet, { r: 1, c: 7 }, null);
    expect(sheet.sparklines).toBeUndefined();
    const re = readWorkbook(writeWorkbook(wb));
    expect(re.sheets[0].sparklines).toBeUndefined();
  });

  it("authors one group spanning several location cells", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxSparklineGroup(sheet, { type: "line", color: "#376092" }, [
      { host: { r: 1, c: 7 }, dataRef: "Sheet1!B1:D1" },
      { host: { r: 2, c: 7 }, dataRef: "Sheet1!B2:D2" },
      { host: { r: 3, c: 7 }, dataRef: "Sheet1!B3:D3" },
    ]);
    // Model has one entry per host.
    expect(sheet.sparklines?.length).toBe(3);
    const re = readWorkbook(writeWorkbook(wb));
    const sps = re.sheets[0].sparklines ?? [];
    expect(sps.length).toBe(3);
    expect(sps.map((s) => `${s.host.r}:${s.host.c}`).sort()).toEqual(["1:7", "2:7", "3:7"]);
    // The three sparklines live in a single <sparklineGroup> (one group, three sparklines).
    const xml = strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]);
    expect((xml.match(/<(?:\w+:)?sparklineGroup[ >]/g) ?? []).length).toBe(1);
    expect((xml.match(/<(?:\w+:)?sparkline[ >]/g) ?? []).length).toBe(3);
  });

  it("replacing the host's sparkline does not stack duplicates", () => {
    const wb = base(); const sheet = wb.sheets[0];
    setXlsxSparkline(sheet, { r: 1, c: 7 }, { type: "line", color: "#376092", dataRef: "Sheet1!B1:D1" });
    setXlsxSparkline(sheet, { r: 1, c: 7 }, { type: "stacked", color: "#c00000", dataRef: "Sheet1!B1:D1" });
    const re = readWorkbook(writeWorkbook(wb));
    const sps = re.sheets[0].sparklines ?? [];
    expect(sps.length).toBe(1);
    expect(sps[0].type).toBe("stacked");
  });
});
