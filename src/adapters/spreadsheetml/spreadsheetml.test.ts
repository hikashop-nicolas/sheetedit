import { describe, expect, it } from "vitest";
import { strToU8 } from "fflate";
import { getCell, type Workbook } from "../../core/model";
import { readWorkbook } from "../../core/workbook";
import { readSpreadsheetMl } from "../tabular";
import { r1c1Range, r1c1ToA1 } from "./r1c1";

const val = (wb: Workbook, si: number, r: number, c: number): string | undefined => getCell(wb.sheets[si]!, r, c)?.value;
const cell = (wb: Workbook, si: number, r: number, c: number) => getCell(wb.sheets[si]!, r, c);

describe("R1C1 to A1", () => {
  const cases: [string, number, number, string][] = [
    ["R1C1", 0, 0, "$A$1"],
    ["RC", 4, 2, "C5"],
    ["R[-1]C[2]", 4, 2, "E4"],
    ["SUM(R[-3]C:R[-1]C)", 3, 0, "SUM(A1:A3)"],
    ["'My Sheet'!R2C3+Sheet2!RC[-1]", 0, 1, "'My Sheet'!$C$2+Sheet2!A1"],
    ["SUM(C2:C3)", 0, 0, "SUM($B:$C)"],
    ["SUM(R2:R5)", 0, 0, "SUM($2:$5)"],
    ["ROUND(RC1,2)&\"R1C1\"", 9, 5, "ROUND($A10,2)&\"R1C1\""],
    ["Rate*COUNT(R1C1:R[1]C)", 0, 0, "Rate*COUNT($A$1:A2)"],
    ["MOD(SUBTOTAL(103,RC1:R2C1),2)=1", 12, 0, "MOD(SUBTOTAL(103,$A13:$A$2),2)=1"],
  ];
  for (const [f, r, c, want] of cases) it(`${f} at R${r + 1}C${c + 1}`, () => expect(r1c1ToA1(f, r, c)).toBe(want));
  it("reads ranges", () => {
    expect(r1c1Range("R2C1:R5C3")).toEqual({ r1: 1, c1: 0, r2: 4, c2: 2 });
    expect(r1c1Range("C2")).toEqual({ r1: 0, c1: 1, r2: 1048575, c2: 1 });
  });
});

const BOOK = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel"><ActiveSheet>1</ActiveSheet></ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Bottom"/><Font ss:FontName="Arial" ss:Size="10"/></Style>
  <Style ss:ID="hdr"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#003366" ss:Pattern="Solid"/>
   <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/></Borders><Alignment ss:Horizontal="Center" ss:WrapText="1"/></Style>
  <Style ss:ID="pct" ss:Parent="hdr"><NumberFormat ss:Format="0.0%"/><Font ss:Italic="1"/></Style>
  <Style ss:ID="date"><NumberFormat ss:Format="Short Date"/></Style>
 </Styles>
 <Names><NamedRange ss:Name="Prices" ss:RefersTo="=Data!R2C2:R3C2"/></Names>
 <Worksheet ss:Name="Data">
  <Names><NamedRange ss:Name="Print_Area" ss:RefersTo="=Data!R1C1:R4C3"/></Names>
  <Table ss:DefaultColumnWidth="48">
   <Column ss:Width="120"/><Column ss:Index="3" ss:Width="30" ss:Span="1" ss:Hidden="1"/>
   <Row ss:Height="30">
    <Cell ss:StyleID="hdr"><Data ss:Type="String">Item</Data></Cell>
    <Cell ss:StyleID="hdr" ss:MergeAcross="1"><Data ss:Type="String">Price</Data></Cell>
   </Row>
   <Row>
    <Cell ss:HRef="https://example.com/a" ss:HRefScreenTip="Open"><Data ss:Type="String">Apple</Data><Comment ss:Author="Ann"><ss:Data><B>Fresh</B> today</ss:Data></Comment></Cell>
    <Cell><Data ss:Type="Number">1.5</Data></Cell>
    <Cell ss:StyleID="date"><Data ss:Type="DateTime">2024-03-15T00:00:00.000</Data></Cell>
    <Cell ss:StyleID="pct"><Data ss:Type="Number">0.25</Data></Cell>
   </Row>
   <Row ss:Index="4" ss:Hidden="1">
    <Cell><Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40"><B>Bold</B><Font html:Color="#FF0000"> red</Font></Data></Cell>
    <Cell ss:Index="2" ss:Formula="=SUM(R[-2]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>
    <Cell><Data ss:Type="Boolean">1</Data></Cell>
    <Cell><Data ss:Type="Error">#N/A</Data></Cell>
   </Row>
   <Row><Cell ss:Index="2" ss:Formula="=Prices*2" ss:ArrayRange="RC:R[1]C"><Data ss:Type="Number">3</Data></Cell></Row>
   <Row><Cell ss:Index="2" ss:Formula="=Prices*2"><Data ss:Type="Number">0</Data></Cell><Cell ss:MergeDown="1"><Data ss:Type="String">Down</Data></Cell></Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup><Layout x:Orientation="Landscape"/></PageSetup>
   <FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane>
   <TabColorIndex>10</TabColorIndex>
  </WorksheetOptions>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel"><Range>R2C1:R10C1</Range><Type>List</Type><Value>&quot;Apple,Pear&quot;</Value></DataValidation>
  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel"><Range>R2C2:R10C2</Range><Type>Decimal</Type><Min>0</Min><Max>R1C5</Max></DataValidation>
  <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel"><Range>R2C2:R10C2</Range>
   <Condition><Qualifier>Greater</Qualifier><Value1>1</Value1><Format Style='color:#9C0006;background:#FFC7CE'/></Condition>
  </ConditionalFormatting>
  <AutoFilter x:Range="R1C1:R6C3" xmlns="urn:schemas-microsoft-com:office:excel"/>
 </Worksheet>
 <Worksheet ss:Name="Hidden">
  <Table><Row><Cell><Data ss:Type="Number">7</Data></Cell></Row></Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><Visible>SheetHidden</Visible></WorksheetOptions>
 </Worksheet>
</Workbook>`;

describe("Excel 2003 XML Spreadsheet", () => {
  for (const useDom of [true, false]) {
    it(`reads cells, styles, formulas and sheet settings${useDom ? "" : " without a DOM"}`, () => {
      const wb = readSpreadsheetMl(BOOK, { useDom });
      expect(wb.importedFrom).toBe("spreadsheetml");
      expect(wb.sheets.map((s) => s.name)).toEqual(["Data", "Hidden"]);
      expect(wb.activeSheet).toBe(1);
      const s = wb.sheets[0]!;
      expect(wb.defaultFontName).toBe("Arial");
      const hdr = cell(wb, 0, 1, 1)!.cellStyle!;
      expect(hdr.bold).toBe(true);
      expect(hdr.color?.toLowerCase()).toBe("#ffffff");
      expect(hdr.bg?.toLowerCase()).toBe("#003366");
      expect(hdr.borders?.bottom).toBeTruthy();
      expect(hdr.align).toBe("center");
      expect(hdr.wrap).toBe(true);
      expect(s.merges).toContainEqual({ r1: 1, c1: 2, r2: 1, c2: 3 });
      expect(s.merges).toContainEqual({ r1: 6, c1: 3, r2: 7, c2: 3 });
      expect(val(wb, 0, 2, 2)).toBe("1.5");
      expect(Number(val(wb, 0, 2, 3))).toBe(45366);
      // The child style inherits the header look and adds its own.
      const pct = cell(wb, 0, 2, 4)!;
      expect(pct.cellStyle?.bold && pct.cellStyle?.italic).toBe(true);
      expect(pct.numFmt).toBe("0.0%");
      expect(cell(wb, 0, 2, 1)?.link?.href).toBe("https://example.com/a");
      expect(cell(wb, 0, 2, 1)?.comments?.[0]?.text).toBe("Fresh today");
      expect(cell(wb, 0, 4, 1)?.richRuns?.map((r) => r.text)).toEqual(["Bold", " red"]);
      // Formulas are converted to A1 and recomputed on open.
      expect(cell(wb, 0, 4, 2)?.formula).toBe("SUM(B2:B3)");
      expect(val(wb, 0, 4, 2)).toBe("1.5");
      expect(cell(wb, 0, 4, 3)?.kind).toBe("b");
      expect(cell(wb, 0, 4, 4)?.kind).toBe("e");
      expect(cell(wb, 0, 5, 2)?.formula).toBe("Prices*2");
      expect(cell(wb, 0, 6, 2)?.formula).toBeUndefined();
      expect(s.hiddenRows?.has(4)).toBe(true);
      expect(s.hiddenCols?.has(3) && s.hiddenCols?.has(4)).toBe(true);
      expect(s.colWidths?.get(1)).toBe(160);
      expect(s.freeze).toEqual({ rows: 1, cols: 0 });
      expect(s.tabColor?.toLowerCase()).toBe("#ff0000");
      expect(s.autoFilter).toEqual({ r1: 1, c1: 1, r2: 6, c2: 3 });
      expect(s.validations?.[0]?.values).toEqual(["Apple", "Pear"]);
      expect(s.condFormats?.length).toBe(1);
      expect(wb.definedNames?.get("Prices")).toBe("Data!$B$2:$B$3");
      expect(wb.sheets[1]!.visibility).toBe("hidden");
    });
  }

  it("reads prefixed namespaces and tolerates broken XML without a DOM parser's help", () => {
    const prefixed = `<?xml version="1.0"?><s:Workbook xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"><s:Worksheet s:Name="One"><s:Table><s:Row><s:Cell><s:Data s:Type="String">A &amp; B</s:Data></s:Cell><s:Cell><s:Data s:Type="Number">2</s:Data></s:Cell></s:Row></s:Table></s:Worksheet></s:Workbook>`;
    const wb = readSpreadsheetMl(prefixed);
    expect(wb.sheets[0]!.name).toBe("One");
    expect(val(wb, 0, 1, 1)).toBe("A & B");
    expect(val(wb, 0, 1, 2)).toBe("2");
    // An unescaped ampersand breaks strict XML; the tolerant tokenizer still reads the cells.
    const broken = prefixed.replace("A &amp; B", "Fish & Chips");
    expect(val(readSpreadsheetMl(broken), 0, 1, 1)).toBe("Fish & Chips");
  });

  it("opens through readWorkbook, whatever the name", () => {
    const wb = readWorkbook(strToU8(BOOK));
    expect(wb.importedFrom).toBe("spreadsheetml");
    expect(wb.sheets).toHaveLength(2);
  });
});
