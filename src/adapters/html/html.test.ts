import { describe, expect, it } from "vitest";
import { strToU8 } from "fflate";
import { getCell, type Workbook } from "../../core/model";
import { readWorkbook } from "../../core/workbook";
import { readHtmlWorkbook } from "../tabular";
import { parseMarkup } from "../tabular/tree";

const val = (wb: Workbook, si: number, r: number, c: number): string | undefined => getCell(wb.sheets[si]!, r, c)?.value;
const cell = (wb: Workbook, si: number, r: number, c: number) => getCell(wb.sheets[si]!, r, c);

// An Excel "Save as Web Page" file in miniature: class styles, x: attributes, its XML island.
const EXCEL_PAGE = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta http-equiv=Content-Type content="text/html; charset=utf-8"><meta name=ProgId content=Excel.Sheet>
<style>
<!--table {mso-displayed-decimal-separator:"\\,"; mso-displayed-thousand-separator:"\\.";}
.style0 {mso-number-format:General; font-size:10.0pt; font-family:Arial;}
td {mso-style-parent:style0; font-size:10.0pt; font-family:Arial; white-space:nowrap;}
.xl65 {font-weight:700; background:#FFFF00; mso-pattern:auto none; border:.5pt solid windowtext;}
.xl66 {mso-number-format:"0\\.00"; text-align:right;}
.xl67 {mso-number-format:"Short Date";}
.xl68 {mso-number-format:"\\@";}
.xl69 {color:red; font-style:italic; white-space:normal;}
-->
</style>
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Ventes 2024</x:Name>
  <x:WorksheetOptions><x:FreezePanes/><x:FrozenNoSplit/><x:SplitHorizontal>1</x:SplitHorizontal><x:TopRowBottomPane>1</x:TopRowBottomPane><x:TabColorIndex>10</x:TabColorIndex></x:WorksheetOptions>
 </x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook>
 <x:ExcelName><x:Name>Total</x:Name><x:Formula>='Ventes 2024'!$B$2:$B$3</x:Formula></x:ExcelName>
</xml><![endif]-->
</head><body>
<table x:str border=0 cellpadding=0 cellspacing=0 width=300>
 <col width=120 style='width:90pt'>
 <col width=80 span=2 style='width:60pt'>
 <tr height=17 style='height:12.75pt'>
  <td class=xl65>Produit</td><td class=xl65>Prix</td><td class=xl65>Date</td>
 </tr>
 <tr height=17>
  <td x:str="'007">007</td><td class=xl66 x:num="1234.5">1.234,50</td><td class=xl67 x:num="45366">15/03/2024</td>
 </tr>
 <tr height=17>
  <td class=xl68>00123</td><td class=xl66 x:num>2,25</td><td x:fmla="=B2+B3" x:num="1236.75">1.236,75</td>
 </tr>
 <tr height=17>
  <td colspan=2 style='mso-ignore:colspan'>Long text that overflows</td><td x:bool="TRUE">VRAI</td>
 </tr>
 <tr height=17>
  <td colspan=2 class=xl69>Merged</td><td x:err="#DIV/0!">#DIV/0!</td>
 </tr>
 <tr height=17>
  <td>edu.codemao.c<span style='display:none'>n</span></td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'><td width=120></td><td width=80></td><td width=80></td></tr>
 <![endif]>
</table>
</body></html>`;

describe("Excel web page (.xls saved as HTML)", () => {
  for (const useDom of [true, false]) {
    it(`reads values, types, styles and the XML island${useDom ? "" : " without a DOM"}`, () => {
      const wb = readHtmlWorkbook(strToU8(EXCEL_PAGE), { useDom });
      expect(wb.importedFrom).toBe("html");
      expect(wb.sheets).toHaveLength(1);
      const s = wb.sheets[0]!;
      expect(s.name).toBe("Ventes 2024");
      expect(s.freeze).toEqual({ rows: 1, cols: 0 });
      expect(s.tabColor?.toLowerCase()).toBe("#ff0000");
      expect(val(wb, 0, 1, 1)).toBe("Produit");
      expect(cell(wb, 0, 1, 1)?.cellStyle?.bold).toBe(true);
      expect(cell(wb, 0, 1, 1)?.cellStyle?.bg?.toLowerCase()).toBe("#ffff00");
      expect(cell(wb, 0, 1, 1)?.cellStyle?.borders?.top).toBeTruthy();
      // x:str with a leading apostrophe is a quoted text value.
      expect(val(wb, 0, 2, 1)).toBe("007");
      expect(cell(wb, 0, 2, 1)?.kind).toBe("s");
      expect(val(wb, 0, 2, 2)).toBe("1234.5");
      expect(cell(wb, 0, 2, 2)?.numFmt).toBe(2);
      expect(val(wb, 0, 2, 3)).toBe("45366");
      // A text-formatted cell keeps its leading zeros; a bare x:num uses the displayed separators.
      expect(val(wb, 0, 3, 1)).toBe("00123");
      expect(val(wb, 0, 3, 2)).toBe("2.25");
      expect(cell(wb, 0, 3, 3)?.formula).toBe("B2+B3");
      expect(Number(val(wb, 0, 3, 3))).toBeCloseTo(1236.75);
      expect(val(wb, 0, 4, 1)).toBe("Long text that overflows");
      expect(cell(wb, 0, 4, 3)?.kind).toBe("b");
      expect(cell(wb, 0, 5, 3)?.kind).toBe("e");
      expect(cell(wb, 0, 5, 1)?.cellStyle?.italic).toBe(true);
      expect(cell(wb, 0, 5, 1)?.cellStyle?.color?.toLowerCase()).toBe("#ff0000");
      expect(cell(wb, 0, 5, 1)?.cellStyle?.wrap).toBe(true);
      // mso-ignore:colspan is overflow, not a merge.
      expect(s.merges).toEqual([{ r1: 5, c1: 1, r2: 5, c2: 2 }]);
      expect(val(wb, 0, 6, 1)).toBe("edu.codemao.cn");
      // The hidden sizing row at the end is dropped, and widths come from <col>.
      expect(s.hiddenRows?.size ?? 0).toBe(0);
      expect(s.colWidths?.get(1)).toBeGreaterThan(s.colWidths?.get(2) ?? 0);
      expect(s.colWidths?.get(3)).toBe(s.colWidths?.get(2));
      expect(wb.definedNames?.get("Total")).toBe("'Ventes 2024'!$B$2:$B$3");
    });
  }
});

const GENERIC = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Report</title></head><body>
<h1>Account statement</h1>
<table border="1">
 <caption>Transactions</caption>
 <thead><tr><th>Id</th><th>Amount</th><th>Share</th><th>Date</th><th>Ref</th></tr></thead>
 <tbody>
  <tr><td>00123</td><td>1,234.50</td><td>12.5%</td><td>2024-03-15</td><td>12345678901234567890</td></tr>
  <tr><td>42</td><td>$99.00</td><td>3%</td><td>2024-03-16 10:30</td><td><b>Bold</b> and <i>italic</i></td></tr>
  <tr><td rowspan="2">Span</td><td>line one<br>line two</td><td>&nbsp;</td><td>TRUE</td><td><a href="https://example.com/x">link</a></td></tr>
  <tr><td>-7</td></tr>
 </tbody>
</table>
<table><tr><td><table><tr><td>inner</td><td>1</td></tr></table></td></tr></table>
<p>Generated by the bank</p>
</body></html>`;

describe("generic HTML tables", () => {
  for (const useDom of [true, false]) {
    it(`types values sensibly, one sheet per data table${useDom ? "" : " without a DOM"}`, () => {
      const wb = readHtmlWorkbook(GENERIC, { useDom });
      expect(wb.sheets.map((s) => s.name)).toEqual(["Transactions", "Sheet2"]);
      // The page title sits above the table, the caption heads it as a merged row.
      expect(val(wb, 0, 1, 1)).toBe("Account statement");
      expect(val(wb, 0, 2, 1)).toBe("Transactions");
      expect(wb.sheets[0]!.merges).toContainEqual({ r1: 2, c1: 1, r2: 2, c2: 5 });
      expect(cell(wb, 0, 3, 1)?.cellStyle?.bold).toBe(true);
      expect(val(wb, 0, 4, 1)).toBe("00123");
      expect(cell(wb, 0, 4, 1)?.kind).toBe("s");
      expect(val(wb, 0, 4, 2)).toBe("1234.5");
      expect(cell(wb, 0, 4, 2)?.numFmt).toBe(4);
      expect(val(wb, 0, 4, 3)).toBe("0.125");
      expect(cell(wb, 0, 4, 4)?.kind).toBe("n");
      expect(Number(val(wb, 0, 4, 4))).toBe(45366);
      expect(val(wb, 0, 4, 5)).toBe("12345678901234567890");
      expect(val(wb, 0, 5, 2)).toBe("99");
      expect(Number(val(wb, 0, 5, 4))).toBeCloseTo(45367 + 10.5 / 24);
      expect(cell(wb, 0, 5, 5)?.richRuns?.length).toBeGreaterThan(1);
      expect(wb.sheets[0]!.merges).toContainEqual({ r1: 6, c1: 1, r2: 7, c2: 1 });
      expect(val(wb, 0, 6, 2)).toBe("line one\nline two");
      expect(val(wb, 0, 6, 3) ?? "").toBe("");
      expect(cell(wb, 0, 6, 4)?.kind).toBe("b");
      expect(cell(wb, 0, 6, 5)?.link?.href).toBe("https://example.com/x");
      // A rowspan shifts the next row's cells right.
      expect(val(wb, 0, 7, 2)).toBe("-7");
      // Table borders reach the cells.
      expect(cell(wb, 0, 4, 1)?.cellStyle?.borders?.left).toBeTruthy();
      // The layout table around the inner one is not a sheet; the text after the last table is kept.
      expect(val(wb, 1, 1, 1)).toBe("inner");
      expect(val(wb, 1, 2, 1)).toBe("Generated by the bank");
    });
  }

  it("reads European numbers and dates when the table writes them", () => {
    const wb = readHtmlWorkbook("<table><tr><td>1 234,56</td><td>0,5</td><td>31.12.2023</td></tr><tr><td>12,00</td><td>3,75</td><td>x</td></tr></table>");
    expect(val(wb, 0, 1, 1)).toBe("1234.56");
    expect(val(wb, 0, 1, 2)).toBe("0.5");
    expect(Number(val(wb, 0, 1, 3))).toBe(45291);
  });

  it("honours the declared charset, BOMs and Windows-1252", () => {
    const body = (s: string): number[] => [...s].map((c) => c.charCodeAt(0));
    const cp1252 = new Uint8Array([...body('<html><head><meta charset="windows-1252"></head><body><table><tr><td>'), 0x80, 0x20, 0xe9, ...body("</td></tr></table></body></html>")]);
    expect(val(readHtmlWorkbook(cp1252), 0, 1, 1)).toBe("\u20ac \u00e9");
    // No declaration and invalid UTF-8: Windows-1252.
    const bare = new Uint8Array([...body("<table><tr><td>caf"), 0xe9, ...body("</td></tr></table>")]);
    expect(val(readHtmlWorkbook(bare), 0, 1, 1)).toBe("caf\u00e9");
    const u16 = "<table><tr><td>\u65e5\u672c</td></tr></table>";
    const le = new Uint8Array(2 + u16.length * 2);
    le[0] = 0xff; le[1] = 0xfe;
    for (let i = 0; i < u16.length; i++) { le[2 + i * 2] = u16.charCodeAt(i) & 0xff; le[3 + i * 2] = u16.charCodeAt(i) >> 8; }
    expect(val(readHtmlWorkbook(le), 0, 1, 1)).toBe("\u65e5\u672c");
    expect(val(readHtmlWorkbook(new Uint8Array([0xef, 0xbb, 0xbf, ...strToU8("<table><tr><td>\u00e9t\u00e9</td></tr></table>")])), 0, 1, 1)).toBe("\u00e9t\u00e9");
  });

  it("opens through readWorkbook, whatever the name", () => {
    const wb = readWorkbook(strToU8(GENERIC));
    expect(wb.importedFrom).toBe("html");
    expect(wb.kind).toBe("xlsx");
  });
});

describe("multi-sheet web pages", () => {
  const qp = (s: string): string => s.replace(/=/g, "=3D");
  const sheet = (v: string): string => `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table><tr><td x:num>${v}</td></tr></table></body></html>`;
  const MHT = [
    "MIME-Version: 1.0",
    'Content-Type: multipart/related; boundary="----=_NextPart_01"',
    "",
    "------=_NextPart_01",
    "Content-Location: file:///C:/Book1.htm",
    "Content-Transfer-Encoding: quoted-printable",
    'Content-Type: text/html; charset="us-ascii"',
    "",
    qp(`<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><xml><x:ExcelWorkbook><x:ExcelWorksheets>
<x:ExcelWorksheet><x:Name>First</x:Name><x:WorksheetSource HRef="Book1_files/sheet001.htm"/></x:ExcelWorksheet>
<x:ExcelWorksheet><x:Name>Second</x:Name><x:WorksheetSource HRef="Book1_files/sheet002.htm"/></x:ExcelWorksheet>
<x:ExcelWorksheet><x:Name>Missing</x:Name><x:WorksheetSource HRef="Book1_files/sheet003.htm"/></x:ExcelWorksheet>
</x:ExcelWorksheets><x:ActiveSheet>1</x:ActiveSheet></x:ExcelWorkbook></xml></head><frameset rows="*,39"><frame src="Book1_files/sheet001.htm"></frameset></html>`),
    "",
    "------=_NextPart_01",
    "Content-Location: file:///C:/Book1_files/sheet001.htm",
    "Content-Transfer-Encoding: quoted-printable",
    "Content-Type: text/html",
    "",
    qp(sheet("11")),
    "",
    "------=_NextPart_01",
    "Content-Location: file:///C:/Book1_files/sheet002.htm",
    "Content-Transfer-Encoding: base64",
    "Content-Type: text/html",
    "",
    btoa(sheet("22")),
    "",
    "------=_NextPart_01--",
  ].join("\r\n");

  it("reads every sheet an MHTML file carries and keeps the names of the missing ones", () => {
    const wb = readWorkbook(strToU8(MHT));
    expect(wb.sheets.map((s) => s.name)).toEqual(["First", "Second", "Missing"]);
    expect(val(wb, 0, 1, 1)).toBe("11");
    expect(val(wb, 1, 1, 1)).toBe("22");
    expect(wb.sheets[2]!.cells.size).toBe(0);
    expect(wb.activeSheet).toBe(1);
  });
});

describe("markup tokenizer", () => {
  it("closes cells and rows the way HTML implies", () => {
    const root = parseMarkup("<table><tr><td>a<td><b>b</b><tr><td>c</table>", false);
    const table = root.kids[0] as { kids: { tag: string; kids: unknown[] }[] };
    expect(table.kids.map((k) => k.tag)).toEqual(["tr", "tr"]);
    expect(table.kids[0]!.kids).toHaveLength(2);
  });
});
