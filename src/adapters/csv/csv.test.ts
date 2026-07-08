import { describe, expect, it } from "vitest";
import { strFromU8, strToU8 } from "fflate";
import { createSheetEditor } from "../../core/editor";
import { readCsv, sniffCsvDelimiter } from "./read";
import { writeCsv } from "./write";
import { csvToXlsx } from "./convert";
import { getCell } from "../../core/model";
import { applyLineOp } from "../../core/structure";
import { readWorkbook, setCellInput, writeWorkbook } from "../../core/workbook";
import { recalc } from "../../core/recalc";

const roundtrip = (text: string) => writeCsv(readCsv(text));

describe("sniffCsvDelimiter", () => {
  it("detects comma, semicolon, tab and pipe", () => {
    expect(sniffCsvDelimiter("a,b\nc,d\n")).toBe(",");
    expect(sniffCsvDelimiter("a;b\nc;d\n")).toBe(";");
    expect(sniffCsvDelimiter("a\tb\nc\td\n")).toBe("\t");
    expect(sniffCsvDelimiter("a|b\nc|d\n")).toBe("|");
  });
  it("ignores delimiters inside quotes and prefers consistency", () => {
    expect(sniffCsvDelimiter('"a,b";c\n"d,e";f\n')).toBe(";");
    expect(sniffCsvDelimiter("name;note\nx;hello, world\ny;z\n")).toBe(";");
  });
});

describe("byte-exact round-trip (unedited)", () => {
  const cases = [
    "a,b,c\nd,e,f\n",
    "a,b\r\nc,d\r\n",
    "a,b\nc,d", // no trailing newline
    'quoted,"with, comma","with ""quotes"""\nplain,2,3\n',
    '"multi\nline",x\ny,z\n',
    "a;b;c\nd;e;f\n",
    "  spaced , kept ,c\n",
    "a,b\n\nc,d\n", // blank line preserved
    "007,1.10,eur\n", // no numeric coercion
  ];
  for (const text of cases) {
    it(JSON.stringify(text.slice(0, 24)), () => expect(roundtrip(text)).toBe(text));
  }
});

describe("editing dirties only the touched row", () => {
  it("rewrites the edited row, keeps the rest byte-exact", () => {
    const wb = readCsv("id, name ,note\r\n1,alice,hi\r\n2,bob,yo\r\n");
    setCellInput(wb.sheets[0]!, 2, 2, "ALICE");
    const out = writeCsv(wb);
    expect(out).toBe("id, name ,note\r\n1,ALICE,hi\r\n2,bob,yo\r\n");
  });
  it("quotes only when needed, with the file's own delimiter", () => {
    const wb = readCsv("a;b\nc;d\n");
    setCellInput(wb.sheets[0]!, 1, 2, "x;y");
    expect(writeCsv(wb)).toBe('a;"x;y"\nc;d\n');
  });
  it("recalc alone never dirties a row (results are not stored in CSV)", () => {
    const wb = readCsv("1,2,=SUM(A1:B1)\n");
    recalc(wb);
    expect(writeCsv(wb)).toBe("1,2,=SUM(A1:B1)\n");
  });
});

describe("formulas", () => {
  it("computes on read and keeps computing on edit", () => {
    const wb = readWorkbook(strToU8("10,20\n=SUM(A1:B1),x\n"), { formatHint: "csv" });
    const s = wb.sheets[0]!;
    expect(getCell(s, 2, 1)?.value).toBe("30");
    setCellInput(s, 1, 1, "15");
    recalc(wb);
    expect(getCell(s, 2, 1)?.value).toBe("35");
    const out = strFromU8(writeWorkbook(wb));
    expect(out).toBe("15,20\n=SUM(A1:B1),x\n");
  });
});

describe("structure ops on csv", () => {
  it("inserting a row keeps other rows byte-exact and extends ranges", () => {
    const wb = readCsv('h1,"h,2"\r\n1,2\r\n3,4\r\n=SUM(A2:A3),\r\n');
    applyLineOp(wb, 0, { axis: "row", kind: "insert", at: 3, count: 1 });
    setCellInput(wb.sheets[0]!, 3, 1, "9");
    recalc(wb);
    expect(getCell(wb.sheets[0]!, 5, 1)?.value).toBe("13");
    expect(writeCsv(wb)).toBe('h1,"h,2"\r\n1,2\r\n9\r\n3,4\r\n=SUM(A2:A4),\r\n');
  });
  it("deleting a row drops its line; refs shrink", () => {
    const wb = readCsv("1,2\n3,4\n=SUM(A1:A2),\n");
    applyLineOp(wb, 0, { axis: "row", kind: "delete", at: 2, count: 1 });
    expect(writeCsv(wb)).toBe("1,2\n=SUM(A1:A1),\n");
  });
  it("a column op reformats the whole file", () => {
    const wb = readCsv("a,b\nc,d\n");
    applyLineOp(wb, 0, { axis: "col", kind: "insert", at: 2, count: 1 });
    setCellInput(wb.sheets[0]!, 1, 2, "new");
    expect(writeCsv(wb)).toBe("a,new,b\nc,,d\n");
  });
});

describe("readWorkbook integration", () => {
  it("routes on hint and on textual fallback; binary junk still throws", () => {
    expect(readWorkbook(strToU8("a,b\nc,d\n")).kind).toBe("csv");
    expect(readWorkbook(strToU8("a\tb\nc\td\n"), { formatHint: "tsv" }).csvDelimiter).toBe("\t");
    expect(() => readWorkbook(new Uint8Array([1, 2, 3]))).toThrow();
  });
});

describe("editor in csv mode", () => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }

  it("hides styling and tabs, shows Convert, and getText round-trips edits", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    let converted: { bytes: Uint8Array; name: string } | null = null;
    const ed = createSheetEditor(host, strToU8("id;nom\n1;alice\n2;bob\n"), {
      formatHint: "csv",
      fileName: "clients.csv",
      onConvert: (bytes, name) => (converted = { bytes, name }),
    });
    // No styling cluster, no sheet tabs, but the convert button is there.
    const labels = [...host.querySelectorAll(".sheetedit-toolbar button")].map((b) => b.textContent);
    expect(labels).not.toContain("B");
    expect(host.querySelector(".sheetedit-tabs")).toBeNull();
    const convertBtn = [...host.querySelectorAll("button")].find((b) => b.textContent === "Convert to XLSX")!;
    expect(convertBtn).toBeTruthy();

    // Untouched: getText is byte-exact; row ops via the header menu still work.
    expect(ed.getText()).toBe("id;nom\n1;alice\n2;bob\n");
    const rn2 = [...host.querySelectorAll("th.rownum")].find((th) => th.textContent === "2")!;
    rn2.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    ([...document.querySelectorAll(".sheetedit-pop-item")][2] as HTMLButtonElement).click(); // delete row 2
    expect(ed.getText()).toBe("id;nom\n2;bob\n");
    expect(ed.isDirty()).toBe(true);

    convertBtn.click();
    expect(converted).not.toBeNull();
    expect(converted!.name).toBe("clients.xlsx");
    const wb2 = readWorkbook(converted!.bytes);
    expect(wb2.kind).toBe("xlsx");
    expect(getCell(wb2.sheets[0]!, 2, 2)?.value).toBe("bob");
    ed.destroy();
    host.remove();
  });
});

describe("csvToXlsx", () => {
  it("produces a workbook that reopens with the same values and formulas", () => {
    const wb = readCsv("name,total\nwidget,=2*3\ncount,5\n");
    recalc(wb);
    const xlsx = csvToXlsx(wb);
    const wb2 = readWorkbook(xlsx);
    expect(wb2.kind).toBe("xlsx");
    const s = wb2.sheets[0]!;
    expect(getCell(s, 1, 1)?.value).toBe("name");
    expect(getCell(s, 2, 2)?.formula).toBe("2*3");
    expect(getCell(s, 3, 2)?.value).toBe("5");
    // The converted workbook is editable and saves through the xlsx writer.
    setCellInput(s, 3, 2, "6");
    const out = readWorkbook(writeWorkbook(wb2));
    expect(getCell(out.sheets[0]!, 3, 2)?.value).toBe("6");
  });
});
