import { afterEach, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { setLocale } from "../i18n";
import { setRegionalLocale } from "../regional";
import { readWorkbook, writeWorkbook } from "../workbook";
import { setupDialogs } from "./dialogs";
import { setupFindBar } from "./findbar";
import { formulaFromLocal, inputFromLocal, listFromLocal, listToLocal, operandFromLocal, operandToLocal } from "../formula-locale";
import { getCell, type Workbook } from "../model";

// The rule dialogs speak the user's conventions both ways: typed "=ET(A1>0;A1<10)", "3,5" or
// "25/11/2026" in French is stored AND(A1>0,A1<10), 3.5 and the date serial, and shown back French.

afterEach(() => {
  setRegionalLocale(null);
  setLocale("en");
});

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function book(): Uint8Array {
  const body = `<sheetData><row r="1"><c r="A1"><v>5</v></c><c r="B1" s="1"><v>46351</v></c></row></sheetData>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}" xmlns:r="${R}">${body}</worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${SS}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs><dxfs count="0"/></styleSheet>`),
  });
}

type Rect = { r1: number; c1: number; r2: number; c2: number };

function harness(wb: Workbook) {
  const wrap = document.createElement("div");
  document.body.appendChild(wrap);
  let sel: Rect = { r1: 1, c1: 1, r2: 3, c2: 1 };
  const dialogs = setupDialogs({
    wb, wrap, active: () => 0, getSelRect: () => sel,
    mark: () => {}, renderGrid: () => {}, refreshShapes: () => {},
    viewBox: () => ({ fromCol: 1, fromRow: 1, fromColOff: 0, fromRowOff: 0, toCol: 2, toRow: 2, toColOff: 0, toRowOff: 0 }),
    applySparkline: () => {},
  });
  const field = (key: string) => wrap.querySelector(`[data-field="${key}"]`) as HTMLInputElement | HTMLSelectElement;
  const fill = (vals: Record<string, string>) => {
    for (const [k, v] of Object.entries(vals)) {
      field(k).value = v;
      field(k).dispatchEvent(new Event("change"));
    }
  };
  const ok = () => (wrap.querySelector('[data-role="ok"]') as HTMLButtonElement).click();
  const cancel = () => (wrap.querySelector('[data-role="cancel"]') as HTMLButtonElement).click();
  return { dialogs, field, fill, ok, cancel, select: (r: Rect) => (sel = r) };
}

const sheetXml = (wb: Workbook): string => strFromU8(unzipSync(writeWorkbook(wb))["xl/worksheets/sheet1.xml"]!);

describe("the data validation dialog in French", () => {
  it("stores a decimal bound and a custom formula invariant, and shows them back French", () => {
    setLocale("fr");
    const wb = readWorkbook(book());
    const h = harness(wb);
    h.dialogs.openDvDialog();
    h.fill({ type: "decimal", op: "between", v1: "3,5", v2: "1 234,5" });
    h.ok();
    expect(wb.sheets[0]!.validations![0]).toMatchObject({ type: "decimal", formula1: "3.5", formula2: "1234.5" });
    let xml = sheetXml(wb);
    expect(xml).toContain("<formula1>3.5</formula1>");
    expect(xml).toContain("<formula2>1234.5</formula2>");

    h.select({ r1: 2, c1: 1, r2: 2, c2: 1 }); // inside the rule: the dialog edits it
    h.dialogs.openDvDialog();
    expect(h.field("type").value).toBe("decimal");
    expect(h.field("v1").value).toBe("3,5");
    expect(h.field("v2").value).toBe("1234,5");
    h.fill({ type: "custom", formula: "=ET(A1>0;A1<10,5)" });
    h.ok();
    expect(wb.sheets[0]!.validations).toHaveLength(1);
    expect(wb.sheets[0]!.validations![0]).toMatchObject({ type: "custom", formula1: "AND(A1>0,A1<10.5)" });
    xml = sheetXml(wb);
    expect(xml).toContain("<formula1>AND(A1&gt;0,A1&lt;10.5)</formula1>");
    h.dialogs.openDvDialog();
    expect(h.field("formula").value).toBe("=ET(A1>0;A1<10,5)");
    h.cancel();
  });

  it("stores a date bound as its serial and a list typed with ;", () => {
    setLocale("fr");
    const wb = readWorkbook(book());
    const h = harness(wb);
    h.select({ r1: 1, c1: 2, r2: 4, c2: 2 });
    h.dialogs.openDvDialog();
    h.fill({ type: "date", op: "greaterThanOrEqual", v1: "25/11/2026" });
    h.ok();
    expect(wb.sheets[0]!.validations![0]!.formula1).toBe("46351");
    expect(sheetXml(wb)).toContain("<formula1>46351</formula1>");
    h.dialogs.openDvDialog();
    expect(h.field("v1").value).toBe("25/11/2026");
    h.cancel();

    h.select({ r1: 1, c1: 3, r2: 4, c2: 3 });
    h.dialogs.openDvDialog();
    h.fill({ type: "list", values: "oui;non;1,5" });
    h.ok();
    expect(sheetXml(wb)).toContain('<formula1>"oui,non,1.5"</formula1>');
    h.dialogs.openDvDialog();
    expect(h.field("values").value).toBe("oui;non;1,5");
    // The list's label names the separator it is read with: ";" here, "," with English conventions.
    const label = () => h.field("values").closest("label, .sheetedit-dialog-row, div")?.textContent ?? "";
    expect(label()).toContain("points-virgules");
    h.cancel();
    setRegionalLocale("en");
    h.dialogs.openDvDialog();
    expect(label()).toContain("virgules");
    expect(label()).not.toContain("points-virgules");
    expect(h.field("values").value).toBe("oui,non,1.5");
    h.cancel();
  });
});

describe("the conditional formatting dialog in French", () => {
  it("stores a highlight value and a formula rule invariant, and shows them back French", () => {
    setLocale("fr");
    const wb = readWorkbook(book());
    const h = harness(wb);
    h.dialogs.openCfDialog();
    h.fill({ kind: "cellIs", operator: "between", value: "2,5", value2: "1 000" });
    h.ok();
    let xml = sheetXml(wb);
    expect(xml).toMatch(/<formula>2\.5<\/formula><formula>1000<\/formula>/);
    h.dialogs.openCfDialog();
    expect(h.field("value").value).toBe("2,5");
    expect(h.field("value2").value).toBe("1000");
    h.fill({ kind: "expression", formula: "=ET(A1>0;A1<10)" });
    h.ok();
    xml = sheetXml(wb);
    expect(xml).toContain("<formula>AND(A1&gt;0,A1&lt;10)</formula>");
    expect(xml).not.toContain("<formula>2.5</formula>");
    h.dialogs.openCfDialog();
    expect(h.field("kind").value).toBe("expression");
    expect(h.field("formula").value).toBe("=ET(A1>0;A1<10)");
    h.cancel();
  });

  it("quotes a text value and reads dates on date cells", () => {
    setLocale("fr");
    const wb = readWorkbook(book());
    const h = harness(wb);
    h.dialogs.openCfDialog();
    h.fill({ kind: "cellIs", operator: "equal", value: "oui" });
    h.ok();
    expect(sheetXml(wb)).toContain("<formula>\"oui\"</formula>");
    h.dialogs.openCfDialog();
    expect(h.field("value").value).toBe("oui");
    h.cancel();

    h.select({ r1: 1, c1: 2, r2: 3, c2: 2 });
    h.dialogs.openCfDialog();
    h.fill({ kind: "cellIs", operator: "greaterThan", value: "25/11/2026" });
    h.ok();
    expect(sheetXml(wb)).toContain("<formula>46351</formula>");
    h.dialogs.openCfDialog();
    expect(h.field("value").value).toBe("25/11/2026");
    h.cancel();
  });

  it("keeps English behaviour in English", () => {
    const wb = readWorkbook(book());
    const h = harness(wb);
    h.dialogs.openCfDialog();
    h.fill({ kind: "expression", formula: "AND(A1>0,A1<10)" });
    h.ok();
    expect(sheetXml(wb)).toContain("<formula>AND(A1&gt;0,A1&lt;10)</formula>");
    h.dialogs.openDvDialog();
    h.fill({ type: "list", values: "yes, no" });
    h.ok();
    expect(sheetXml(wb)).toContain('<formula1>"yes,no"</formula1>');
  });
});

describe("rule operands", () => {
  it("convert both ways", () => {
    setRegionalLocale("fr");
    expect(operandFromLocal("3,5")).toBe("3.5");
    expect(operandFromLocal("=A1*2,5")).toBe("A1*2.5");
    expect(operandFromLocal("$B$2")).toBe("$B$2");
    expect(operandFromLocal("13:30", { time: true })).toBe("0.5625");
    expect(operandToLocal("0.5625", { time: true })).toBe("13:30");
    expect(operandToLocal("SUM(A1,1.5)")).toBe("=SOMME(A1;1,5)");
    expect(listFromLocal("a; b ;2,5")).toEqual(["a", "b", "2.5"]);
    expect(listToLocal(["a", "2.5"])).toBe("a;2,5");
    expect(formulaFromLocal("ET(1;2)")).toBe("AND(1,2)");
  });
});

describe("find and replace in French", () => {
  it("finds a French function name and replaces it in the stored formula", () => {
    setLocale("fr");
    const wb = readWorkbook(book());
    const sheet = wb.sheets[0]!;
    getCell(sheet, 1, 1)!.formula = "SUM(B1,1.5)";
    const host = document.createElement("div");
    const batches: { r: number; c: number; raw: string }[][] = [];
    const bar = setupFindBar({
      container: host, beforeEl: null, getWorkbook: () => wb, getActiveSheet: () => 0,
      setActiveSheet: () => {}, focusCell: () => {}, commitValue: () => {},
      applyBatch: (ch) => batches.push(ch),
      textOf: (si, r, c) => { const cell = getCell(wb.sheets[si]!, r, c); return cell?.formula != null ? "=SOMME(B1;1,5)" : cell?.value ?? ""; },
      fromText: (t) => inputFromLocal(t),
    });
    void bar;
    const [find, replace] = host.querySelectorAll("input");
    find!.value = "SOMME";
    find!.dispatchEvent(new Event("input"));
    replace!.value = "MOYENNE";
    const all = [...host.querySelectorAll("button")].find((b) => /tout|all/i.test(b.textContent ?? b.title));
    all!.click();
    expect(batches[0]).toEqual([{ r: 1, c: 1, raw: "=AVERAGE(B1,1.5)" }]);
  });
});
