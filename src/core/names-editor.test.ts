import { afterEach, beforeAll, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createSheetEditor } from "./editor";
import { setLocale } from "./i18n";
import { setRegionalLocale } from "./regional";
import { readWorkbook } from "./workbook";

// The Name Manager and the name box, driven as a user would: toolbar button, dialog fields, Enter
// in the name box, Ctrl+Z. Refers-to reads and types the local way and is stored invariant.

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});
afterEach(() => {
  setRegionalLocale(null);
  setLocale("en");
  document.body.textContent = "";
});

// Each case drives a whole editor and saves; give it room on a loaded machine.
const SLOW = { timeout: 30000 };
const frame = () => new Promise<void>((res) => setTimeout(res, 30));
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function book(names = ""): Uint8Array {
  const body = `<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>4</v></c></row>` +
    `<row r="4"><c r="B4"><f>SUM(Old)</f><v>0</v></c></row></sheetData>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>${names}<calcPr calcId="1"/></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}" xmlns:r="${R}">${body}</worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${SS}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`),
  });
}

function odsBook(): Uint8Array {
  const cell = (v: number) => `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
  const content = `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row>${cell(3)}</table:table-row><table:table-row>${cell(4)}</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
    "content.xml": strToU8(content),
  } as Record<string, Uint8Array | [Uint8Array, { level: 0 }]>);
}

async function mount(bytes = book(), fileName = "s.xlsx") {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const ed = createSheetEditor(host, bytes, { fileName });
  await frame();
  return { host, ed };
}

const nameBox = (host: HTMLElement) => host.querySelector(".sheetedit-fxref") as HTMLInputElement;
async function typeInNameBox(host: HTMLElement, text: string): Promise<void> {
  const box = nameBox(host);
  box.focus();
  box.value = text;
  box.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
  await frame();
}
const cellInput = (host: HTMLElement, rc: string) => host.querySelector(`td[data-rc="${rc}"] input`) as HTMLInputElement;
async function typeCell(host: HTMLElement, rc: string, text: string): Promise<void> {
  const input = cellInput(host, rc);
  input.focus();
  input.value = text;
  input.dispatchEvent(new Event("input"));
  input.blur();
  await frame();
}
const selected = (host: HTMLElement) => [...host.querySelectorAll("td.sheetedit-sel")].map((td) => (td as HTMLElement).dataset.rc).sort();

/** Open the Name Manager from its toolbar button. */
function openManager(host: HTMLElement, title = "Name Manager"): HTMLElement {
  const btn = [...host.querySelectorAll("button")].find((b) => b.title === title)!;
  btn.click();
  return document.querySelector(".sheetedit-names") as HTMLElement;
}
const role = (dlg: HTMLElement, r: string) => dlg.querySelector(`[data-role="${r}"]`) as HTMLButtonElement;
const field = (dlg: HTMLElement, k: string) => dlg.querySelector(`[data-field="${k}"]`) as HTMLInputElement;
const formOk = (dlg: HTMLElement) => dlg.querySelector('.sheetedit-names-form [data-role="ok"]') as HTMLButtonElement;
const rowRef = (dlg: HTMLElement, name: string) =>
  [...dlg.querySelectorAll("tr[data-name]")].find((tr) => (tr as HTMLElement).dataset.name === name)?.querySelector(".sheetedit-names-ref")?.textContent;
const ctrl = async (host: HTMLElement, key: string, shift = false) => {
  cellInput(host, "1:1").focus();
  cellInput(host, "1:1").dispatchEvent(new KeyboardEvent("keydown", { key, ctrlKey: true, shiftKey: shift, bubbles: true, cancelable: true }));
  await frame();
};
const wbXml = async (ed: { getBytes: () => Promise<Uint8Array> }) => strFromU8(unzipSync(await ed.getBytes())["xl/workbook.xml"]!);

describe("the Name Manager in English", SLOW, () => {
  it("defines a name from the selection, saves it to .xlsx, reopens it and undoes it", async () => {
    const { host, ed } = await mount();
    await typeInNameBox(host, "A1:A3"); // select the range
    expect(selected(host)).toEqual(["1:1", "2:1", "3:1"]);
    const dlg = openManager(host);
    expect(dlg.getAttribute("role")).toBe("dialog");
    role(dlg, "new").click();
    expect(field(dlg, "refersTo").value).toBe("=S!$A$1:$A$3"); // absolute, sheet-qualified
    field(dlg, "name").value = "Sales";
    formOk(dlg).click();
    expect(rowRef(dlg, "Sales")).toBe("=S!$A$1:$A$3");
    expect(ed.definedNames()).toEqual({ Sales: "S!$A$1:$A$3" });
    role(dlg, "close").click();
    await frame();

    await typeCell(host, "5:1", "=SUM(sales)");
    expect(cellInput(host, "5:1").value).toBe("7");
    const xml = await wbXml(ed);
    expect(xml).toContain('<definedName name="Sales">S!$A$1:$A$3</definedName>');
    expect(xml.indexOf("<definedNames>")).toBeLessThan(xml.indexOf("<calcPr"));
    const again = readWorkbook(await ed.getBytes());
    expect(again.definedNames?.get("Sales")).toBe("S!$A$1:$A$3");

    await ctrl(host, "z"); // the formula
    await ctrl(host, "z"); // the name
    expect(ed.definedNames()).toEqual({});
    expect(await wbXml(ed)).not.toContain("Sales");
    await ctrl(host, "z", true);
    expect(ed.definedNames()).toEqual({ Sales: "S!$A$1:$A$3" });
    ed.destroy();
  });

  it("refuses bad names in the dialog and keeps it open", async () => {
    const { host, ed } = await mount(book(`<definedNames><definedName name="Old">S!$A$1:$A$2</definedName></definedNames>`));
    const dlg = openManager(host);
    role(dlg, "new").click();
    const err = dlg.querySelector(".sheetedit-names-error")!;
    for (const [name, msg] of [["A1", /cell reference/], ["two words", /no spaces/], ["old", /already exists/], ["", /Type a name/]] as const) {
      field(dlg, "name").value = name;
      formOk(dlg).click();
      expect(err.textContent).toMatch(msg);
      expect(field(dlg, "name").getAttribute("aria-invalid")).toBe("true");
    }
    expect(ed.definedNames()).toEqual({ Old: "S!$A$1:$A$2" });
    // Escape leaves the form first, then the dialog.
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect((dlg.querySelector(".sheetedit-names-form") as HTMLElement).hidden).toBe(true);
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(document.querySelector(".sheetedit-names")).toBeNull();
    ed.destroy();
  });

  it("renames a name into the formulas that use it, deletes one, and shows sheet-scoped names", async () => {
    const names = `<definedNames><definedName name="Old">S!$A$1:$A$2</definedName><definedName name="Local" localSheetId="0">S!$A$3</definedName></definedNames>`;
    const { host, ed } = await mount(book(names));
    const dlg = openManager(host);
    const localRow = [...dlg.querySelectorAll("tr[data-name]")].find((tr) => (tr as HTMLElement).dataset.name === "Local")!;
    expect(localRow.textContent).toContain("S");
    expect((localRow.querySelector('[data-role="edit"]') as HTMLButtonElement).disabled).toBe(true);
    const oldRow = [...dlg.querySelectorAll("tr[data-name]")].find((tr) => (tr as HTMLElement).dataset.name === "Old")!;
    (oldRow.querySelector('[data-role="edit"]') as HTMLButtonElement).click();
    field(dlg, "name").value = "Fresh";
    field(dlg, "refersTo").value = "=S!$A$1:$A$3";
    formOk(dlg).click();
    await frame();
    expect(ed.definedNames()).toEqual({ Fresh: "S!$A$1:$A$3" });
    expect(ed.cellInputs().find((c) => c.r === 4 && c.c === 2)?.input).toBe("=SUM(Fresh)");
    expect(cellInput(host, "4:2").value).toBe("7");
    let xml = await wbXml(ed);
    expect(xml).toContain('<definedName name="Fresh">S!$A$1:$A$3</definedName>');
    expect(xml).toContain('<definedName name="Local" localSheetId="0">S!$A$3</definedName>');
    expect(strFromU8(unzipSync(await ed.getBytes())["xl/worksheets/sheet1.xml"]!)).toContain("<f>SUM(Fresh)</f>");

    const fresh = [...dlg.querySelectorAll("tr[data-name]")].find((tr) => (tr as HTMLElement).dataset.name === "Fresh")!;
    (fresh.querySelector('[data-role="delete"]') as HTMLButtonElement).click();
    expect(ed.definedNames()).toEqual({});
    xml = await wbXml(ed);
    expect(xml).not.toContain('name="Fresh"');
    expect(xml).toContain('name="Local"');
    role(dlg, "close").click();
    await ctrl(host, "z"); // back to Fresh
    await ctrl(host, "z"); // back to Old, formula too
    expect(ed.definedNames()).toEqual({ Old: "S!$A$1:$A$2" });
    expect(ed.cellInputs().find((c) => c.r === 4 && c.c === 2)?.input).toBe("=SUM(Old)");
    ed.destroy();
  });
});

describe("the Name Manager in French", SLOW, () => {
  it("takes a refers-to typed the French way, stores it invariant and shows it French", async () => {
    setLocale("fr");
    const { host, ed } = await mount();
    const dlg = openManager(host, "Gestionnaire de noms");
    role(dlg, "new").click();
    field(dlg, "name").value = "Total_TTC";
    field(dlg, "refersTo").value = "=SOMME(S!$A$1:$A$2;0,5)";
    formOk(dlg).click();
    expect(ed.definedNames()).toEqual({ Total_TTC: "SUM(S!$A$1:$A$2,0.5)" });
    expect(rowRef(dlg, "Total_TTC")).toBe("=SOMME(S!$A$1:$A$2;0,5)");
    role(dlg, "new").click();
    field(dlg, "name").value = "Taux";
    field(dlg, "refersTo").value = "0,2";
    formOk(dlg).click();
    expect(ed.definedNames().Taux).toBe("0.2");
    expect(rowRef(dlg, "Taux")).toBe("=0,2");
    role(dlg, "close").click();
    await frame();
    await typeCell(host, "6:1", "=Total_TTC*(1+Taux)");
    expect(cellInput(host, "6:1").value).toBe("4,2"); // (1+2+0.5)*1.2
    const xml = await wbXml(ed);
    expect(xml).toContain('<definedName name="Total_TTC">SUM(S!$A$1:$A$2,0.5)</definedName>');
    expect(xml).toContain('<definedName name="Taux">0.2</definedName>');
    // Reopened, the dialog shows the stored names French again.
    const { host: host2 } = await mount(await ed.getBytes());
    const dlg2 = openManager(host2, "Gestionnaire de noms");
    expect(rowRef(dlg2, "Total_TTC")).toBe("=SOMME(S!$A$1:$A$2;0,5)");
    ed.destroy();
  });
});

describe("the name box", SLOW, () => {
  it("jumps to a name, shows it for its exact range, and defines a new one from the selection", async () => {
    const { host, ed } = await mount(book(`<definedNames><definedName name="Old">S!$A$1:$A$2</definedName></definedNames>`));
    await typeInNameBox(host, "old");
    expect(selected(host)).toEqual(["1:1", "2:1"]);
    expect(nameBox(host).value).toBe("Old");
    await typeInNameBox(host, "A2:A3");
    expect(nameBox(host).value).toBe("A2");
    await typeInNameBox(host, "Pair");
    expect(ed.definedNames().Pair).toBe("S!$A$2:$A$3");
    expect(nameBox(host).value).toBe("Pair");
    await typeInNameBox(host, "bad name");
    expect(ed.definedNames()).toEqual({ Old: "S!$A$1:$A$2", Pair: "S!$A$2:$A$3" });
    expect(host.querySelector(".sheetedit-notice")?.textContent).toContain("bad name");
    await ctrl(host, "z");
    expect(ed.definedNames()).toEqual({ Old: "S!$A$1:$A$2" });
    ed.destroy();
  });

  it("defines a name in an .ods and writes it as a named range", async () => {
    const { host, ed } = await mount(odsBook(), "d.ods");
    await typeInNameBox(host, "A1:A2");
    await typeInNameBox(host, "Pair");
    expect(ed.definedNames()).toEqual({ Pair: "Data!$A$1:$A$2" });
    const content = strFromU8(unzipSync(await ed.getBytes())["content.xml"]!);
    expect(content).toMatch(/<table:named-range [^>]*table:name="Pair"[^>]*table:cell-range-address="\$Data\.\$A\$1:\.\$A\$2"/);
    const again = readWorkbook(await ed.getBytes());
    expect(again.definedNames?.get("Pair")).toBe("Data!A1:A2");
    ed.destroy();
  });
});
