import { afterEach, beforeAll, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createSheetEditor } from "./editor";
import { setLocale } from "./i18n";
import { setRegionalLocale } from "./regional";

// A French user reads and types "=SOMME(A1;0,5)" and "1 234,5"; the file keeps SUM(A1,0.5)
// and 1234.5, exactly as Excel in French does.

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
});

const frame = () => new Promise<void>((res) => setTimeout(res, 40));
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function book(): Uint8Array {
  const body =
    `<sheetData><row r="1"><c r="A1"><v>1.5</v></c></row>` +
    `<row r="2"><c r="A2"><f>SUM(A1,1)</f><v>2.5</v></c></row>` +
    `<row r="5"><c r="A5" s="1"><v>46351</v></c></row></sheetData>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${SS}" xmlns:r="${R}">${body}</worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${SS}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>`),
  });
}

const cellInput = (host: HTMLElement, rc: string) => host.querySelector(`td[data-rc="${rc}"] input`) as HTMLInputElement;

async function type(host: HTMLElement, rc: string, text: string): Promise<void> {
  const input = cellInput(host, rc);
  input.focus();
  input.value = text;
  input.dispatchEvent(new Event("input"));
  input.blur();
  await frame();
}

describe("a workbook edited in French", () => {
  it("shows and takes French numbers, dates and formulas while the file stays invariant", async () => {
    setLocale("fr");
    const host = document.createElement("div");
    document.body.appendChild(host);
    const editor = createSheetEditor(host, book(), { filename: "s.xlsx" });
    await frame();

    expect(cellInput(host, "1:1").value).toBe("1,5");
    expect(cellInput(host, "2:1").value).toBe("2,5");
    expect(cellInput(host, "5:1").value).toBe("25/11/2026");
    cellInput(host, "2:1").focus();
    expect(cellInput(host, "2:1").value).toBe("=SOMME(A1;1)");
    cellInput(host, "2:1").blur();
    await frame();

    await type(host, "3:1", "=SOMME(A1;0,5)");
    await type(host, "4:1", "1 234,5");
    expect(cellInput(host, "3:1").value).toBe("2");
    expect(cellInput(host, "4:1").value).toBe("1234,5");

    const inputs = new Map(editor.cellInputs().map((i) => [`${i.r}:${i.c}`, i.input]));
    expect(inputs.get("3:1")).toBe("=SUM(A1,0.5)");
    expect(inputs.get("4:1")).toBe("1234.5");
    const xml = strFromU8(unzipSync(await editor.getBytes())["xl/worksheets/sheet1.xml"]!);
    expect(xml).toContain("<f>SUM(A1,0.5)</f>");
    expect(xml).toContain("<v>1234.5</v>");
  });

  it("stays English when the host forces English conventions", async () => {
    setLocale("fr");
    setRegionalLocale("en");
    const host = document.createElement("div");
    document.body.appendChild(host);
    createSheetEditor(host, book(), { filename: "s.xlsx" });
    await frame();
    expect(cellInput(host, "1:1").value).toBe("1.5");
    cellInput(host, "2:1").focus();
    expect(cellInput(host, "2:1").value).toBe("=SUM(A1,1)");
  });
});
