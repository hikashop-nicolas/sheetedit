import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";
import { readWorkbook } from "../index";

// The sheet a workbook opens on. A file records the one it was last left on, and that is where
// every other spreadsheet opens it; starting at sheet one instead drops the reader somewhere the
// author never left them - on a workbook whose first sheet is a changelog, that is the changelog.

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

const frame = () => new Promise<void>((res) => setTimeout(res, 40));
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function book(view: string, states: string[] = ["", "", ""]): Uint8Array {
  const sheets = states.map((st, i) => `<sheet name="S${i + 1}" sheetId="${i + 1}"${st} r:id="rId${i + 1}"/>`).join("");
  const rels = states.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("");
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdW" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${SS}" xmlns:r="${R}">${view}<sheets>${sheets}</sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${SS}"/>`),
  };
  states.forEach((_, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(`<worksheet xmlns="${SS}"><sheetData><row r="1"><c r="A1" t="str"><v>on sheet ${i + 1}</v></c></row></sheetData></worksheet>`);
  });
  return zipSync(files);
}

const VIEW = (tab: number) => `<bookViews><workbookView activeTab="${tab}"/></bookViews>`;

async function openedOn(bytes: Uint8Array): Promise<string> {
  const host = document.createElement("div");
  document.body.appendChild(host);
  createSheetEditor(host, bytes, { filename: "s.xlsx" });
  await frame();
  const tab = [...host.querySelectorAll<HTMLElement>(".sheetedit-tab")].find((b) => b.getAttribute("aria-selected") === "true");
  const name = tab?.textContent ?? "";
  host.remove();
  return name;
}

describe("which sheet a workbook opens on", () => {
  it("reads the tab the file was left on", () => {
    expect(readWorkbook(book(VIEW(2))).activeSheet).toBe(2);
  });

  it("opens there", async () => {
    expect(await openedOn(book(VIEW(2)))).toBe("S3");
    expect(await openedOn(book(VIEW(1)))).toBe("S2");
  });

  it("opens on the first sheet when the file names none", async () => {
    expect(await openedOn(book(""))).toBe("S1");
  });

  // A hidden sheet has no tab, so landing on it would leave the grid showing a sheet with nothing
  // selected in the strip.
  it("skips past an active sheet that is hidden", async () => {
    expect(await openedOn(book(VIEW(1), ["", ' state="hidden"', ""]))).toBe("S1");
  });
});
