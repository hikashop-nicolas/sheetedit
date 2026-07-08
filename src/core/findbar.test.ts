import { beforeAll, describe, expect, it } from "vitest";
import { strFromU8, strToU8, zipSync } from "fflate";
import { createSheetEditor } from "./editor";
import { replaceEvery, replaceOnce } from "./ui/findbar";

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

describe("replace helpers", () => {
  it("replaceOnce is case-insensitive and single-shot", () => {
    expect(replaceOnce("Alpha beta ALPHA", "alpha", "X")).toBe("X beta ALPHA");
    expect(replaceOnce("none here", "zzz", "X")).toBe("none here");
  });
  it("replaceEvery replaces all occurrences and cannot loop on growing output", () => {
    expect(replaceEvery("aAa", "a", "aa")).toBe("aaaaaa");
    expect(replaceEvery("x-y-x", "x", "z")).toBe("z-y-z");
  });
});

const openEditor = (bytes: Uint8Array, hint?: "csv") => {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const ed = createSheetEditor(host, bytes, hint ? { formatHint: hint } : {});
  return { host, ed };
};

const barOf = (host: HTMLElement) => {
  const bar = host.querySelector(".sheetedit-findbar") as HTMLElement;
  const [find, replace] = [...bar.querySelectorAll("input")] as HTMLInputElement[];
  const btn = (label: string) => [...bar.querySelectorAll("button")].find((b) => b.textContent === label) as HTMLButtonElement;
  return { bar, find: find!, replace: replace!, btn };
};

describe("find and replace in the editor", () => {
  it("counts matches, steps with focus, replaces one and all", () => {
    const { host, ed } = openEditor(strToU8("apple,pear\napple pie,apple\n"), "csv");
    const { bar, find, replace, btn } = barOf(host);
    expect(bar.hidden).toBe(true);
    // Ctrl+F opens the bar.
    host.querySelector(".sheetedit-wrap")!.dispatchEvent(new KeyboardEvent("keydown", { key: "f", ctrlKey: true, bubbles: true }));
    expect(bar.hidden).toBe(false);

    find.value = "apple";
    find.dispatchEvent(new Event("input", { bubbles: true }));
    expect(bar.querySelector(".sheetedit-findcount")!.textContent).toBe("1 / 3");

    replace.value = "kiwi";
    btn("Replace").click(); // replaces A1
    expect(ed.getText()).toBe("kiwi,pear\napple pie,apple\n");

    btn("Replace all").click();
    expect(ed.getText()).toBe("kiwi,pear\nkiwi pie,kiwi\n");
    expect(bar.querySelector(".sheetedit-findcount")!.textContent).toContain("No matches");
    ed.destroy();
    host.remove();
  });

  it("finds across sheets and switches the active sheet when navigating", () => {
    const wbXml = `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="rId1"/><sheet name="Two" sheetId="2" r:id="rId2"/></sheets></workbook>`;
    const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`;
    const ws = (v: string) => `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${v}</t></is></c></row></sheetData></worksheet>`;
    const bytes = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "xl/workbook.xml": strToU8(wbXml),
      "xl/_rels/workbook.xml.rels": strToU8(rels),
      "xl/worksheets/sheet1.xml": strToU8(ws("hello")),
      "xl/worksheets/sheet2.xml": strToU8(ws("target here")),
    });
    const { host, ed } = openEditor(bytes);
    const { find, btn } = barOf(host);
    find.value = "target";
    find.dispatchEvent(new Event("input", { bubbles: true }));
    btn("›").click(); // wraps to the only match, on sheet Two
    const selectedTab = host.querySelector('.sheetedit-tab[aria-selected="true"]')!;
    expect(selectedTab.textContent).toBe("Two");
    expect((document.activeElement as HTMLInputElement)?.value).toBe("target here");
    ed.destroy();
    host.remove();
  });

  it("replace-all edits survive a save", async () => {
    const { host, ed } = openEditor(strToU8("total 10,total 20\n"), "csv");
    const { find, replace, btn } = barOf(host);
    find.value = "total";
    find.dispatchEvent(new Event("input", { bubbles: true }));
    replace.value = "sum";
    btn("Replace all").click();
    expect(strFromU8(await ed.getBytes())).toBe("sum 10,sum 20\n");
    ed.destroy();
    host.remove();
  });
});
