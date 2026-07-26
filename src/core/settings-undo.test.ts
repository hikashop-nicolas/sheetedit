import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { createSheetEditor } from "./editor";
import { readWorkbook } from "./workbook";

// Protection, page setup, panes, outline grouping and the workbook theme live on the sheet or the
// workbook rather than in a cell, so they need their own undo step. These drive the editor the way
// a user does (menus and the keyboard) rather than calling the internals, because the point is that
// Ctrl+Z reaches them.

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

const frame = () => new Promise<void>((res) => setTimeout(res, 30));
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";

function book(): Uint8Array {
  const rows = [1, 2, 3, 4, 5]
    .map((r) => `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>a${r}</t></is></c><c r="B${r}" t="inlineStr"><is><t>b${r}</t></is></c></row>`)
    .join("");
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}"><sheetData>${rows}</sheetData></worksheet>`),
    "xl/styles.xml": strToU8(`<styleSheet xmlns="${MAIN}"/>`),
  });
}

const mount = async () => {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const ed = createSheetEditor(host, book(), { fileName: "s.xlsx" });
  await frame();
  return { host, ed, done: () => { ed.destroy(); host.remove(); } };
};

/** Ctrl+Z / Ctrl+Shift+Z on the focused cell input, which is where the grid listens. */
const press = async (host: HTMLElement, key: string, shift = false): Promise<void> => {
  const input = host.querySelector<HTMLInputElement>(`td[data-rc="1:1"] input`)!;
  input.focus();
  input.dispatchEvent(new KeyboardEvent("keydown", { key, ctrlKey: true, shiftKey: shift, bubbles: true, cancelable: true }));
  await frame();
};

/** Click a toolbar button by its tooltip, then a menu item by its text. */
const menu = async (host: HTMLElement, title: string, item: string): Promise<void> => {
  const btn = [...host.querySelectorAll("button")].find((b) => b.title === title)!;
  btn.click();
  await frame();
  const entry = [...document.querySelectorAll(".sheetedit-pop-item, .sheetedit-theme-opt")].find((b) => b.textContent?.startsWith(item)) as HTMLElement;
  entry.click();
  await frame();
};

/** The model behind a live editor, read back from the bytes it would save. */
const saved = async (ed: { getBytes: () => Promise<Uint8Array> }) => readWorkbook(await ed.getBytes());

describe("undo reaches the sheet-level settings", () => {
  it("undoes and redoes a freeze", async () => {
    const { host, ed, done } = await mount();
    await menu(host, "Freeze panes", "Freeze top row");
    expect((await saved(ed)).sheets[0]!.freeze).toEqual({ rows: 1, cols: 0 });
    await press(host, "z");
    expect((await saved(ed)).sheets[0]!.freeze).toBeUndefined();
    await press(host, "z", true); // redo
    expect((await saved(ed)).sheets[0]!.freeze).toEqual({ rows: 1, cols: 0 });
    done();
  });

  it("undoes protecting a sheet", async () => {
    const { host, ed, done } = await mount();
    await menu(host, "Protection", "Protect sheet");
    // The protect dialog applies on OK.
    (document.querySelector('[data-role="ok"]') as HTMLElement).click();
    await frame();
    expect((await saved(ed)).sheets[0]!.protection?.sheet).toBe(true);
    await press(host, "z");
    expect((await saved(ed)).sheets[0]!.protection).toBeUndefined();
    done();
  });

  it("undoes a page-setup change", async () => {
    const { host, ed, done } = await mount();
    await menu(host, "Page setup", "Page setup");
    const orient = document.querySelector<HTMLSelectElement>('[data-field="orientation"]')!;
    orient.value = "landscape";
    (document.querySelector('[data-role="ok"]') as HTMLElement).click();
    await frame();
    expect((await saved(ed)).sheets[0]!.printSetup?.orientation).toBe("landscape");
    await press(host, "z");
    expect((await saved(ed)).sheets[0]!.printSetup?.orientation).not.toBe("landscape");
    done();
  });

  it("undoes a print area, which the grid also stops outlining", async () => {
    const { host, ed, done } = await mount();
    host.querySelector<HTMLInputElement>(`td[data-rc="1:1"] input`)!.focus();
    await frame();
    await menu(host, "Page setup", "Set print area");
    expect((await saved(ed)).sheets[0]!.printSetup?.printArea?.length).toBe(1);
    expect(host.querySelector("td.pa-top")).not.toBeNull();
    await press(host, "z");
    expect((await saved(ed)).sheets[0]!.printSetup?.printArea).toBeUndefined();
    expect(host.querySelector("td.pa-top")).toBeNull();
    done();
  });

  it("undoes a theme switch, putting back every colour it changed", async () => {
    const { host, ed, done } = await mount();
    const before = (await saved(ed)).theme!.colors.accent1;
    await menu(host, "Workbook theme", "Berlin");
    expect((await saved(ed)).theme!.colors.accent1).toBe("#e97b1f");
    await press(host, "z");
    expect((await saved(ed)).theme!.colors.accent1).toBe(before);
    done();
  });

  it("undoes an outline group", async () => {
    const { host, ed, done } = await mount();
    const rowHead = host.querySelector<HTMLElement>(`th.rownum[data-r="2"]`)!;
    rowHead.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    await frame();
    (([...document.querySelectorAll(".sheetedit-pop-item")].find((b) => b.textContent === "Group")) as HTMLElement).click();
    await frame();
    expect((await saved(ed)).sheets[0]!.rowOutline?.get(2)).toBe(1);
    await press(host, "z");
    expect((await saved(ed)).sheets[0]!.rowOutline?.get(2)).toBeUndefined();
    done();
  });

  it("still writes the undone state after a save has already cleared the dirty flag", async () => {
    // Save, then undo: the model has diverged from the file, so the part must be re-emitted. A
    // restored "clean" flag would leave the undone change missing from the next save.
    const { host, ed, done } = await mount();
    await menu(host, "Freeze panes", "Freeze top row");
    const first = await ed.getBytes();
    expect(new TextDecoder().decode(unzipSync(first)["xl/worksheets/sheet1.xml"]!)).toContain("<pane");
    await press(host, "z");
    const second = await ed.getBytes();
    expect(new TextDecoder().decode(unzipSync(second)["xl/worksheets/sheet1.xml"]!)).not.toContain("<pane");
    done();
  });

  it("does not hand back the same map object, which a later edit would mutate", async () => {
    const { host, ed, done } = await mount();
    const rowHead = host.querySelector<HTMLElement>(`th.rownum[data-r="2"]`)!;
    rowHead.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    await frame();
    (([...document.querySelectorAll(".sheetedit-pop-item")].find((b) => b.textContent === "Group")) as HTMLElement).click();
    await frame();
    await press(host, "z");
    await press(host, "z", true); // redo
    expect((await saved(ed)).sheets[0]!.rowOutline?.get(2)).toBe(1);
    await press(host, "z"); // and undo again: the snapshot must still be intact
    expect((await saved(ed)).sheets[0]!.rowOutline?.get(2)).toBeUndefined();
    done();
  });
});
