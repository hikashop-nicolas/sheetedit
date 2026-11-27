// The toolbar font menu over fontshelf (detection and catalog tests live in fontshelf).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createFontMenu } from "./font-menu";
import type { FontOptions } from "../fonts";
import { setLocale } from "../i18n";
import { getFontLocale } from "fontshelf";

describe("the font menu", () => {
  beforeEach(() => {
    (globalThis as unknown as Record<symbol, unknown>)[Symbol.for("fontshelf.state.v1")] = undefined;
    (globalThis as unknown as Record<symbol, unknown>)[Symbol.for("fontshelf.loaded.v1")] = undefined;
    const mem = new Map<string, string>();
    vi.stubGlobal("localStorage", { getItem: (k: string) => mem.get(k) ?? null, setItem: (k: string, v: string) => void mem.set(k, v) });
  });
  afterEach(() => { vi.unstubAllGlobals(); document.body.replaceChildren(); });

  const mount = (fontOptions?: FontOptions, current = "Calibri") => {
    const wrap = document.createElement("div");
    document.body.appendChild(wrap);
    const button = document.createElement("button");
    wrap.appendChild(button);
    const apply = vi.fn();
    const m = createFontMenu({
      wrap, button, docFonts: ["Calibri", "Klee One"], fontOptions, current: () => current, apply,
      bindMenu: (b, menu) => b.addEventListener("click", () => { menu.hidden = !menu.hidden; }),
      closeMenus: () => { m.menu.hidden = true; },
    });
    const items = () => Array.from(m.menu.querySelectorAll<HTMLButtonElement>(".sheetedit-font-item"));
    const labels = () => items().map((b) => b.textContent ?? "");
    return { m, wrap, button, apply, items, labels };
  };

  it("offers web fonts only when the host turns them on", () => {
    expect(mount().labels().some((l) => l.startsWith("Get more fonts"))).toBe(false);
    expect(mount({ remote: false }).labels().some((l) => l.startsWith("Get more fonts"))).toBe(false);
    expect(mount({ remote: { provider: "google" } }).labels().some((l) => l.startsWith("Get more fonts"))).toBe(true);
    expect(mount({ local: false }).labels().some((l) => l.startsWith("Show my fonts"))).toBe(false);
  });

  it("lists the workbook's fonts first, checks the active one and applies a pick", () => {
    const { m, items, apply } = mount();
    expect(m.fonts().slice(0, 2)).toEqual(["Calibri", "Klee One"]);
    const cal = items().find((b) => b.dataset.font === "Calibri")!;
    expect(cal.getAttribute("role")).toBe("menuitemradio");
    expect(cal.getAttribute("aria-checked")).toBe("true");
    items().find((b) => b.dataset.font === "Klee One")!.click();
    expect(apply).toHaveBeenCalledWith("Klee One");
    expect(m.menu.hidden).toBe(true);
  });

  it("shows each font in its own face", () => {
    const { items } = mount();
    expect(items().find((b) => b.dataset.font === "Klee One")!.style.fontFamily).toContain("Klee One");
  });

  it("asks for consent inside the editor before any catalog, without applying anything", () => {
    const { items, wrap, apply } = mount({ remote: { provider: "google" } });
    items().find((b) => b.textContent?.startsWith("Get more fonts"))!.click();
    expect(apply).not.toHaveBeenCalled();
    expect(wrap.querySelector(".fontshelf-consent")?.textContent).toContain("Google Fonts");
  });

  it("moves with the arrow keys and closes on Escape, back to its button", () => {
    const { m, button, items } = mount();
    button.click(); // detail 0: opened from the keyboard
    expect(m.menu.hidden).toBe(false);
    expect(document.activeElement).toBe(items().find((b) => b.dataset.font === "Calibri"));
    m.menu.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    expect(document.activeElement).toBe(items()[1]);
    m.menu.dispatchEvent(new KeyboardEvent("keydown", { key: "End", bubbles: true }));
    expect(document.activeElement).toBe(items().at(-1));
    m.menu.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(m.menu.hidden).toBe(true);
    expect(document.activeElement).toBe(button);
  });

  it("tells assistive tech whether it is open", async () => {
    const { button } = mount();
    expect(button.getAttribute("aria-haspopup")).toBe("menu");
    button.click();
    await Promise.resolve();
    expect(button.getAttribute("aria-expanded")).toBe("true");
  });

  it("loads a font file from the menu and applies its family", async () => {
    vi.stubGlobal("indexedDB", undefined);
    const { items, apply, wrap } = mount();
    expect(mount({ file: false }).labels()).not.toContain("Load a font file…");
    let input: HTMLInputElement | null = null;
    const click = vi.spyOn(HTMLInputElement.prototype, "click").mockImplementation(function (this: HTMLInputElement) { input = this; });
    items().find((b) => b.textContent === "Load a font file…")!.click();
    expect(input!.type).toBe("file");
    expect(wrap.contains(input)).toBe(true);
    // An sfnt with no tables: the family comes from the file name.
    const bytes = new Uint8Array(12);
    new DataView(bytes.buffer).setUint32(0, 0x00010000);
    Object.defineProperty(input!, "files", { value: [new File([bytes], "Sheet Face.ttf")] });
    input!.dispatchEvent(new Event("change"));
    await vi.waitFor(() => expect(apply).toHaveBeenCalledWith("Sheet Face"));
    click.mockRestore();
  });

  it("follows the workbook's fonts as sheets load", () => {
    const { m } = mount();
    m.setDocFonts(["Calibri", "Klee One", "Lobster"]);
    m.refresh();
    expect(m.fonts().slice(0, 3)).toEqual(["Calibri", "Klee One", "Lobster"]);
  });

  it("speaks the editor's language", () => {
    setLocale("fr");
    expect(getFontLocale()).toBe("fr");
    expect(mount().labels()).toContain("Calibri");
    expect(Array.from(document.querySelectorAll(".sheetedit-font-group")).map((h) => h.textContent)).toContain("Dans ce document");
    setLocale("en");
    expect(getFontLocale()).toBe("en");
  });
});
