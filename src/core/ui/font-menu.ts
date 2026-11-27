// The toolbar font picker: sheetedit's own button + menu, filled from fontshelf's shared state.
// Detection, the catalog, consent and its dialogs live in fontshelf.
import {
  BASE_FONTS, addDocumentFonts, buildFontGroups, canDraw, cssFamily, fontT, getFontState, isAdded,
  onFontsChange, requestFontFile, requestRemoteFont, restoreFonts, setRemoteEnabled, showFontToast, showInstalledFonts,
} from "fontshelf";
import type { FontOptions } from "../fonts";

export interface FontMenuDeps {
  wrap: HTMLElement; // hosts the menu, the dialogs and the toast
  button: HTMLButtonElement;
  docFonts: readonly string[];
  fontOptions?: FontOptions;
  /** The font of the active cell, marked as checked. */
  current(): string | undefined;
  apply(family: string): void;
  /** Wire the button to the menu (position, toggle, outside click). */
  bindMenu(button: HTMLElement, menu: HTMLElement): void;
  /** Hide every toolbar menu. */
  closeMenus(): void;
  /** A dialog is about to open / a downloaded family is about to be applied. */
  captureSel?(): void;
  restoreSel?(): void;
  notify?(text: string): void;
}

export interface FontMenu {
  menu: HTMLElement;
  /** Every font offered (document fonts first). */
  fonts(): string[];
  setDocFonts(names: readonly string[]): void;
  /** Rebuild the items (also done on open and on every fontshelf change). */
  refresh(): void;
  teardown(): void;
}

const ITEM = "sheetedit-btn sheetedit-more-item sheetedit-font-item";

export function createFontMenu(deps: FontMenuDeps): FontMenu {
  const { wrap, button } = deps;
  const remote = !!deps.fontOptions?.remote && deps.fontOptions.remote.provider === "google";
  const localOn = deps.fontOptions?.local !== false;
  const fileOn = deps.fontOptions?.file !== false;
  if (remote) setRemoteEnabled(true);
  let docFonts = [...deps.docFonts];
  let unregister = addDocumentFonts(docFonts);
  let all: string[] = [];
  let disposed = false;

  const menu = document.createElement("div");
  menu.className = "sheetedit-tb-groupmenu sheetedit-tb-listmenu sheetedit-fontmenu";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", fontT("font"));
  button.setAttribute("aria-haspopup", "menu");
  button.setAttribute("aria-expanded", "false");

  const done = (): void => {
    const hadFocus = menu.contains(document.activeElement);
    deps.closeMenus();
    if (hadFocus) button.focus();
  };
  const item = (label: string, role: "menuitem" | "menuitemradio", run: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = ITEM;
    b.setAttribute("role", role);
    b.tabIndex = -1;
    b.textContent = label;
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the cell's focus and text selection
    b.addEventListener("click", run);
    return b;
  };
  const heading = (text: string): HTMLElement => {
    const h = document.createElement("div");
    h.className = "sheetedit-font-group";
    h.setAttribute("role", "presentation");
    h.textContent = text;
    return h;
  };
  const fontItem = (name: string, label: string, cur: string): HTMLButtonElement => {
    const b = item(label, "menuitemradio", () => {
      done();
      deps.apply(name);
    });
    b.dataset.font = name;
    b.style.fontFamily = `${cssFamily(name)}, system-ui, sans-serif`; // each name in its own face
    b.setAttribute("aria-checked", name.toLowerCase() === cur ? "true" : "false");
    return b;
  };

  const render = (): void => {
    const st = getFontState();
    const groups = buildFontGroups({ docFonts, base: BASE_FONTS, extra: [...st.installed, ...st.added], canDraw: (n) => isAdded(n) || canDraw(n) });
    all = [...groups.doc.map((d) => d.name), ...groups.fonts];
    const cur = (deps.current() ?? "").toLowerCase();
    const parts: HTMLElement[] = [];
    if (groups.doc.length) {
      parts.push(heading(fontT("fontsInDocument")));
      for (const d of groups.doc) parts.push(fontItem(d.name, d.missing ? fontT("fontMissing", { name: d.name }) : d.name, cur));
    }
    if (groups.fonts.length) {
      parts.push(heading(fontT("fontsAll")));
      for (const f of groups.fonts) parts.push(fontItem(f, f, cur));
    }
    const actions: HTMLElement[] = [];
    if (localOn && !st.installedFull) actions.push(item(fontT("fontsShowMine"), "menuitem", () => { done(); void showMine(); }));
    if (fileOn) actions.push(item(fontT("fontsLoadFile"), "menuitem", () => { done(); void loadFile(); }));
    if (remote) actions.push(item(fontT("fontsGetMore"), "menuitem", () => { done(); void getMore(); }));
    if (actions.length) parts.push(heading(fontT("fontsMore")), ...actions);
    const keepScroll = menu.scrollTop;
    menu.replaceChildren(...parts);
    menu.scrollTop = keepScroll;
  };

  const notify = (text: string): void => {
    if (deps.notify) deps.notify(text);
    else showFontToast(wrap, text);
  };
  const showMine = async (): Promise<void> => {
    const before = all.length;
    await showInstalledFonts({ gesture: true });
    if (disposed) return;
    render();
    const n = all.length - before;
    notify(n > 0 ? fontT("fontsFound", { n }) : fontT("fontsNoneFound"));
  };
  const picked = (family: string | null): void => {
    if (!family || disposed) return;
    render();
    deps.restoreSel?.();
    deps.apply(family);
  };
  const getMore = async (): Promise<void> => {
    deps.captureSel?.();
    picked(await requestRemoteFont({ container: wrap }));
  };
  // The file input opens in this click, so the browser allows it.
  const loadFile = async (): Promise<void> => {
    deps.captureSel?.();
    picked(await requestFontFile({ container: wrap, notify }));
  };

  // Keyboard: arrows, Home/End move between items, Escape closes back to the button.
  const items = (): HTMLButtonElement[] => Array.from(menu.querySelectorAll<HTMLButtonElement>(".sheetedit-font-item"));
  menu.addEventListener("keydown", (e) => {
    const list = items();
    const i = list.indexOf(document.activeElement as HTMLButtonElement);
    let next = -1;
    if (e.key === "ArrowDown") next = i < 0 ? 0 : (i + 1) % list.length;
    else if (e.key === "ArrowUp") next = i <= 0 ? list.length - 1 : i - 1;
    else if (e.key === "Home") next = 0;
    else if (e.key === "End") next = list.length - 1;
    else if (e.key === "Escape" || e.key === "Tab") {
      if (e.key === "Escape") e.preventDefault();
      done();
      return;
    } else return;
    e.preventDefault();
    list[next]?.focus();
  });
  // Rebuilt before the toolbar's own click handler shows it, so it opens current.
  button.addEventListener("click", () => { if (menu.hidden) render(); });
  deps.bindMenu(button, menu);
  // Opened from the keyboard (click detail 0): focus the checked font, else the first item.
  button.addEventListener("click", (e) => {
    button.setAttribute("aria-expanded", menu.hidden ? "false" : "true");
    if (menu.hidden || e.detail !== 0) return;
    const list = items();
    const target = list.find((b) => b.getAttribute("aria-checked") === "true") ?? list[0];
    target?.focus();
    target?.scrollIntoView?.({ block: "nearest" });
  });
  button.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown" && menu.hidden) {
      e.preventDefault();
      button.click();
    }
  });
  const hideObserver = new MutationObserver(() => button.setAttribute("aria-expanded", menu.hidden ? "false" : "true"));
  hideObserver.observe(menu, { attributes: true, attributeFilter: ["hidden"] });

  const off = onFontsChange(() => { if (!disposed && !menu.hidden) render(); });
  render();
  restoreFonts({ local: localOn, remote, docFonts });

  wrap.append(menu);
  return {
    menu,
    fonts: () => all,
    setDocFonts(names) {
      if (names.length === docFonts.length && names.every((n, i) => n === docFonts[i])) return;
      docFonts = [...names];
      unregister();
      unregister = addDocumentFonts(docFonts);
    },
    refresh: render,
    teardown() {
      disposed = true;
      off();
      unregister();
      hideObserver.disconnect();
      menu.remove();
    },
  };
}
