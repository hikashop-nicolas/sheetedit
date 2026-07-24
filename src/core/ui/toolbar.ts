import { localeCode, t } from "../i18n";
import { numFmtPresets } from "../dates";
import type { CellStyle, StyleChange } from "../model";

// The style toolbar, split out of the editor. The styling controls form one
// collapsible cluster (richdoc's toolbar pattern): inline while there is room,
// folded into a single "Aa" dropdown when the toolbar runs out of width.

export const tbBtn = (label: string, title: string, onClick: () => void): HTMLButtonElement => {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "sheetedit-btn";
  b.textContent = label;
  b.title = title;
  b.setAttribute("aria-label", title);
  b.addEventListener("click", onClick);
  return b;
};
export const tbIcon = (svg: string, title: string, onClick: () => void): HTMLButtonElement => {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "sheetedit-btn";
  b.innerHTML = svg;
  b.title = title;
  b.setAttribute("aria-label", title);
  b.addEventListener("click", onClick);
  return b;
};

export const ICON = {
  find: `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" aria-hidden="true"><circle cx="7" cy="7" r="4.2"/><path d="m10.2 10.2 3.3 3.3"/></svg>`,
  undo: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3 3 6l3 3"/><path d="M3 6h6a4 4 0 0 1 0 8H7"/></svg>`,
  redo: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="m10 3 3 3-3 3"/><path d="M13 6H7a4 4 0 0 0 0 8h2"/></svg>`,
  left: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M2 8h7M2 12h10"/></svg>`,
  center: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M4.5 8h7M3 12h10"/></svg>`,
  right: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M7 8h7M4 12h10"/></svg>`,
  borders: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4"><rect x="2" y="2" width="12" height="12"/><path d="M8 2v12M2 8h12"/></svg>`,
  merge: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><rect x="1.5" y="4" width="13" height="8"/><path d="M5 6 7.5 8 5 10M11 6 8.5 8 11 10"/></svg>`,
  valignTop: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 2.5h12M8 6v8M5.5 8.5 8 6l2.5 2.5"/></svg>`,
  valignMiddle: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 8h12M8 1.5v4M6 3.5 8 5.5l2-2M8 14.5v-4M6 12.5l2 2 2-2"/></svg>`,
  valignBottom: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 13.5h12M8 2v8M5.5 7.5 8 10l2.5-2.5"/></svg>`,
  wrap: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M2 4h12M2 8h8.5a2.25 2.25 0 0 1 0 4.5H8"/><path d="m9.5 10.5-2 2 2 2"/><path d="M2 12h3"/></svg>`,
};

export interface ToolbarHandle {
  relayout(): void;
  teardown(): void;
  /** Place trailing (authoring) controls the editor adds; they fold into a "⋯" menu when tight. */
  setTrailing(elements: HTMLElement[]): void;
  /** Reflect the active cell's style as pressed states on the toggle buttons (bold, align, ...). */
  syncActive(): void;
}

export function buildToolbar(ctx: {
  toolbar: HTMLElement;
  wrap: HTMLElement;
  /** xlsx/ods: the style model is known, so the styling cluster is shown. */
  styled: boolean;
  /** csv mode: shown as a "Convert to XLSX" button; null hides it. */
  convert?: (() => void) | null;
  findReplace(): void;
  onUndo(): void;
  onRedo(): void;
  addRows(): void;
  addCols(): void;
  applyStyle(change: StyleChange): void;
  applyNumFmt(fmt: string | number | undefined, currency?: string): void;
  curStyle(): CellStyle | undefined;
  openBorderPopover(anchor: HTMLElement): void;
  toggleMerge(): void;
  /** Open the furigana (phonetic reading) editor for the active cell, anchored at `btn`. */
  editFurigana(btn: HTMLElement): void;
}): ToolbarHandle {
  const { toolbar, wrap } = ctx;
  toolbar.innerHTML = "";
  const sep = () => {
    const d = document.createElement("div");
    d.className = "sheetedit-tb-sep";
    return d;
  };
  const colorInput = (title: string, def: string, apply: (v: string) => void) => {
    const i = document.createElement("input");
    i.type = "color";
    i.title = title;
    i.setAttribute("aria-label", title);
    i.className = "sheetedit-color";
    i.value = def;
    i.addEventListener("change", () => apply(i.value));
    return i;
  };
  const undoBtn = tbIcon(ICON.undo, t("undo"), ctx.onUndo);
  const redoBtn = tbIcon(ICON.redo, t("redo"), ctx.onRedo);
  // Keep the grid's focus (and pending selection) when clicking undo/redo.
  undoBtn.addEventListener("mousedown", (e) => e.preventDefault());
  redoBtn.addEventListener("mousedown", (e) => e.preventDefault());
  toolbar.append(
    undoBtn,
    redoBtn,
    sep(),
    tbBtn(t("addRow"), t("addRows"), ctx.addRows),
    tbBtn(t("addCol"), t("addCols"), ctx.addCols),
    tbIcon(ICON.find, t("findReplace"), ctx.findReplace),
  );
  if (ctx.convert) toolbar.append(sep(), tbBtn(t("convertXlsx"), t("convertXlsxTitle"), ctx.convert));
  if (!ctx.styled) return { relayout: () => undefined, teardown: () => undefined, setTrailing: () => undefined, syncActive: () => undefined };

  const bold = tbBtn("B", t("bold"), () => ctx.applyStyle({ bold: !ctx.curStyle()?.bold }));
  bold.style.fontWeight = "700";
  const italic = tbBtn("I", t("italic"), () => ctx.applyStyle({ italic: !ctx.curStyle()?.italic }));
  italic.style.fontStyle = "italic";
  const underline = tbBtn("U", t("underline"), () => ctx.applyStyle({ underline: !ctx.curStyle()?.underline }));
  underline.style.textDecoration = "underline";
  const strike = tbBtn("S", t("strikethrough"), () => ctx.applyStyle({ strike: !ctx.curStyle()?.strike }));
  strike.style.textDecoration = "line-through";
  const borderBtn = tbIcon(ICON.borders, t("borders"), () => ctx.openBorderPopover(borderBtn));
  // Furigana editor: a compact "ふ" button (the feature is inherently Japanese; the title explains).
  const furiBtn = tbBtn("ふ", t("furiganaTitle"), () => ctx.editFurigana(furiBtn));

  // Font family / size: stateless menus (like the toolbar's other controls, they
  // read nothing back); the placeholder row re-selects itself after each apply.
  const picker = (title: string, placeholder: string, options: [string, string][], apply: (v: string) => void) => {
    const s = document.createElement("select");
    s.className = "sheetedit-tb-select";
    s.title = title;
    s.setAttribute("aria-label", title);
    const ph = document.createElement("option");
    ph.value = "";
    ph.textContent = placeholder;
    ph.disabled = true;
    ph.selected = true;
    s.appendChild(ph);
    for (const [v, label] of options) {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = label;
      s.appendChild(o);
    }
    s.addEventListener("change", () => {
      if (s.value) apply(s.value);
      s.selectedIndex = 0;
    });
    return s;
  };
  const FAMILIES = ["Arial", "Calibri", "Courier New", "Georgia", "Helvetica", "Times New Roman", "Verdana"];
  const famSel = picker(t("fontFamily"), "Aa", FAMILIES.map((f) => [f, f]), (v) => ctx.applyStyle({ fontFamily: v }));
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
  const sizeSel = picker(t("fontSize"), "pt", SIZES.map((n) => [String(n), String(n)]), (v) => ctx.applyStyle({ fontSize: Number(v) }));

  // Number-format picker: a button opening a preset menu (General, number,
  // percent, currency, date, time), shapes localized via numFmtPresets.
  const fmtBtn = tbBtn("123 ▾", t("numFormat"), () => {
    const r = fmtBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    fmtMenu.style.top = `${r.bottom - wr.top + 2}px`;
    fmtMenu.style.left = `${r.left - wr.left}px`;
    fmtMenu.hidden = !fmtMenu.hidden;
  });
  fmtBtn.addEventListener("mousedown", (e) => e.preventDefault());
  const fmtMenu = document.createElement("div");
  fmtMenu.className = "sheetedit-tb-groupmenu sheetedit-fmtmenu";
  fmtMenu.hidden = true;
  for (const preset of numFmtPresets(localeCode())) {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "sheetedit-btn";
    item.textContent = t(preset.key);
    item.addEventListener("mousedown", (e) => e.preventDefault());
    item.addEventListener("click", () => {
      fmtMenu.hidden = true;
      ctx.applyNumFmt(preset.fmt, preset.currency);
    });
    fmtMenu.appendChild(item);
  }
  wrap.append(fmtMenu);
  const closeFmtMenu = (e: MouseEvent) => {
    if (!fmtMenu.hidden && !fmtMenu.contains(e.target as Node) && !fmtBtn.contains(e.target as Node)) fmtMenu.hidden = true;
  };
  document.addEventListener("click", closeFmtMenu);

  const alignL = tbIcon(ICON.left, t("alignLeft"), () => ctx.applyStyle({ align: "left" }));
  const alignC = tbIcon(ICON.center, t("alignCentre"), () => ctx.applyStyle({ align: "center" }));
  const alignR = tbIcon(ICON.right, t("alignRight"), () => ctx.applyStyle({ align: "right" }));
  const valignT = tbIcon(ICON.valignTop, t("valignTop"), () => ctx.applyStyle({ valign: "top" }));
  const valignM = tbIcon(ICON.valignMiddle, t("valignMiddle"), () => ctx.applyStyle({ valign: "middle" }));
  const valignB = tbIcon(ICON.valignBottom, t("valignBottom"), () => ctx.applyStyle({ valign: "bottom" }));
  const wrapBtn = tbIcon(ICON.wrap, t("wrapText"), () => ctx.applyStyle({ wrap: !ctx.curStyle()?.wrap }));

  const styleControls: HTMLElement[] = [
    fmtBtn,
    famSel,
    sizeSel,
    sep(),
    bold,
    italic,
    underline,
    strike,
    colorInput(t("textColour"), "#000000", (v) => ctx.applyStyle({ color: v })),
    colorInput(t("fillColour"), "#ffff00", (v) => ctx.applyStyle({ bg: v })),
    sep(),
    alignL, alignC, alignR,
    valignT, valignM, valignB,
    wrapBtn,
    sep(),
    borderBtn,
    tbIcon(ICON.merge, t("merge"), ctx.toggleMerge),
    furiBtn,
  ];

  // Reflect the active cell's style as pressed toggle buttons. A stateless button (fmt/font pickers)
  // reads nothing back; only the on/off + mutually-exclusive controls surface a current value.
  const setPressed = (btn: HTMLElement, on: boolean) => { btn.classList.toggle("is-active", on); btn.setAttribute("aria-pressed", on ? "true" : "false"); };
  const syncActive = () => {
    const st = ctx.curStyle();
    setPressed(bold, !!st?.bold);
    setPressed(italic, !!st?.italic);
    setPressed(underline, !!st?.underline);
    setPressed(strike, !!st?.strike);
    setPressed(alignL, st?.align === "left");
    setPressed(alignC, st?.align === "center");
    setPressed(alignR, st?.align === "right");
    setPressed(valignT, st?.valign === "top");
    setPressed(valignM, st?.valign === "middle");
    setPressed(valignB, st?.valign === "bottom");
    setPressed(wrapBtn, !!st?.wrap);
  };

  // Collapsible cluster: inline when it fits, otherwise one "Aa" button + popover.
  const slot = document.createElement("span");
  slot.className = "sheetedit-tb-slot";
  const groupBtn = tbBtn("Aa ▾", t("formatting"), () => {
    const r = groupBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    menu.style.top = `${r.bottom - wr.top + 2}px`;
    menu.style.left = `${r.left - wr.left}px`;
    menu.hidden = !menu.hidden;
  });
  groupBtn.addEventListener("mousedown", (e) => e.preventDefault());
  const menu = document.createElement("div");
  menu.className = "sheetedit-tb-groupmenu";
  menu.hidden = true;
  let collapsed = false;
  const expand = () => {
    if (!collapsed) return;
    slot.replaceChildren(...styleControls);
    menu.hidden = true;
    collapsed = false;
  };
  const collapse = () => {
    if (collapsed) return;
    menu.replaceChildren(...styleControls);
    slot.replaceChildren(groupBtn);
    collapsed = true;
  };
  slot.replaceChildren(...styleControls);
  toolbar.append(sep(), slot);
  wrap.append(menu);
  const closeMenu = (e: MouseEvent) => {
    if (!menu.hidden && !menu.contains(e.target as Node) && !groupBtn.contains(e.target as Node)) menu.hidden = true;
  };
  document.addEventListener("click", closeMenu);

  // Overflow "⋯" menu: trailing authoring controls (added by the editor) that don't fit fold into a
  // dropdown, each shown as an icon + label row that forwards to the original (hidden) button.
  let trailing: HTMLElement[] = [];
  const moreBtn = tbBtn("⋯", t("more"), () => {
    const r = moreBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    moreMenu.style.top = `${r.bottom - wr.top + 2}px`;
    moreMenu.style.left = `${Math.max(0, r.right - wr.left - 180)}px`;
    moreMenu.hidden = !moreMenu.hidden;
  });
  moreBtn.addEventListener("mousedown", (e) => e.preventDefault());
  moreBtn.style.display = "none";
  const moreMenu = document.createElement("div");
  moreMenu.className = "sheetedit-tb-groupmenu sheetedit-tb-moremenu";
  moreMenu.hidden = true;
  wrap.append(moreMenu);
  const moreItem = (btn: HTMLElement): HTMLElement => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "sheetedit-btn sheetedit-more-item";
    row.innerHTML = `${btn.innerHTML}<span>${btn.getAttribute("aria-label") ?? btn.title ?? ""}</span>`;
    row.addEventListener("mousedown", (e) => e.preventDefault());
    row.addEventListener("click", () => { moreMenu.hidden = true; (btn as HTMLElement).click(); });
    return row;
  };
  const closeMoreMenu = (e: MouseEvent) => { if (!moreMenu.hidden && !moreMenu.contains(e.target as Node) && !moreBtn.contains(e.target as Node)) moreMenu.hidden = true; };
  document.addEventListener("click", closeMoreMenu);
  const setTrailing = (els: HTMLElement[]) => {
    for (const el of trailing) el.remove();
    trailing = els;
    for (const el of trailing) toolbar.insertBefore(el, moreBtn);
    relayout();
  };
  toolbar.append(moreBtn);

  const fits = () => toolbar.scrollWidth <= toolbar.clientWidth + 1;
  const relayout = () => {
    // Reset to the fully-expanded state, then fold only as much as needed.
    expand();
    for (const el of trailing) el.style.display = "";
    moreMenu.replaceChildren();
    moreBtn.style.display = "none";
    if (fits()) return;
    // Fold the trailing authoring controls into the "⋯" menu first (from the end).
    moreBtn.style.display = "";
    for (let i = trailing.length - 1; i >= 0; i--) {
      if (fits()) break;
      trailing[i]!.style.display = "none";
      moreMenu.prepend(moreItem(trailing[i]!));
    }
    if (!moreMenu.children.length) moreBtn.style.display = "none";
    // Only when every authoring control has folded and it still overflows, collapse the style
    // cluster into its "Aa" menu as a last resort.
    if (!fits()) collapse();
  };
  relayout();
  requestAnimationFrame(relayout);
  const observer = new ResizeObserver(relayout);
  observer.observe(toolbar);
  return {
    relayout,
    setTrailing,
    syncActive,
    teardown() {
      observer.disconnect();
      document.removeEventListener("click", closeMenu);
      document.removeEventListener("click", closeFmtMenu);
      document.removeEventListener("click", closeMoreMenu);
      menu.remove();
      fmtMenu.remove();
      moreMenu.remove();
    },
  };
}
