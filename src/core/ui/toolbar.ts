import { t } from "../i18n";
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
  undo: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3 3 6l3 3"/><path d="M3 6h6a4 4 0 0 1 0 8H7"/></svg>`,
  redo: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="m10 3 3 3-3 3"/><path d="M13 6H7a4 4 0 0 0 0 8h2"/></svg>`,
  left: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M2 8h7M2 12h10"/></svg>`,
  center: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M4.5 8h7M3 12h10"/></svg>`,
  right: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M7 8h7M4 12h10"/></svg>`,
  borders: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4"><rect x="2" y="2" width="12" height="12"/><path d="M8 2v12M2 8h12"/></svg>`,
  merge: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><rect x="1.5" y="4" width="13" height="8"/><path d="M5 6 7.5 8 5 10M11 6 8.5 8 11 10"/></svg>`,
};

export interface ToolbarHandle {
  relayout(): void;
  teardown(): void;
}

export function buildToolbar(ctx: {
  toolbar: HTMLElement;
  wrap: HTMLElement;
  /** xlsx/ods: the style model is known, so the styling cluster is shown. */
  styled: boolean;
  /** csv mode: shown as a "Convert to XLSX" button; null hides it. */
  convert?: (() => void) | null;
  onUndo(): void;
  onRedo(): void;
  addRows(): void;
  addCols(): void;
  applyStyle(change: StyleChange): void;
  curStyle(): CellStyle | undefined;
  openBorderPopover(anchor: HTMLElement): void;
  toggleMerge(): void;
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
  );
  if (ctx.convert) toolbar.append(sep(), tbBtn(t("convertXlsx"), t("convertXlsxTitle"), ctx.convert));
  if (!ctx.styled) return { relayout: () => undefined, teardown: () => undefined };

  const bold = tbBtn("B", t("bold"), () => ctx.applyStyle({ bold: !ctx.curStyle()?.bold }));
  bold.style.fontWeight = "700";
  const italic = tbBtn("I", t("italic"), () => ctx.applyStyle({ italic: !ctx.curStyle()?.italic }));
  italic.style.fontStyle = "italic";
  const borderBtn = tbIcon(ICON.borders, t("borders"), () => ctx.openBorderPopover(borderBtn));
  const styleControls: HTMLElement[] = [
    bold,
    italic,
    colorInput(t("textColour"), "#000000", (v) => ctx.applyStyle({ color: v })),
    colorInput(t("fillColour"), "#ffff00", (v) => ctx.applyStyle({ bg: v })),
    sep(),
    tbIcon(ICON.left, t("alignLeft"), () => ctx.applyStyle({ align: "left" })),
    tbIcon(ICON.center, t("alignCentre"), () => ctx.applyStyle({ align: "center" })),
    tbIcon(ICON.right, t("alignRight"), () => ctx.applyStyle({ align: "right" })),
    sep(),
    borderBtn,
    tbIcon(ICON.merge, t("merge"), ctx.toggleMerge),
  ];

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

  const fits = () => toolbar.scrollWidth <= toolbar.clientWidth + 1;
  const relayout = () => {
    expand();
    if (!fits()) collapse();
  };
  relayout();
  requestAnimationFrame(relayout);
  const observer = new ResizeObserver(relayout);
  observer.observe(toolbar);
  return {
    relayout,
    teardown() {
      observer.disconnect();
      document.removeEventListener("click", closeMenu);
      menu.remove();
    },
  };
}
