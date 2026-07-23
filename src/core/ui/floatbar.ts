import { t } from "../i18n";
import { ICON, tbIcon } from "./toolbar";
import type { CellStyle, StyleChange } from "../model";

// Floating style bar (desktop only): a compact bold/italic/colour/align bar that
// appears near the selection when the mouse approaches it, mirroring richdoc's
// float bar, so common styling doesn't require the top toolbar. Hidden on
// coarse-pointer (touch) devices.

export function setupFloatBar(deps: {
  wrap: HTMLElement;
  /** Viewport rect of the grid area; the bar only reacts and shows inside it. */
  bounds: () => DOMRect;
  /** Viewport rect of the current selection, or null when nothing should show. */
  selRect: () => DOMRect | null;
  curStyle: () => CellStyle | undefined;
  applyStyle: (change: StyleChange) => void;
  /** Sparkline actions: shown only when the focused single cell hosts a sparkline. */
  spark?: { has: () => boolean; edit: () => void; remove: () => void };
}): { teardown(): void } {
  const coarse = typeof window.matchMedia === "function" && window.matchMedia("(hover: none), (pointer: coarse)").matches;
  if (coarse) return { teardown: () => undefined };

  const bar = document.createElement("div");
  bar.className = "sheetedit-floatbar";
  bar.hidden = true;
  let hovered = false;
  let hideTimer = 0;

  const btn = (label: string, title: string, apply: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sheetedit-btn";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the selection/focus
    b.addEventListener("click", apply);
    return b;
  };
  const color = (title: string, def: string, apply: (v: string) => void): HTMLInputElement => {
    const c = document.createElement("input");
    c.type = "color";
    c.value = def;
    c.title = title;
    c.setAttribute("aria-label", title);
    c.className = "sheetedit-color";
    c.addEventListener("mousedown", () => {
      hovered = true; // keep the bar open while the native picker is up
      window.clearTimeout(hideTimer);
    });
    c.addEventListener("input", () => apply(c.value));
    c.addEventListener("change", () => {
      hovered = false;
    });
    return c;
  };
  const bold = btn("B", t("bold"), () => deps.applyStyle({ bold: !deps.curStyle()?.bold }));
  bold.style.fontWeight = "700";
  const italic = btn("I", t("italic"), () => deps.applyStyle({ italic: !deps.curStyle()?.italic }));
  italic.style.fontStyle = "italic";
  const alignBtn = (svg: string, title: string, align: "left" | "center" | "right") => {
    const b = tbIcon(svg, title, () => deps.applyStyle({ align }));
    b.addEventListener("mousedown", (e) => e.preventDefault());
    return b;
  };
  const underline = btn("U", t("underline"), () => deps.applyStyle({ underline: !deps.curStyle()?.underline }));
  underline.style.textDecoration = "underline";
  const strike = btn("S", t("strikethrough"), () => deps.applyStyle({ strike: !deps.curStyle()?.strike }));
  strike.style.textDecoration = "line-through";
  bar.append(
    bold,
    italic,
    underline,
    strike,
    color(t("textColour"), "#000000", (v) => deps.applyStyle({ color: v })),
    color(t("fillColour"), "#ffff00", (v) => deps.applyStyle({ bg: v })),
    alignBtn(ICON.left, t("alignLeft"), "left"),
    alignBtn(ICON.center, t("alignCentre"), "center"),
    alignBtn(ICON.right, t("alignRight"), "right"),
  );

  // Sparkline actions live behind a divider and only appear when the focused cell hosts one.
  let sparkGroup: HTMLElement[] = [];
  if (deps.spark) {
    const sp = deps.spark;
    const iconBtn = (svg: string, title: string, run: () => void): HTMLButtonElement => {
      const b = tbIcon(svg, title, run);
      b.addEventListener("mousedown", (e) => e.preventDefault());
      return b;
    };
    const sep = document.createElement("span");
    sep.className = "sheetedit-floatbar-sep";
    const SPARK_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 11l3-4 3 2 3-5 3 3"/></svg>`;
    const TRASH_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 4h10M6.5 4V2.5h3V4M5 4l.6 9a1 1 0 0 0 1 .9h2.8a1 1 0 0 0 1-.9L11 4"/></svg>`;
    const edit = iconBtn(SPARK_ICON, t("sparkEdit"), () => sp.edit());
    const del = iconBtn(TRASH_ICON, t("sparkDelete"), () => sp.remove());
    sparkGroup = [sep, edit, del];
    bar.append(sep, edit, del);
  }
  deps.wrap.appendChild(bar);

  const hide = () => {
    bar.hidden = true;
  };
  const scheduleHide = () => {
    window.clearTimeout(hideTimer);
    hideTimer = window.setTimeout(() => {
      if (!hovered) hide();
    }, 350);
  };
  const showAt = (rect: DOMRect) => {
    const showSpark = !!deps.spark?.has();
    for (const el of sparkGroup) el.hidden = !showSpark;
    bar.hidden = false;
    const grid = deps.bounds();
    const bw = bar.offsetWidth || 200;
    const bh = bar.offsetHeight || 32;
    let left = rect.left + rect.width / 2 - bw / 2;
    left = Math.max(grid.left + 4, Math.min(left, grid.right - bw - 4));
    // Never leave the grid area (the toolbar and formula bar live above it).
    let top = rect.top - bh - 8;
    if (top < grid.top + 4) top = rect.bottom + 8;
    top = Math.min(top, grid.bottom - bh - 4);
    bar.style.left = `${left}px`;
    bar.style.top = `${top}px`;
  };
  const onMove = (e: MouseEvent) => {
    if (hovered) return;
    const grid = deps.bounds();
    if (e.clientX < grid.left || e.clientX > grid.right || e.clientY < grid.top || e.clientY > grid.bottom) {
      scheduleHide();
      return;
    }
    const rect = deps.selRect();
    if (!rect) {
      scheduleHide();
      return;
    }
    const dx = Math.max(rect.left - e.clientX, 0, e.clientX - rect.right);
    const dy = Math.max(rect.top - e.clientY, 0, e.clientY - rect.bottom);
    if (dx < 110 && dy < 90) {
      window.clearTimeout(hideTimer);
      showAt(rect);
    } else {
      scheduleHide();
    }
  };
  document.addEventListener("mousemove", onMove);
  bar.addEventListener("mouseenter", () => {
    hovered = true;
  });
  bar.addEventListener("mouseleave", () => {
    hovered = false;
    scheduleHide();
  });

  return {
    teardown() {
      document.removeEventListener("mousemove", onMove);
      window.clearTimeout(hideTimer);
      bar.remove();
    },
  };
}
