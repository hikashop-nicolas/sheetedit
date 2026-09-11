import { t } from "../i18n";
import type { SheetShape } from "../model";

// The bar that appears beside a selected shape. It replaces a properties dialog reached by
// double-clicking the shape: a dialog nobody knows is there is a dialog nobody opens, and every
// property then needs a home in two places that drift apart.
//
// Only properties the writers persist correctly are offered. Anything the editor can show but not
// save would be a control that silently does nothing the next time the file is opened.

export interface ShapeBarDeps {
  wrap: HTMLElement;
  /** Viewport rect of the grid area: the bar stays inside it. */
  bounds: () => DOMRect;
  /** Viewport rect of the selected shape, or null when nothing is selected. */
  shapeRect: () => DOMRect | null;
  /** The shape the bar is acting on, or null. */
  shape: () => SheetShape | null;
  /** A property changed: repaint the shape and mark the workbook. */
  onChange: (paintChanged: boolean) => void;
  /** Move the shape to the front / back of the drawing. Absent when the format cannot. */
  reorder?: (to: "front" | "back") => void;
  /** Open the shape gallery to swap the selected shape's geometry, anchored at `btn`. */
  swapGeometry: (btn: HTMLElement) => void;
  /** Edit the shape's text, anchored at `btn`. */
  editText: (btn: HTMLElement) => void;
}

/** Lines and connectors: no inside to fill, no text, but they do carry arrowheads. */
const isLine = (sh: SheetShape): boolean =>
  sh.geom === "line" || sh.geom === "elbow" || sh.geom === "curve" || sh.geom === "arc";

export function setupShapeBar(deps: ShapeBarDeps): { refresh(): void; teardown(): void } {
  const bar = document.createElement("div");
  bar.className = "sheetedit-floatbar sheetedit-shapebar";
  bar.hidden = true;
  bar.setAttribute("role", "toolbar");
  bar.setAttribute("aria-label", t("shapeEdit"));

  const groups: { el: HTMLElement; show: (sh: SheetShape) => boolean }[] = [];
  /** Apply a change to the selected shape and tell the host whether the paint moved. */
  const set = (fn: (sh: SheetShape) => void, paint = true): void => {
    const sh = deps.shape();
    if (!sh) return;
    fn(sh);
    sh.dirty = true;
    if (paint) sh.styleDirty = true;
    deps.onChange(paint);
  };

  const color = (title: string, get: (sh: SheetShape) => string, apply: (sh: SheetShape, v: string) => void): HTMLInputElement => {
    const c = document.createElement("input");
    c.type = "color";
    c.className = "sheetedit-color";
    c.title = title;
    c.setAttribute("aria-label", title);
    c.addEventListener("input", () => set((sh) => apply(sh, c.value)));
    (c as HTMLInputElement & { sync?: (sh: SheetShape) => void }).sync = (sh) => { c.value = get(sh); };
    return c;
  };

  /** A button that opens a list of choices, styled like the toolbar's own menus. */
  const menu = (label: string, title: string, items: [string, () => void][]): HTMLElement => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "sheetedit-btn";
    btn.textContent = label;
    btn.title = title;
    btn.setAttribute("aria-label", title);
    const list = document.createElement("div");
    list.className = "sheetedit-tb-groupmenu sheetedit-tb-listmenu";
    list.hidden = true;
    list.setAttribute("role", "menu");
    for (const [text, run] of items) {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "sheetedit-btn sheetedit-more-item";
      item.setAttribute("role", "menuitem");
      item.textContent = text;
      item.addEventListener("mousedown", (e) => e.preventDefault());
      item.addEventListener("click", () => { list.hidden = true; run(); });
      list.appendChild(item);
    }
    btn.addEventListener("mousedown", (e) => e.preventDefault());
    btn.addEventListener("click", () => {
      const open = list.hidden;
      for (const m of bar.querySelectorAll<HTMLElement>(".sheetedit-tb-groupmenu")) m.hidden = true;
      if (open) {
        const r = btn.getBoundingClientRect();
        list.style.left = `${r.left}px`;
        list.style.top = `${r.bottom + 2}px`;
      }
      list.hidden = !open;
    });
    deps.wrap.appendChild(list);
    return btn;
  };

  const sep = (): HTMLElement => {
    const d = document.createElement("div");
    d.className = "sheetedit-floatbar-sep";
    return d;
  };

  const group = (els: HTMLElement[], show: (sh: SheetShape) => boolean): void => {
    const wrapEl = document.createElement("span");
    wrapEl.className = "sheetedit-shapebar-group";
    for (const el of els) wrapEl.appendChild(el);
    bar.appendChild(wrapEl);
    groups.push({ el: wrapEl, show });
  };

  // --- fill and outline -----------------------------------------------------
  const fill = color(t("shapeFill"), (sh) => sh.fill ?? "#4c8bf5", (sh, v) => { sh.fill = v; sh.fillGradient = undefined; });
  const noFill = document.createElement("button");
  noFill.type = "button";
  noFill.className = "sheetedit-btn";
  noFill.textContent = "∅";
  noFill.title = t("shapeNoFill");
  noFill.setAttribute("aria-label", t("shapeNoFill"));
  noFill.addEventListener("mousedown", (e) => e.preventDefault());
  noFill.addEventListener("click", () => set((sh) => { sh.fill = undefined; sh.fillGradient = undefined; }));
  const opacity = menu("◑", t("shapeOpacity"), ([100, 75, 50, 25] as const).map((pct) =>
    [`${pct}%`, () => set((sh) => { sh.fillOpacity = pct === 100 ? undefined : pct / 100; })] as [string, () => void]));
  group([fill, noFill, opacity], (sh) => !isLine(sh));

  const stroke = color(t("shapeOutline"), (sh) => sh.stroke ?? "#1f3a5f", (sh, v) => { sh.stroke = v; });
  const width = menu("▬", t("shapeOutlineWidth"), [1, 2, 3, 4, 6].map((px) =>
    [`${px} px`, () => set((sh) => { sh.strokeWidth = px; })] as [string, () => void]));
  const dash = menu("┄", t("shapeDash"), ([
    [t("shapeDashSolid"), undefined],
    [t("shapeDashDash"), "dash"],
    [t("shapeDashDot"), "sysDot"],
  ] as const).map(([label, v]) => [label, () => set((sh) => { sh.dash = v; })] as [string, () => void]));
  group([sep(), stroke, width, dash], () => true);

  // --- arrowheads, for the shapes that have ends ---------------------------
  const END_KINDS = [
    [t("shapeEndNone"), undefined],
    [t("shapeEndArrow"), "triangle"],
    [t("shapeEndOval"), "oval"],
    [t("shapeEndDiamond"), "diamond"],
  ] as const;
  // A line has two ends and either can carry a head; offering only the far one meant the near
  // one could be read from a file and never changed.
  const headEnd = menu("←", t("shapeEndStart"), END_KINDS.map(([label, v]) =>
    [label, () => set((sh) => { sh.headEnd = v; })] as [string, () => void]));
  const tailEnd = menu("→", t("shapeEndFinish"), END_KINDS.map(([label, v]) =>
    [label, () => set((sh) => { sh.tailEnd = v; })] as [string, () => void]));
  group([sep(), headEnd, tailEnd], (sh) => isLine(sh));

  // --- text ----------------------------------------------------------------
  const textEdit = document.createElement("button");
  textEdit.type = "button";
  textEdit.className = "sheetedit-btn";
  textEdit.textContent = "T";
  textEdit.title = t("shapeTextEdit");
  textEdit.setAttribute("aria-label", t("shapeTextEdit"));
  textEdit.addEventListener("mousedown", (e) => e.preventDefault());
  textEdit.addEventListener("click", () => deps.editText(textEdit));
  const textColour = color(t("shapeTextColour"), (sh) => sh.textColor ?? "#000000", (sh, v) => { sh.textColor = v; });
  const textSize = menu("A", t("shapeTextSize"), [9, 11, 14, 18, 24, 40].map((pt) =>
    [`${pt}`, () => set((sh) => { sh.textSize = pt; })] as [string, () => void]));
  const textAlign = menu("≡", t("shapeTextAlign"), ([
    ["left", "center", "right"] as const,
  ][0]).map((a) => [a, () => set((sh) => { sh.textAlign = a; })] as [string, () => void]));
  group([sep(), textEdit, textColour, textSize, textAlign], (sh) => !isLine(sh));

  // --- the shape itself: which geometry, and which way round -----------------
  const swap = document.createElement("button");
  swap.type = "button";
  swap.className = "sheetedit-btn";
  swap.textContent = "◇";
  swap.title = t("shapeSwap");
  swap.setAttribute("aria-label", t("shapeSwap"));
  swap.addEventListener("mousedown", (e) => e.preventDefault());
  swap.addEventListener("click", () => deps.swapGeometry(swap));
  const flip = (label: string, title: string, apply: (sh: SheetShape) => void): HTMLElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sheetedit-btn";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("mousedown", (e) => e.preventDefault());
    // A mirror is geometry, not paint: recording it must not rewrite a themed fill.
    b.addEventListener("click", () => set((sh) => { apply(sh); sh.xfrmDirty = true; }, false));
    return b;
  };
  group([
    sep(), swap,
    flip("⇄", t("shapeFlipH"), (sh) => { sh.flipH = sh.flipH ? undefined : true; }),
    flip("⇅", t("shapeFlipV"), (sh) => { sh.flipV = sh.flipV ? undefined : true; }),
  ], () => true);

  // --- stacking -------------------------------------------------------------
  if (deps.reorder) {
    const front = document.createElement("button");
    front.type = "button";
    front.className = "sheetedit-btn";
    front.textContent = "⤒";
    front.title = t("shapeFront");
    front.setAttribute("aria-label", t("shapeFront"));
    front.addEventListener("mousedown", (e) => e.preventDefault());
    front.addEventListener("click", () => deps.reorder!("front"));
    const back = document.createElement("button");
    back.type = "button";
    back.className = "sheetedit-btn";
    back.textContent = "⤓";
    back.title = t("shapeBack");
    back.setAttribute("aria-label", t("shapeBack"));
    back.addEventListener("mousedown", (e) => e.preventDefault());
    back.addEventListener("click", () => deps.reorder!("back"));
    // A grouped shape has no anchor of its own: the group is what stacks.
    group([sep(), front, back], (sh) => !sh.within);
  }

  deps.wrap.appendChild(bar);

  const refresh = (): void => {
    const sh = deps.shape();
    const rect = sh ? deps.shapeRect() : null;
    if (!sh || !rect) {
      bar.hidden = true;
      for (const m of bar.querySelectorAll<HTMLElement>(".sheetedit-tb-groupmenu")) m.hidden = true;
      for (const m of deps.wrap.querySelectorAll<HTMLElement>(".sheetedit-shapebar ~ .sheetedit-tb-groupmenu")) m.hidden = true;
      return;
    }
    for (const g of groups) g.el.hidden = !g.show(sh);
    for (const el of bar.querySelectorAll<HTMLInputElement & { sync?: (s: SheetShape) => void }>("input.sheetedit-color")) el.sync?.(sh);
    bar.hidden = false;
    // Above the shape where there is room, below it otherwise, and never outside the grid.
    const grid = deps.bounds();
    const bw = bar.offsetWidth || 260;
    const bh = bar.offsetHeight || 32;
    let left = rect.left + rect.width / 2 - bw / 2;
    left = Math.max(grid.left + 4, Math.min(left, grid.right - bw - 4));
    let top = rect.top - bh - 10;
    if (top < grid.top + 4) top = rect.bottom + 10;
    top = Math.min(top, grid.bottom - bh - 4);
    bar.style.left = `${left}px`;
    bar.style.top = `${top}px`;
  };

  return {
    refresh,
    teardown() {
      for (const m of deps.wrap.querySelectorAll(".sheetedit-tb-groupmenu")) m.remove();
      bar.remove();
    },
  };
}
