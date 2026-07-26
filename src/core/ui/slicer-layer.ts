import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet, SheetSlicer, Workbook } from "../model";
import type { ChartGeom } from "./chart-overlay";

// The slicer overlay: a titled panel of item buttons floated over the grid, anchored to cells and
// glued while scrolling. Clicking an item filters the linked pivot(s): a plain click selects only
// that item, ctrl/cmd-click toggles it within the current selection, and the "clear filter" button
// selects everything again. The host recomputes the pivots and persists the new selection.

export interface SlicerLayerDeps {
  wrap: HTMLElement;
  /** The grid's viewports: the element plus whether it draws the header / row numbers. */
  panes: () => { el: HTMLElement; header: boolean; rowHeader: boolean }[];
  getSheet: () => Sheet | undefined;
  /** The workbook, for its user-defined slicer styles. */
  getWorkbook?: () => Workbook;
  geom: () => ChartGeom;
  /** Selection changed: the model is already updated and dirty set. */
  onChange?: (s: SheetSlicer) => void;
}


/** The accent colour of a built-in slicer style (SlicerStyleLight1..6 / Dark1..6 / Other1..2).
    Excel's built-ins differ mainly by accent, so the family index picks a theme-ish accent. */
export function styleAccent(style: string | undefined): string | undefined {
  if (!style) return undefined;
  const m = /^SlicerStyle(Light|Dark|Other)(\d+)$/i.exec(style.trim());
  if (!m) return undefined;
  const n = Math.max(1, Math.min(6, Number(m[2]) || 1));
  // Accent1..6 of the default Office theme; Light1 is the neutral grey one.
  const accents = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"];
  const family = m[1]!.toLowerCase();
  if (family === "light" && n === 1) return "#7f7f7f";
  const base = accents[(n - 1) % accents.length]!;
  // The Dark family uses a deeper tone of the same accent.
  return family === "dark" ? shade(base, -0.25) : base;
}
function shade(hex: string, amount: number): string {
  const c = hex.replace("#", "");
  const ch = (i: number): string => {
    const v = parseInt(c.slice(i, i + 2), 16);
    const out = amount < 0 ? v * (1 + amount) : v + (255 - v) * amount;
    return Math.max(0, Math.min(255, Math.round(out))).toString(16).padStart(2, "0");
  };
  return `#${ch(0)}${ch(2)}${ch(4)}`;
}

export function setupSlicerLayer(deps: SlicerLayerDeps): { refresh(): void; teardown(): void } {
  const hosts = setupOverlayHosts({
    wrap: deps.wrap,
    panes: deps.panes,
    geom: () => { const g = deps.geom(); return { rnW: g.rnW, headerH: g.headerH, yOfRow: g.yOfRow }; },
    className: "sheetedit-slicerlayer",
    innerClassName: "sheetedit-slicerlayer-inner",
  });

  const refresh = (): void => {
    hosts.clear();
    const sheet = deps.getSheet();
    const slicers = sheet?.slicers ?? [];
    hosts.setVisible(slicers.length > 0);
    if (!slicers.length) return;
    const g = deps.geom();
    for (const sl of slicers) {
      const a = sl.anchor;
      // A slicer with no drawing anchor still needs somewhere to live.
      const x = a ? g.xOfCol(a.fromCol) + a.fromColOff : 20;
      const y = a ? g.yOfRow(a.fromRow) + a.fromRowOff : 20;
      const w = a ? Math.max(120, g.xOfCol(a.toCol) + a.toColOff - x) : 160;
      const h = a ? Math.max(80, g.yOfRow(a.toRow) + a.toRowOff - y) : 180;
      const readonly = sl.kind === "olap";
      const box = document.createElement("div");
      box.className = "sheetedit-slicerbox" + (readonly ? " readonly" : "");
      Object.assign(box.style, { left: `${x}px`, top: `${y}px`, width: `${w}px`, height: `${h}px` });
      // A user-defined style carries real colours; the built-ins only differ by accent, so map the
      // family to a theme accent instead.
      const custom = sl.style ? deps.getWorkbook?.().slicerStyles?.get(sl.style) : undefined;
      const accent = custom?.selectedFill ?? styleAccent(sl.style);
      if (accent) box.style.setProperty("--se-slicer-accent", accent);
      if (custom?.selectedText) box.style.setProperty("--se-slicer-on-fg", custom.selectedText);
      if (custom?.unselectedFill) box.style.setProperty("--se-slicer-off-bg", custom.unselectedFill);
      if (custom?.unselectedText) box.style.setProperty("--se-slicer-off-fg", custom.unselectedText);
      if (custom?.unselectedFill || custom?.unselectedText) box.classList.add("styled");
      box.dataset.slicer = sl.name;

      const head = document.createElement("div");
      head.className = "sheetedit-slicer-head";
      const title = document.createElement("span");
      title.className = "sheetedit-slicer-title";
      title.textContent = sl.caption || sl.sourceName || sl.name;
      title.title = title.textContent;
      const clear = document.createElement("button");
      clear.type = "button";
      clear.className = "sheetedit-slicer-clear";
      clear.textContent = "⊗";
      clear.title = "Clear filter";
      clear.addEventListener("click", (e) => {
        e.stopPropagation();
        for (const it of sl.items) it.selected = true;
        sl.dirty = true;
        deps.onChange?.(sl);
      });
      head.append(title, clear);
      box.appendChild(head);

      const list = document.createElement("div");
      list.className = "sheetedit-slicer-items";
      list.style.gridTemplateColumns = `repeat(${Math.max(1, sl.columnCount ?? 1)}, minmax(0, 1fr))`;
      for (const item of sl.items) {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-slicer-item" + (item.selected ? " on" : "");
        b.textContent = item.label === "" ? "(blank)" : item.label;
        b.title = b.textContent;
        b.dataset.item = String(item.x);
        if (readonly) { b.title = `${b.textContent} (OLAP slicer: shown for reference, not filterable here)`; list.appendChild(b); continue; }
        b.addEventListener("click", (e) => {
          const additive = e.ctrlKey || e.metaKey;
          const allOn = sl.items.every((i) => i.selected);
          if (additive) item.selected = !item.selected;
          // A plain click on an item in a "no filter" slicer narrows to just it; clicking the only
          // selected item clears the filter again.
          else if (allOn) for (const i of sl.items) i.selected = i === item;
          else if (item.selected && sl.items.filter((i) => i.selected).length === 1) for (const i of sl.items) i.selected = true;
          else for (const i of sl.items) i.selected = i === item;
          if (!sl.items.some((i) => i.selected)) for (const i of sl.items) i.selected = true; // never empty
          sl.dirty = true;
          deps.onChange?.(sl);
        });
        list.appendChild(b);
      }
      box.appendChild(list);
      hosts.hostFor(sl.anchor?.fromRow ?? 1).appendChild(box);
    }
    hosts.layout();
  };

  return { refresh, teardown() { hosts.teardown(); } };
}
