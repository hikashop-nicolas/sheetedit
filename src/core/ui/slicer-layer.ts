import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet, SheetSlicer, Workbook } from "../model";
import type { ChartGeom } from "./chart-overlay";

// The slicer overlay: a titled panel of item buttons floated over the grid, anchored to cells and
// glued while scrolling. Clicking an item filters the linked pivot(s): a plain click selects only
// that item, ctrl/cmd-click toggles it within the current selection, and the "clear filter" button
// selects everything again. The host recomputes the pivots and persists the new selection.

export interface SlicerLayerDeps {
  wrap: HTMLElement;
  /** The grid's scroll containers, top-most first (two when a row split is on). */
  panes: () => HTMLElement[];
  getSheet: () => Sheet | undefined;
  /** The workbook, for its user-defined slicer styles. */
  getWorkbook?: () => Workbook;
  geom: () => ChartGeom;
  /** Selection changed: the model is already updated and dirty set. */
  onChange?: (s: SheetSlicer) => void;
}

const STYLE_ID = "sheetedit-slicer-style";
function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-slicerlayer { position:absolute; overflow:hidden; pointer-events:none; z-index:6; }
    .sheetedit-slicerlayer-inner { position:absolute; inset:0; }
    .sheetedit-slicerbox { position:absolute; pointer-events:auto; display:flex; flex-direction:column;
      background:var(--sheetedit-chrome,#fff); color:var(--sheetedit-text,#1c1f24);
      border:1px solid var(--sheetedit-border,#c8ccd2); border-radius:6px;
      box-shadow:0 2px 10px rgba(0,0,0,.18); font:12px system-ui,sans-serif; overflow:hidden; }
    .sheetedit-slicer-head { display:flex; align-items:center; gap:6px; padding:5px 7px; font-weight:600;
      border-bottom:1px solid var(--sheetedit-border,#c8ccd2); background:rgba(127,127,127,.08); }
    .sheetedit-slicer-title { flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    .sheetedit-slicer-clear { border:0; background:none; cursor:pointer; color:inherit; opacity:.65;
      font:inherit; padding:1px 4px; border-radius:4px; }
    .sheetedit-slicer-clear:hover { opacity:1; background:rgba(127,127,127,.18); }
    .sheetedit-slicer-items { flex:1; overflow:auto; padding:4px; display:grid; gap:3px; }
    .sheetedit-slicer-item { border:1px solid var(--sheetedit-border,#c8ccd2); border-radius:4px;
      background:var(--se-slicer-off-bg,transparent); color:var(--se-slicer-off-fg,inherit);
      font:inherit; padding:3px 6px; cursor:pointer;
      text-align:left; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; opacity:.45; }
    .sheetedit-slicer-item:hover { border-color:var(--sheetedit-accent,#4c8bf5); }
    .sheetedit-slicer-item.on { opacity:1; background:var(--se-slicer-accent,var(--sheetedit-accent,#4c8bf5));
      color:var(--se-slicer-on-fg,#fff);
      border-color:var(--se-slicer-accent,var(--sheetedit-accent,#4c8bf5)); }
    /* A custom style says exactly how an unselected item looks, so do not dim it on top. */
    .sheetedit-slicerbox.styled .sheetedit-slicer-item { opacity:1; }
    /* OLAP slicers have no source we can filter, so they show their items but do not react. */
    .sheetedit-slicerbox.readonly .sheetedit-slicer-item { cursor:default; }
    .sheetedit-slicerbox.readonly .sheetedit-slicer-clear { display:none; }
  `;
  document.head.appendChild(s);
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
  injectStyles();
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
      box.style.cssText += `left:${x}px;top:${y}px;width:${w}px;height:${h}px`;
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
