import type { Sheet, SheetSlicer } from "../model";
import type { ChartGeom } from "./chart-overlay";

// The slicer overlay: a titled panel of item buttons floated over the grid, anchored to cells and
// glued while scrolling. Clicking an item filters the linked pivot(s): a plain click selects only
// that item, ctrl/cmd-click toggles it within the current selection, and the "clear filter" button
// selects everything again. The host recomputes the pivots and persists the new selection.

export interface SlicerLayerDeps {
  wrap: HTMLElement;
  gridScroll: HTMLElement;
  getSheet: () => Sheet | undefined;
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
      background:transparent; color:inherit; font:inherit; padding:3px 6px; cursor:pointer;
      text-align:left; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; opacity:.45; }
    .sheetedit-slicer-item:hover { border-color:var(--sheetedit-accent,#4c8bf5); }
    .sheetedit-slicer-item.on { opacity:1; background:var(--sheetedit-accent,#4c8bf5); color:#fff;
      border-color:var(--sheetedit-accent,#4c8bf5); }
  `;
  document.head.appendChild(s);
}

export function setupSlicerLayer(deps: SlicerLayerDeps): { refresh(): void; teardown(): void } {
  injectStyles();
  const { wrap, gridScroll } = deps;
  const layer = document.createElement("div");
  layer.className = "sheetedit-slicerlayer";
  const inner = document.createElement("div");
  inner.className = "sheetedit-slicerlayer-inner";
  layer.appendChild(inner);
  wrap.appendChild(layer);

  const positionLayer = (): void => {
    const g = deps.geom();
    const gr = gridScroll.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    layer.style.left = `${gr.left - wr.left + g.rnW}px`;
    layer.style.top = `${gr.top - wr.top + g.headerH}px`;
    layer.style.width = `${Math.max(0, gr.width - g.rnW)}px`;
    layer.style.height = `${Math.max(0, gr.height - g.headerH)}px`;
  };
  const syncScroll = (): void => { inner.style.transform = `translate(${-gridScroll.scrollLeft}px, ${-gridScroll.scrollTop}px)`; };

  const refresh = (): void => {
    inner.textContent = "";
    const sheet = deps.getSheet();
    const slicers = sheet?.slicers ?? [];
    layer.style.display = slicers.length ? "block" : "none";
    if (!slicers.length) return;
    const g = deps.geom();
    for (const sl of slicers) {
      const a = sl.anchor;
      // A slicer with no drawing anchor still needs somewhere to live.
      const x = a ? g.xOfCol(a.fromCol) + a.fromColOff : 20;
      const y = a ? g.yOfRow(a.fromRow) + a.fromRowOff : 20;
      const w = a ? Math.max(120, g.xOfCol(a.toCol) + a.toColOff - x) : 160;
      const h = a ? Math.max(80, g.yOfRow(a.toRow) + a.toRowOff - y) : 180;
      const box = document.createElement("div");
      box.className = "sheetedit-slicerbox";
      box.style.cssText += `left:${x}px;top:${y}px;width:${w}px;height:${h}px`;
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
      inner.appendChild(box);
    }
    positionLayer();
    syncScroll();
  };

  gridScroll.addEventListener("scroll", syncScroll, { passive: true });
  return {
    refresh,
    teardown() { gridScroll.removeEventListener("scroll", syncScroll); layer.remove(); },
  };
}
