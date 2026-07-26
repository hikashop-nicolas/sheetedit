import type { Sheet } from "../model";
import { maxOutlineLevel, outlineGroups, showOutlineLevel } from "../outline";

// The row-outline gutter: the strip of bars and +/- buttons Excel draws left of the row numbers.
// One column per outline level, plus a level header at the top with the 1/2/3 buttons. The strip is
// an overlay glued to the grid's vertical scroll, the same pattern the chart / slicer layers use;
// the row-number column is widened to make room, so nothing overlaps.

/** Width in px of one outline level column. */
export const OUTLINE_COL_W = 14;

/** How wide the gutter must be for a sheet (0 when nothing is grouped). */
export function outlineGutterWidth(sheet: Sheet | undefined): number {
  const levels = sheet ? maxOutlineLevel(sheet, "row") : 0;
  return levels ? (levels + 1) * OUTLINE_COL_W : 0;
}

export interface OutlineLayerDeps {
  wrap: HTMLElement;
  gridScroll: HTMLElement;
  getSheet: () => Sheet | undefined;
  /**
   * Where a row actually sits, measured from the rendered grid. The declared row height and the
   * rendered one can differ (a row's content can be taller than its <row ht>), so the gutter reads
   * the DOM rather than the layout model; a row outside the rendered window returns null.
   */
  geom: () => { rowRect: (r: number) => { top: number; height: number } | null; headerH: number; totalRows: number };
  /** A group's button was clicked: collapse or expand it. */
  onToggle: (level: number, line: number, collapse: boolean) => void;
  /** The level header's 1..N buttons. */
  onLevel: (level: number) => void;
}


export function setupOutlineLayer(deps: OutlineLayerDeps): { refresh(): void; teardown(): void } {
  const { wrap, gridScroll } = deps;
  const layer = document.createElement("div");
  layer.className = "sheetedit-outline";
  const inner = document.createElement("div");
  inner.className = "sheetedit-outline-inner";
  const head = document.createElement("div");
  head.className = "sheetedit-outline-head";
  layer.append(head, inner);
  wrap.appendChild(layer);

  const syncScroll = (): void => { inner.style.transform = `translateY(${-gridScroll.scrollTop}px)`; };

  const refresh = (): void => {
    const sheet = deps.getSheet();
    const width = outlineGutterWidth(sheet);
    layer.style.display = width ? "block" : "none";
    inner.textContent = "";
    head.textContent = "";
    if (!width || !sheet) return;
    const g = deps.geom();
    const gr = gridScroll.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    layer.style.left = `${gr.left - wr.left}px`;
    layer.style.top = `${gr.top - wr.top}px`;
    layer.style.width = `${width}px`;
    layer.style.height = `${gr.height}px`;
    head.style.width = `${width}px`;
    head.style.height = `${g.headerH}px`;
    inner.style.top = `${g.headerH}px`;
    inner.style.height = `${Math.max(0, gr.height - g.headerH)}px`;

    const levels = maxOutlineLevel(sheet, "row");
    for (let l = 1; l <= levels + 1; l++) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-outline-lvl";
      b.textContent = String(l);
      b.title = `Level ${l}`;
      b.dataset.level = String(l);
      // Button N shows levels 0..N-1, so the last button expands everything.
      b.addEventListener("click", () => deps.onLevel(l - 1));
      head.appendChild(b);
    }

    // One vertical bar per group, with the toggle button on its summary line. Only the part of a
    // group inside the rendered window can be measured, which is also the only part on screen.
    for (const grp of outlineGroups(sheet, "row", g.totalRows)) {
      const x = (grp.level - 1) * OUTLINE_COL_W + 6;
      let first: { top: number; height: number } | null = null, last: { top: number; height: number } | null = null;
      for (let r = grp.from; r <= grp.to; r++) {
        const rc = g.rowRect(r);
        if (!rc) continue;
        first ??= rc;
        last = rc;
      }
      if (!grp.collapsed && first && last) {
        const top = first.top, bottom = last.top + last.height;
        const bar = document.createElement("div");
        bar.className = "sheetedit-outline-bar";
        Object.assign(bar.style, { left: `${x}px`, top: `${top}px`, height: `${Math.max(0, bottom - top - 2)}px` });
        inner.appendChild(bar);
        const foot = document.createElement("div");
        foot.className = "sheetedit-outline-foot";
        Object.assign(foot.style, { left: `${x}px`, top: `${bottom - 3}px` });
        inner.appendChild(foot);
      }
      // A group whose summary line is itself hidden (by an outer collapsed group) draws nothing:
      // its button would pile up on the outer group's line.
      if (sheet.hiddenRows?.has(grp.summary)) continue;
      const sc = g.rowRect(grp.summary);
      if (!sc) continue;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "sheetedit-outline-btn";
      btn.textContent = grp.collapsed ? "+" : "−";
      btn.title = grp.collapsed ? "Expand group" : "Collapse group";
      btn.dataset.group = `${grp.level}:${grp.from}`;
      Object.assign(btn.style, { left: `${x - 5}px`, top: `${sc.top + Math.max(0, (sc.height - 11) / 2)}px` });
      btn.addEventListener("click", () => deps.onToggle(grp.level, grp.from, !grp.collapsed));
      inner.appendChild(btn);
    }
    syncScroll();
  };

  gridScroll.addEventListener("scroll", syncScroll, { passive: true });
  return {
    refresh,
    teardown() { gridScroll.removeEventListener("scroll", syncScroll); layer.remove(); },
  };
}

/** Re-export so the host can drive the level buttons without importing the model module twice. */
export { showOutlineLevel };
