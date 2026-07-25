import type { Sheet } from "../model";

// The pane dividers: the bars sitting on a sheet's frozen / split boundary. Dragging one moves the
// boundary (snapping to the nearest line edge on screen), and double-clicking removes it.
//
// Both the position and the extent are measured off the RENDERED grid, never from the layout model:
// a row's declared height and its rendered height differ (content can be taller than <row ht>), and
// a container's rect can still be stale when a refresh runs before layout settles. The bars live in
// a clipping box laid over the grid's client area, so they can never spill past it.

export interface PaneDividerDeps {
  wrap: HTMLElement;
  /** Every scroll container making up the grid, top-most first (two when a row split is on). */
  panes: () => HTMLElement[];
  getSheet: () => Sheet | undefined;
  /** Where the boundary is on screen, measured from the rendered headers. */
  geom: () => {
    /** Client y of the boundary below `rows`, or null when that row is not rendered. */
    rowBoundaryY: (rows: number) => number | null;
    /** Client x of the boundary right of `cols`, or null when that column is not rendered. */
    colBoundaryX: (cols: number) => number | null;
    /** Client y where the first data row starts (below the column header). */
    bodyTop: () => number;
    nearestRow: (clientY: number) => number;
    nearestCol: (clientX: number) => number;
  };
  onMove: (rows: number, cols: number) => void;
}

const STYLE_ID = "sheetedit-pane-style";
function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    /* Clips the bars to the grid area, so neither can reach the scrollbar or the sheet tabs. */
    .sheetedit-panediv-layer { position:absolute; overflow:hidden; pointer-events:none; z-index:11; }
    .sheetedit-panediv { position:absolute; pointer-events:auto; background:var(--sheetedit-accent,#6e7bff); opacity:.55; }
    .sheetedit-panediv:hover, .sheetedit-panediv.dragging { opacity:1; }
    .sheetedit-panediv-h { left:0; right:0; height:3px; cursor:row-resize; }
    .sheetedit-panediv-v { top:0; bottom:0; width:3px; cursor:col-resize; }
    /* A wider invisible grab area, so the thin bar is still easy to catch. */
    .sheetedit-panediv::after { content:""; position:absolute; inset:-4px; }
  `;
  document.head.appendChild(s);
}

export function setupPaneDividers(deps: PaneDividerDeps): { refresh(): void; teardown(): void } {
  injectStyles();
  const { wrap } = deps;
  const layer = document.createElement("div");
  layer.className = "sheetedit-panediv-layer";
  layer.style.display = "none";
  wrap.appendChild(layer);
  let dragging: "row" | "col" | null = null;
  let pending = 0;

  const paint = (): void => {
    if (dragging) return; // a live drag owns the bars
    layer.textContent = "";
    const sheet = deps.getSheet();
    const rows = sheet?.freeze?.rows ?? 0, cols = sheet?.freeze?.cols ?? 0;
    const els = deps.panes();
    if ((!rows && !cols) || !els.length) { layer.style.display = "none"; return; }
    layer.style.display = "block";
    const wr = wrap.getBoundingClientRect();
    const firstEl = els[0]!, lastEl = els[els.length - 1]!;
    const first = firstEl.getBoundingClientRect(), last = lastEl.getBoundingClientRect();
    // The clipping box spans every pane, so a column bar crosses a row split's two viewports. It
    // uses the CLIENT box, so a bar never runs over a scrollbar.
    const left = first.left, top = first.top;
    const right = first.left + firstEl.clientWidth;
    const bottom = last.top + lastEl.clientHeight;
    layer.style.left = `${left - wr.left}px`;
    layer.style.top = `${top - wr.top}px`;
    layer.style.width = `${Math.max(0, right - left)}px`;
    layer.style.height = `${Math.max(0, bottom - top)}px`;

    const g = deps.geom();
    const bar = (kind: "row" | "col", pos: number, title: string): void => {
      const el = document.createElement("div");
      el.className = `sheetedit-panediv sheetedit-panediv-${kind === "row" ? "h" : "v"}`;
      if (kind === "row") el.style.top = `${pos - top - 1}px`;
      else el.style.left = `${pos - left - 1}px`;
      el.title = title;
      el.dataset.pane = kind;
      el.addEventListener("pointerdown", (e) => startDrag(e, kind));
      el.addEventListener("dblclick", () => deps.onMove(kind === "row" ? 0 : rows, kind === "col" ? 0 : cols));
      layer.appendChild(el);
    };
    const hint = "Drag to move the split, double-click to remove";
    if (rows > 0) { const y = g.rowBoundaryY(rows); if (y != null) bar("row", y, hint); }
    if (cols > 0) { const x = g.colBoundaryX(cols); if (x != null) bar("col", x, hint); }
  };

  // Measuring right after a render can read a container that has not been laid out yet, so settle
  // on the next frame; a plain refresh() call still paints immediately for the tests.
  const refresh = (): void => {
    paint();
    if (pending) cancelAnimationFrame(pending);
    pending = typeof requestAnimationFrame === "function" ? requestAnimationFrame(() => { pending = 0; paint(); }) : 0;
  };

  const startDrag = (e: PointerEvent, kind: "row" | "col"): void => {
    const sheet = deps.getSheet();
    if (!sheet) return;
    e.preventDefault();
    dragging = kind;
    const el = e.currentTarget as HTMLElement;
    el.classList.add("dragging");
    el.setPointerCapture(e.pointerId);
    const g = deps.geom();
    const box = layer.getBoundingClientRect();
    const home = { top: el.style.top, left: el.style.left };
    const onMove = (ev: PointerEvent): void => {
      // Follow the pointer live; the boundary itself only moves on release.
      if (kind === "row") el.style.top = `${ev.clientY - box.top - 1}px`;
      else el.style.left = `${ev.clientX - box.left - 1}px`;
    };
    const onUp = (ev: PointerEvent): void => {
      el.releasePointerCapture(ev.pointerId);
      el.removeEventListener("pointermove", onMove);
      el.removeEventListener("pointerup", onUp);
      dragging = null;
      const rows = sheet.freeze?.rows ?? 0, cols = sheet.freeze?.cols ?? 0;
      const next = kind === "row" ? Math.max(0, g.nearestRow(ev.clientY)) : Math.max(0, g.nearestCol(ev.clientX));
      // A plain click lands back on the same line. Rebuilding the bar - whether by committing the
      // move or by repainting - would swap the element out between the two clicks of a double-click,
      // so the dblclick would never fire. Put THIS element back where it was instead.
      if (next === (kind === "row" ? rows : cols)) {
        el.classList.remove("dragging");
        el.style.top = home.top;
        el.style.left = home.left;
        return;
      }
      if (kind === "row") deps.onMove(next, cols); else deps.onMove(rows, next);
    };
    el.addEventListener("pointermove", onMove);
    el.addEventListener("pointerup", onUp);
  };

  return { refresh, teardown() { if (pending) cancelAnimationFrame(pending); layer.remove(); } };
}
