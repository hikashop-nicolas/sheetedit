import type { Sheet } from "../model";

// The pane dividers: the bars sitting on a sheet's frozen / split boundary. Dragging one moves the
// boundary (snapping to the nearest line edge on screen), and double-clicking removes it. Both
// kinds render the same, because the leading pane stays put either way; only what the file records
// differs, so dragging a SPLIT keeps it a split where the format can say so.

export interface PaneDividerDeps {
  wrap: HTMLElement;
  gridScroll: HTMLElement;
  getSheet: () => Sheet | undefined;
  /** Where the boundary sits on screen, and which line a pointer position lands on. */
  geom: () => {
    headerH: number;
    gutterW: number;
    /** Left edge of the boundary between the frozen and scrolling columns, in grid coordinates. */
    boundary: (rows: number, cols: number) => { x: number; y: number };
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
    .sheetedit-panediv { position:absolute; z-index:11; background:var(--sheetedit-accent,#6e7bff); opacity:.55; }
    .sheetedit-panediv:hover, .sheetedit-panediv.dragging { opacity:1; }
    .sheetedit-panediv-h { height:3px; cursor:row-resize; }
    .sheetedit-panediv-v { width:3px; cursor:col-resize; }
    /* A wider invisible grab area, so the thin bar is still easy to catch. */
    .sheetedit-panediv::after { content:""; position:absolute; inset:-3px; }
  `;
  document.head.appendChild(s);
}

export function setupPaneDividers(deps: PaneDividerDeps): { refresh(): void; teardown(): void } {
  injectStyles();
  const { wrap, gridScroll } = deps;
  const bars: HTMLElement[] = [];
  let dragging: "row" | "col" | null = null;

  const clear = (): void => { for (const b of bars.splice(0)) b.remove(); };

  const refresh = (): void => {
    if (dragging) return; // a live drag owns the bars
    clear();
    const sheet = deps.getSheet();
    const rows = sheet?.freeze?.rows ?? 0, cols = sheet?.freeze?.cols ?? 0;
    if (!rows && !cols) return;
    const g = deps.geom();
    const gr = gridScroll.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    const at = g.boundary(rows, cols);
    const left = gr.left - wr.left, top = gr.top - wr.top;

    const bar = (kind: "row" | "col", style: string, title: string): HTMLElement => {
      const el = document.createElement("div");
      el.className = `sheetedit-panediv sheetedit-panediv-${kind === "row" ? "h" : "v"}`;
      el.style.cssText += style;
      el.title = title;
      el.dataset.pane = kind;
      el.addEventListener("pointerdown", (e) => startDrag(e, kind));
      el.addEventListener("dblclick", () => deps.onMove(kind === "row" ? 0 : rows, kind === "col" ? 0 : cols));
      wrap.appendChild(el);
      bars.push(el);
      return el;
    };
    if (rows > 0) bar("row", `left:${left + g.gutterW}px;top:${top + at.y - 1}px;width:${Math.max(0, gr.width - g.gutterW)}px`, "Drag to move the split, double-click to remove");
    if (cols > 0) bar("col", `left:${left + at.x - 1}px;top:${top + g.headerH}px;height:${Math.max(0, gr.height - g.headerH)}px`, "Drag to move the split, double-click to remove");
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
    const wr = wrap.getBoundingClientRect();
    const home = { top: el.style.top, left: el.style.left };
    const onMove = (ev: PointerEvent): void => {
      // Follow the pointer live; the boundary itself only moves on release.
      if (kind === "row") el.style.top = `${ev.clientY - wr.top - 1}px`;
      else el.style.left = `${ev.clientX - wr.left - 1}px`;
    };
    const onUp = (ev: PointerEvent): void => {
      el.releasePointerCapture(ev.pointerId);
      el.removeEventListener("pointermove", onMove);
      el.removeEventListener("pointerup", onUp);
      dragging = null;
      const rows = sheet.freeze?.rows ?? 0, cols = sheet.freeze?.cols ?? 0;
      const next = kind === "row" ? Math.max(0, g.nearestRow(ev.clientY)) : Math.max(0, g.nearestCol(ev.clientX));
      // A plain click lands back on the same line. Rebuilding the bar - whether by committing the
      // move or by refreshing - would swap the element out between the two clicks of a double-click,
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

  return { refresh, teardown() { clear(); } };
}
