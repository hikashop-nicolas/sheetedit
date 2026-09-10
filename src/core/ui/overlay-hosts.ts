// Multi-pane plumbing for the floating layers (charts, images, shapes, slicers, timelines, pivot
// tags). Every one of them is the same shape: a box inside the grid's own scroller, positioned in
// sheet coordinates, with children placed relative to it.
//
// The layer lives INSIDE the viewport it belongs to, so the browser scrolls it with the cells for
// free. Laying it over the grid and translating it from a scroll listener meant it moved a frame
// or more after the content did: on any real scroll the shapes and images visibly trailed the
// sheet, and no amount of listener tuning closes that gap, because the content is scrolled on the
// compositor while the correction is a main-thread style change.
//
// With a row split there are two scrolling viewports, so each layer gets one box PER pane and an
// object is appended to the pane that currently shows its anchor row. Only the top pane carries the
// column header, so the lower box starts at its own top edge.

export interface OverlayHostDeps {
  /** The editor root. Kept for callers that place their own chrome against it. */
  wrap: HTMLElement;
  /** The grid's viewports: the element plus whether it draws the header / row numbers. */
  panes: () => { el: HTMLElement; header: boolean; rowHeader: boolean }[];
  geom: () => { rnW: number; headerH: number; yOfRow: (r: number) => number };
  /** Class for the outer box, so each layer keeps its own styling hooks. */
  className: string;
  /** Class for the scrolled inner element. */
  innerClassName: string;
}

export interface OverlayHosts {
  /** The inner element to append an object anchored at (`row`, `col`) into. */
  hostFor: (row: number, col?: number) => HTMLElement;
  /** The first (top) pane's inner, for objects with no row anchor. */
  main: () => HTMLElement;
  /** Empty every pane's inner. */
  clear: () => void;
  /** Show or hide the whole overlay. */
  setVisible: (on: boolean) => void;
  /** Re-place the boxes at their pane's grid origin (past the header and row numbers). */
  layout: () => void;
  /** Subscribe to an event on every pane's scroll container. */
  onPanes: (type: string, fn: (e: Event) => void, opts?: AddEventListenerOptions) => void;
  /** All inner elements, for lookups that must search everything on screen. */
  inners: () => HTMLElement[];
  teardown: () => void;
}

export function setupOverlayHosts(deps: OverlayHostDeps): OverlayHosts {
  const boxes: { pane: HTMLElement; layer: HTMLElement; inner: HTMLElement; header: boolean; rowHeader: boolean }[] = [];
  const listeners: { type: string; fn: (e: Event) => void; opts?: AddEventListenerOptions }[] = [];
  let visible = true;

  /** Match the boxes to the current pane list, creating and dropping as the split comes and goes. */
  const build = (): void => {
    const panes = deps.panes().map((p) => p.el);
    const specs = deps.panes();
    while (boxes.length > panes.length) {
      const b = boxes.pop()!;
      for (const l of listeners) b.pane.removeEventListener(l.type, l.fn, l.opts);
      b.layer.remove();
    }
    panes.forEach((pane, i) => {
      if (boxes[i] && boxes[i]!.pane === pane) {
        boxes[i]!.header = specs[i]!.header;
        boxes[i]!.rowHeader = specs[i]!.rowHeader;
        // A full grid render empties the scroller, which now holds the layer too. Re-attach it
        // rather than rebuild: the boxes and their listeners are still good.
        if (!boxes[i]!.layer.isConnected) pane.appendChild(boxes[i]!.layer);
        return;
      }
      if (boxes[i]) {
        const old = boxes[i]!;
        for (const l of listeners) old.pane.removeEventListener(l.type, l.fn, l.opts);
        old.layer.remove();
      }
      const layer = document.createElement("div");
      layer.className = deps.className;
      const inner = document.createElement("div");
      inner.className = deps.innerClassName;
      layer.appendChild(inner);
      pane.appendChild(layer); // inside the scroller: the browser moves it with the cells
      const box = { pane, layer, inner, header: specs[i]!.header, rowHeader: specs[i]!.rowHeader };
      boxes[i] = box;
      for (const l of listeners) pane.addEventListener(l.type, l.fn, l.opts);
    });
  };

  const layout = (): void => {
    build();
    const g = deps.geom();
    boxes.forEach((b) => {
      // In the scroller's own content coordinates, so the offsets are the chrome each band draws:
      // the header on the top band, the row numbers on the left one. The scroller does the
      // clipping, and the sticky header and row numbers outrank the layers and cover them.
      const headerH = b.header ? g.headerH : 0;
      const rnW = b.rowHeader ? g.rnW : 0;
      b.layer.style.display = visible ? "block" : "none";
      b.layer.style.left = `${rnW}px`;
      b.layer.style.top = `${headerH}px`;
    });
  };

  /**
   * The pane that actually shows `row` on screen, decided from the RENDERED row element rather than
   * the layout model: the model's uniform row height differs from the rendered one by a pixel or so,
   * which is enough to put an object in the wrong pane. Falls back to the top pane.
   */
  const hostFor = (row: number, col?: number): HTMLElement => {
    build(); // a split may have appeared since the last layout
    if (boxes.length < 2) return boxes[0]!.inner;
    const g = deps.geom();
    let best = boxes[0]!.inner, bestSeen = -1;
    for (const b of boxes) {
      const pr = b.pane.getBoundingClientRect();
      const top = pr.top + (b.header ? g.headerH : 0), bottom = pr.top + b.pane.clientHeight;
      const left = pr.left + (b.rowHeader ? g.rnW : 0), right = pr.left + b.pane.clientWidth;
      const th = b.pane.querySelector(`th.rownum[data-r="${row}"]`) as HTMLElement | null;
      // The right band draws no row numbers, so fall back to any cell of that row there.
      const rowEl = th ?? (b.pane.querySelector(`td[data-rc^="${row}:"]`) as HTMLElement | null);
      if (!rowEl) continue;
      const rr = rowEl.getBoundingClientRect();
      let seen = Math.min(rr.bottom, bottom) - Math.max(rr.top, top);
      if (col != null) {
        const ch = b.pane.querySelector(`th.colhead[data-c="${col}"]`) as HTMLElement | null;
        const colEl = ch ?? (b.pane.querySelector(`td[data-rc$=":${col}"]`) as HTMLElement | null);
        if (!colEl) continue;
        const cr = colEl.getBoundingClientRect();
        seen = Math.min(seen, Math.min(cr.right, right) - Math.max(cr.left, left));
      }
      if (seen > bestSeen) { bestSeen = seen; best = b.inner; }
    }
    return best;
  };

  build();
  return {
    hostFor,
    main: () => boxes[0]!.inner,
    clear: () => { build(); for (const b of boxes) b.inner.textContent = ""; },
    setVisible: (on: boolean) => { visible = on; for (const b of boxes) b.layer.style.display = on ? "block" : "none"; },
    layout,
    onPanes: (type, fn, opts) => {
      listeners.push({ type, fn, opts });
      for (const b of boxes) b.pane.addEventListener(type, fn, opts);
    },
    inners: () => boxes.map((b) => b.inner),
    teardown: () => {
      for (const b of boxes) {
        for (const l of listeners) b.pane.removeEventListener(l.type, l.fn, l.opts);
        b.layer.remove();
      }
      boxes.length = 0;
    },
  };
}
