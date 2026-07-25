// Multi-pane plumbing for the floating layers (charts, images, shapes, slicers, timelines, pivot
// tags). Every one of them is the same shape: a box laid over the grid area with an inner element
// translated by the grid's scroll, and children positioned in sheet coordinates.
//
// With a row split there are two scrolling viewports, so each layer gets one box PER pane and an
// object is appended to the pane that currently shows its anchor row. Only the top pane carries the
// column header, so the lower box starts at its own top edge.

export interface OverlayHostDeps {
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
  /** Re-place the boxes over their panes and re-apply each pane's scroll offset. */
  layout: () => void;
  /** Subscribe to an event on every pane's scroll container. */
  onPanes: (type: string, fn: (e: Event) => void, opts?: AddEventListenerOptions) => void;
  /** All inner elements, for lookups that must search everything on screen. */
  inners: () => HTMLElement[];
  teardown: () => void;
}

export function setupOverlayHosts(deps: OverlayHostDeps): OverlayHosts {
  const { wrap } = deps;
  const boxes: { pane: HTMLElement; layer: HTMLElement; inner: HTMLElement; header: boolean; rowHeader: boolean }[] = [];
  const listeners: { type: string; fn: (e: Event) => void; opts?: AddEventListenerOptions }[] = [];
  let visible = true;

  const sync = (b: { pane: HTMLElement; inner: HTMLElement }): void => {
    b.inner.style.transform = `translate(${-b.pane.scrollLeft}px, ${-b.pane.scrollTop}px)`;
  };
  const onScroll = (): void => { for (const b of boxes) sync(b); };

  /** Match the boxes to the current pane list, creating and dropping as the split comes and goes. */
  const build = (): void => {
    const panes = deps.panes().map((p) => p.el);
    const specs = deps.panes();
    while (boxes.length > panes.length) {
      const b = boxes.pop()!;
      for (const l of listeners) b.pane.removeEventListener(l.type, l.fn, l.opts);
      b.pane.removeEventListener("scroll", onScroll);
      b.layer.remove();
    }
    panes.forEach((pane, i) => {
      if (boxes[i] && boxes[i]!.pane === pane) { boxes[i]!.header = specs[i]!.header; boxes[i]!.rowHeader = specs[i]!.rowHeader; return; }
      if (boxes[i]) {
        const old = boxes[i]!;
        for (const l of listeners) old.pane.removeEventListener(l.type, l.fn, l.opts);
        old.pane.removeEventListener("scroll", onScroll);
        old.layer.remove();
      }
      const layer = document.createElement("div");
      layer.className = deps.className;
      const inner = document.createElement("div");
      inner.className = deps.innerClassName;
      layer.appendChild(inner);
      wrap.appendChild(layer);
      const box = { pane, layer, inner, header: specs[i]!.header, rowHeader: specs[i]!.rowHeader };
      boxes[i] = box;
      pane.addEventListener("scroll", onScroll, { passive: true });
      for (const l of listeners) pane.addEventListener(l.type, l.fn, l.opts);
    });
  };

  const layout = (): void => {
    build();
    const g = deps.geom();
    const wr = wrap.getBoundingClientRect();
    boxes.forEach((b) => {
      const pr = b.pane.getBoundingClientRect();
      // Each band carries only its own chrome: the header on the top band, the row numbers on the
      // left one, so the boxes past a boundary start flush with their pane.
      const headerH = b.header ? g.headerH : 0;
      const rnW = b.rowHeader ? g.rnW : 0;
      b.layer.style.display = visible ? "block" : "none";
      b.layer.style.left = `${pr.left - wr.left + rnW}px`;
      b.layer.style.top = `${pr.top - wr.top + headerH}px`;
      b.layer.style.width = `${Math.max(0, b.pane.clientWidth - rnW)}px`;
      b.layer.style.height = `${Math.max(0, b.pane.clientHeight - headerH)}px`;
      sync(b);
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
        b.pane.removeEventListener("scroll", onScroll);
        b.layer.remove();
      }
      boxes.length = 0;
    },
  };
}
