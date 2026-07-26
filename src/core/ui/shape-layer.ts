import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet, SheetShape } from "../model";
import type { ChartGeom } from "./chart-overlay";
import { shapePoints } from "../shape-geom";

// An overlay that floats a sheet's drawing shapes over the grid as SVG, anchored to cells and glued
// while scrolling (the same layer pattern as the image / chart overlays). Shapes can be selected,
// moved (drag the box) and resized (drag the corner handle); the new cell anchor is committed to the
// model and the host persists it. Double-click opens a property editor (fill / outline / text).

export interface ShapeLayerDeps {
  wrap: HTMLElement;
  /** The grid's viewports: the element plus whether it draws the header / row numbers. */
  panes: () => { el: HTMLElement; header: boolean; rowHeader: boolean }[];
  getSheet: () => Sheet | undefined;
  geom: () => ChartGeom;
  onEdit?: (sh: SheetShape) => void;      // after a move/resize
  onActivate?: (sh: SheetShape) => void;  // double-click -> edit properties
  onDelete?: (sh: SheetShape) => void;    // the selected shape's delete handle
  editable?: () => boolean;
}


// SVG gradient ids have to be unique in the document, since a fill refers to one by id and the
// first match wins across every overlay on the page.
let gradSeq = 0;

/** A <linearGradient> def for a shape's gradient fill, plus the url() that names it. */
function gradientDef(g: NonNullable<SheetShape["fillGradient"]>): { def: string; url: string } {
  const id = `sheetedit-grad-${++gradSeq}`;
  // DrawingML measures the angle clockwise from east, which is the direction SVG's y-down user
  // space already runs, so the vector goes straight across the bounding box through its centre.
  const rad = (g.angle * Math.PI) / 180;
  const dx = Math.cos(rad), dy = Math.sin(rad);
  const [x1, y1, x2, y2] = [0.5 - dx / 2, 0.5 - dy / 2, 0.5 + dx / 2, 0.5 + dy / 2];
  const stops = g.stops
    .map((s) => `<stop offset="${Math.max(0, Math.min(1, s.pos))}" stop-color="${s.color}"/>`)
    .join("");
  return {
    def: `<defs><linearGradient id="${id}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops}</linearGradient></defs>`,
    url: `url(#${id})`,
  };
}

/** Build the SVG markup for one shape at the given pixel size. */
export function shapeSvg(sh: SheetShape, w: number, h: number): string {
  const grad = sh.fillGradient?.stops.length ? gradientDef(sh.fillGradient) : undefined;
  const fill = grad?.url ?? sh.fill ?? "none";
  const stroke = sh.stroke ?? (sh.fill ? "none" : "#000000");
  const sw = sh.strokeWidth ?? 1;
  const inset = sw / 2; // keep the stroke inside the box
  const iw = Math.max(0, w - sw), ih = Math.max(0, h - sw);
  let body: string;
  switch (sh.geom) {
    case "ellipse":
      body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${iw / 2}" ry="${ih / 2}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    case "roundRect": {
      const r = Math.min(w, h) * 0.15;
      body = `<rect x="${inset}" y="${inset}" width="${iw}" height="${ih}" rx="${r}" ry="${r}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    }
    case "line":
      body = `<line x1="${inset}" y1="${inset}" x2="${w - inset}" y2="${h - inset}" stroke="${stroke === "none" ? "#000000" : stroke}" stroke-width="${sw}"/>`;
      break;
    default: {
      // Polygon shapes (triangle / diamond / hexagon / pentagon / star / arrow / parallelogram),
      // inset a touch so the stroke stays inside the box; else a plain rectangle.
      const pts = shapePoints(sh.geom, iw, ih);
      if (pts) body = `<polygon points="${pts.map(([x, y]) => `${x + inset},${y + inset}`).join(" ")}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}" stroke-linejoin="round"/>`;
      else body = `<rect x="${inset}" y="${inset}" width="${iw}" height="${ih}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
    }
  }
  let label = "";
  if (sh.text && sh.geom !== "line") {
    const esc = sh.text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    label = `<text x="${w / 2}" y="${h / 2}" text-anchor="middle" dominant-baseline="central" fill="${sh.textColor ?? "#000000"}" font-size="13" font-family="sans-serif">${esc}</text>`;
  }
  return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">${grad?.def ?? ""}${body}${label}</svg>`;
}

export function setupShapeLayer(deps: ShapeLayerDeps): { refresh(): void; teardown(): void } {
  const hosts = setupOverlayHosts({
    wrap: deps.wrap,
    panes: deps.panes,
    geom: () => { const g = deps.geom(); return { rnW: g.rnW, headerH: g.headerH, yOfRow: g.yOfRow }; },
    className: "sheetedit-shapelayer",
    innerClassName: "sheetedit-shapelayer-inner",
  });

  let selected: SheetShape | null = null;
  const boxes = new Map<SheetShape, HTMLElement>();


  const rectToAnchor = (sh: SheetShape, x: number, y: number, w: number, h: number): void => {
    const g = deps.geom();
    const set = (px: number, at: (p: number) => number, of: (i: number) => number): [number, number] => { const i = Math.max(1, at(px)); return [i, Math.max(0, px - of(i))]; };
    const [fc, fco] = set(x, g.colAt, g.xOfCol);
    const [fr, fro] = set(y, g.rowAt, g.yOfRow);
    const [tc, tco] = set(x + w, g.colAt, g.xOfCol);
    const [tr, tro] = set(y + h, g.rowAt, g.yOfRow);
    sh.anchor = { fromCol: fc, fromRow: fr, fromColOff: fco, fromRowOff: fro, toCol: tc, toRow: tr, toColOff: tco, toRowOff: tro };
    if (sh.odsShapeEl && sh.odsAnchorCol != null && sh.odsAnchorRow != null) {
      sh.odsFrame = { x: x - g.xOfCol(sh.odsAnchorCol), y: y - g.yOfRow(sh.odsAnchorRow), w, h };
    }
    sh.dirty = true;
  };

  const select = (sh: SheetShape | null): void => {
    selected = sh;
    for (const [s, b] of boxes) b.classList.toggle("selected", s === sh);
  };

  const attachDrag = (box: HTMLElement, handle: HTMLElement, sh: SheetShape): void => {
    const start = (e: PointerEvent, mode: "move" | "resize"): void => {
      e.preventDefault();
      e.stopPropagation();
      select(sh);
      const sx = e.clientX, sy = e.clientY;
      const x0 = parseFloat(box.style.left) || 0, y0 = parseFloat(box.style.top) || 0;
      const w0 = box.offsetWidth, h0 = box.offsetHeight;
      const svg = box.querySelector("svg");
      const onMove = (ev: PointerEvent): void => {
        const dx = ev.clientX - sx, dy = ev.clientY - sy;
        if (mode === "move") { box.style.left = `${Math.max(0, x0 + dx)}px`; box.style.top = `${Math.max(0, y0 + dy)}px`; }
        else {
          const w = Math.max(8, w0 + dx), h = Math.max(8, h0 + dy);
          box.style.width = `${w}px`; box.style.height = `${h}px`;
          if (svg) svg.outerHTML = shapeSvg(sh, w, h); // redraw so geometry tracks the new size
        }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        const x = parseFloat(box.style.left) || 0, y = parseFloat(box.style.top) || 0;
        // A plain click (no drag) must not commit or refresh, or a double-click's box is replaced
        // between its two clicks and never fires.
        if (x === x0 && y === y0 && box.offsetWidth === w0 && box.offsetHeight === h0) return;
        rectToAnchor(sh, x, y, box.offsetWidth, box.offsetHeight);
        deps.onEdit?.(sh);
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    box.addEventListener("pointerdown", (e) => { if (e.target !== handle) start(e, "move"); });
    handle.addEventListener("pointerdown", (e) => start(e, "resize"));
  };

  const refresh = (): void => {
    hosts.clear();
    boxes.clear();
    const sheet = deps.getSheet();
    const shapes = sheet?.shapes ?? [];
    hosts.setVisible(shapes.length > 0);
    if (!shapes.length) { selected = null; return; }
    const editable = deps.editable?.() ?? false;
    const g = deps.geom();
    for (const sh of shapes) {
      const a = sh.anchor;
      const x = g.xOfCol(a.fromCol) + a.fromColOff;
      const y = g.yOfRow(a.fromRow) + a.fromRowOff;
      const w = Math.max(1, g.xOfCol(a.toCol) + a.toColOff - x);
      const h = Math.max(1, g.yOfRow(a.toRow) + a.toRowOff - y);
      const box = document.createElement("div");
      box.className = "sheetedit-shapebox" + (editable ? " editable" : "") + (sh === selected ? " selected" : "");
      box.style.left = `${x}px`;
      box.style.top = `${y}px`;
      box.style.width = `${w}px`;
      box.style.height = `${h}px`;
      box.innerHTML = shapeSvg(sh, w, h);
      if (editable) {
        const handle = document.createElement("div");
        handle.className = "sheetedit-shape-resize";
        box.appendChild(handle);
        attachDrag(box, handle, sh);
        box.title = deps.onActivate ? "Double-click to edit" : "";
        box.addEventListener("dblclick", (e) => { e.preventDefault(); e.stopPropagation(); select(sh); deps.onActivate?.(sh); });
        if (deps.onDelete) {
          const del = document.createElement("button");
          del.className = "sheetedit-shape-del";
          del.type = "button";
          del.textContent = "×";
          del.title = "Delete shape";
          del.addEventListener("pointerdown", (e) => e.stopPropagation());
          del.addEventListener("click", (e) => { e.preventDefault(); e.stopPropagation(); deps.onDelete!(sh); });
          box.appendChild(del);
        }
      }
      hosts.hostFor(sh.anchor.fromRow).appendChild(box);
      boxes.set(sh, box);
    }
    hosts.layout();
  };


  const onGridDown = (): void => { if (selected) select(null); };
  hosts.onPanes("pointerdown", onGridDown as (e: Event) => void);
  return {
    refresh,
    teardown() { hosts.teardown(); },
  };
}
