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
  /** The grid's scroll containers, top-most first (two when a row split is on). */
  panes: () => HTMLElement[];
  getSheet: () => Sheet | undefined;
  geom: () => ChartGeom;
  onEdit?: (sh: SheetShape) => void;      // after a move/resize
  onActivate?: (sh: SheetShape) => void;  // double-click -> edit properties
  onDelete?: (sh: SheetShape) => void;    // the selected shape's delete handle
  editable?: () => boolean;
}

const STYLE_ID = "sheetedit-shape-style";
function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-shapelayer { position:absolute; overflow:hidden; pointer-events:none; z-index:5; }
    .sheetedit-shapelayer-inner { position:absolute; inset:0; }
    .sheetedit-shapebox { position:absolute; }
    .sheetedit-shapebox svg { width:100%; height:100%; display:block; overflow:visible; pointer-events:none; }
    .sheetedit-shapebox.editable { pointer-events:auto; cursor:move; }
    .sheetedit-shapebox.selected { outline:1.5px dashed var(--sheetedit-accent,#4c8bf5); outline-offset:2px; }
    .sheetedit-shape-resize { position:absolute; right:-5px; bottom:-5px; width:12px; height:12px; border-radius:3px;
      background:var(--sheetedit-accent,#4c8bf5); border:1.5px solid #fff; cursor:nwse-resize; pointer-events:auto; display:none; }
    .sheetedit-shapebox.selected .sheetedit-shape-resize { display:block; }
    .sheetedit-shape-del { position:absolute; right:-9px; top:-9px; width:16px; height:16px; border-radius:50%; padding:0;
      background:#e03131; color:#fff; border:1.5px solid #fff; cursor:pointer; pointer-events:auto; display:none;
      font:700 11px/13px sans-serif; text-align:center; }
    .sheetedit-shapebox.selected .sheetedit-shape-del { display:block; }
  `;
  document.head.appendChild(s);
}

/** Build the SVG markup for one shape at the given pixel size. */
export function shapeSvg(sh: SheetShape, w: number, h: number): string {
  const fill = sh.fill ?? "none";
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
  return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">${body}${label}</svg>`;
}

export function setupShapeLayer(deps: ShapeLayerDeps): { refresh(): void; teardown(): void } {
  injectStyles();
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
