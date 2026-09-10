import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet, SheetShape } from "../model";
import type { ChartGeom } from "./chart-overlay";
import { shapeOutlinePath, shapePoints } from "../shape-geom";

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
  /** Run the macro a shape names. Absent when the workbook carries no macros at all. */
  runMacro?: (name: string) => void;
  /** Tooltip prefix for a shape that has a macro assigned. */
  macroTitle?: string;
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

// Line ends. DrawingML names a shape (triangle, stealth, ...) and the renderer draws it at the
// line's end, in the line's colour, scaled to its width. A connector with no marker reads as a
// plain rule, which is the wrong statement entirely: an arrow points at something.
const MARKER_PATHS: Record<string, string> = {
  triangle: "M0,0 L6,3 L0,6 z",
  arrow: "M0,0 L6,3 L0,6 z",
  stealth: "M0,0 L6,3 L0,6 L1.6,3 z",
  diamond: "M0,3 L3,0 L6,3 L3,6 z",
  oval: "M0,3 a3,3 0 1,0 6,0 a3,3 0 1,0 -6,0",
};

/** A <marker> def for one line end, plus the url() that names it, or null for "none". */
function markerDef(type: string | undefined, color: string): { def: string; url: string } | null {
  const path = type ? MARKER_PATHS[type] : undefined;
  if (!path) return null;
  const id = `sheetedit-marker-${++gradSeq}`;
  return {
    def: `<marker id="${id}" viewBox="0 0 6 6" refX="5.4" refY="3" markerWidth="4" markerHeight="4" markerUnits="strokeWidth" orient="auto"><path d="${path}" fill="${color}"/></marker>`,
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
  let markerDefs = "";
  switch (sh.geom) {
    case "ellipse":
      body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${iw / 2}" ry="${ih / 2}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    case "roundRect": {
      const r = Math.min(w, h) * 0.15;
      body = `<rect x="${inset}" y="${inset}" width="${iw}" height="${ih}" rx="${r}" ry="${r}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    }
    case "line": {
      // A line is the diagonal of its box; flipH/flipV say which diagonal, so a flipped connector
      // runs from the other corner instead of doubling back on itself.
      const colour = stroke === "none" ? "#000000" : stroke;
      const [x1, x2] = sh.flipH ? [w - inset, inset] : [inset, w - inset];
      const [y1, y2] = sh.flipV ? [h - inset, inset] : [inset, h - inset];
      const head = markerDef(sh.headEnd, colour);
      const tail = markerDef(sh.tailEnd, colour);
      const ends = [head, tail].filter(Boolean).map((m) => m!.def).join("");
      if (ends) markerDefs = `<defs>${ends}</defs>`;
      const attrs = `${head ? ` marker-start="${head.url}"` : ""}${tail ? ` marker-end="${tail.url}"` : ""}`;
      body = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${colour}" stroke-width="${sw}"${attrs}/>`;
      break;
    }
    default: {
      // Braces and brackets: an open stroked outline, never filled.
      const path = shapeOutlinePath(sh.geom, iw, ih);
      if (path) {
        body = `<path d="${path}" transform="translate(${inset},${inset})" fill="none" stroke="${stroke === "none" ? "#000000" : stroke}" stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round"/>`;
        break;
      }
      // Polygon shapes (triangle / diamond / hexagon / pentagon / star / arrow / parallelogram),
      // inset a touch so the stroke stays inside the box; else a plain rectangle.
      const pts = shapePoints(sh.geom, iw, ih);
      if (pts) body = `<polygon points="${pts.map(([x, y]) => `${x + inset},${y + inset}`).join(" ")}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}" stroke-linejoin="round"/>`;
      else body = `<rect x="${inset}" y="${inset}" width="${iw}" height="${ih}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
    }
  }
  let label = "";
  if (sh.text && sh.geom !== "line") {
    // The text goes in a foreignObject rather than an <svg:text>: a single text node cannot wrap
    // and cannot hold the shape's paragraphs, so a label longer than its box was drawn as one
    // line running out of both sides of it. HTML in the box wraps, keeps the line breaks, and
    // can be selected and copied like any other text on the page.
    const esc = sh.text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const justify = sh.textValign === "top" ? "flex-start" : sh.textValign === "bottom" ? "flex-end" : "center";
    const style = [
      "width:100%", "height:100%", "box-sizing:border-box", "padding:2px 5px",
      "display:flex", "flex-direction:column", `justify-content:${justify}`,
      `text-align:${sh.textAlign ?? "center"}`, "white-space:pre-wrap", "overflow-wrap:break-word",
      "font:13px sans-serif", "line-height:1.25", `color:${sh.textColor ?? "#000000"}`,
    ].join(";");
    label = `<foreignObject x="0" y="0" width="${w}" height="${h}"><div xmlns="http://www.w3.org/1999/xhtml" class="sheetedit-shapetext" style="${style}">${esc}</div></foreignObject>`;
  }
  return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">${grad?.def ?? ""}${markerDefs}${body}${label}</svg>`;
}

/**
 * Draw a shape into its box at the given size.
 *
 * A rotated shape's anchor is its bounding box AFTER the turn, so the shape itself is drawn in a
 * box with the two axes swapped for a quarter turn and rotated back into the anchor. Without this
 * a brace turned on its side to bracket a run of columns was drawn upright and tall, across the
 * rows it was meant to sit under.
 */
function paintShape(box: HTMLElement, sh: SheetShape, w: number, h: number): void {
  const rot = sh.rotation ?? 0;
  if (!rot) {
    box.innerHTML = shapeSvg(sh, w, h);
    return;
  }
  const quarter = Math.abs((((rot % 180) + 180) % 180) - 90) < 1;
  const [dw, dh] = quarter ? [h, w] : [w, h];
  box.innerHTML = shapeSvg(sh, dw, dh);
  const svg = box.firstElementChild as HTMLElement | null;
  if (!svg) return;
  svg.style.position = "absolute";
  svg.style.left = "50%";
  svg.style.top = "50%";
  svg.style.width = `${dw}px`;
  svg.style.height = `${dh}px`;
  svg.style.transform = `translate(-50%, -50%) rotate(${rot}deg)`;
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
      // Only the primary button drags. A right-click has its own job (the context menu), and
      // starting a move on it leaves the shape following the pointer with no button held down.
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      select(sh);
      const sx = e.clientX, sy = e.clientY;
      const x0 = parseFloat(box.style.left) || 0, y0 = parseFloat(box.style.top) || 0;
      const w0 = box.offsetWidth, h0 = box.offsetHeight;
      const onMove = (ev: PointerEvent): void => {
        const dx = ev.clientX - sx, dy = ev.clientY - sy;
        if (mode === "move") { box.style.left = `${Math.max(0, x0 + dx)}px`; box.style.top = `${Math.max(0, y0 + dy)}px`; }
        else {
          const w = Math.max(8, w0 + dx), h = Math.max(8, h0 + dy);
          box.style.width = `${w}px`; box.style.height = `${h}px`;
          paintShape(box, sh, w, h); // redraw so geometry tracks the new size
        }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        const x = parseFloat(box.style.left) || 0, y = parseFloat(box.style.top) || 0;
        // A plain click (no drag) must not commit or refresh, or a double-click's box is replaced
        // between its two clicks and never fires.
        if (x === x0 && y === y0 && box.offsetWidth === w0 && box.offsetHeight === h0) {
          // Not a drag but a click: a shape with a macro is a button, so run it.
          if (sh.macro && deps.runMacro) deps.runMacro(sh.macro);
          return;
        }
        rectToAnchor(sh, x, y, box.offsetWidth, box.offsetHeight);
        deps.onEdit?.(sh);
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    // Dragging from the text of a shape that is already selected selects the text instead, the
    // way clicking into a shape twice does elsewhere: the first click picks the shape up, the
    // second one gets at what it says. Otherwise a text box could never be read out of.
    const onText = (e: PointerEvent): boolean =>
      sh === selected && e.target instanceof Element && !!e.target.closest(".sheetedit-shapetext");
    box.addEventListener("pointerdown", (e) => { if (e.target !== handle && !onText(e)) start(e, "move"); });
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
      paintShape(box, sh, w, h);
      if (sh.macro && deps.runMacro) {
        box.classList.add("macro");
        box.title = `${deps.macroTitle ?? ""} ${sh.macro}`.trim();
      }
      if (editable) {
        const handle = document.createElement("div");
        handle.className = "sheetedit-shape-resize";
        box.appendChild(handle);
        attachDrag(box, handle, sh);
        if (!sh.macro || !deps.runMacro) box.title = deps.onActivate ? "Double-click to edit" : "";
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
      if (!editable && sh.macro && deps.runMacro) {
        box.addEventListener("click", (e) => { e.preventDefault(); e.stopPropagation(); deps.runMacro!(sh.macro!); });
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
