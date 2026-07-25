import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet, SheetImage } from "../model";
import type { ChartGeom } from "./chart-overlay";

// An overlay that floats a sheet's embedded pictures over the grid, anchored to cells and glued
// while scrolling (the same layer pattern the chart overlay uses). xlsx images can be moved (drag
// the box) and resized (drag the corner handle); the new cell anchor is committed back to the model
// and the host persists it. ods images are render-only (no drawing write-back yet).

export interface ImageLayerDeps {
  wrap: HTMLElement;
  /** The grid's viewports: the element plus whether it draws the header / row numbers. */
  panes: () => { el: HTMLElement; header: boolean; rowHeader: boolean }[];
  getSheet: () => Sheet | undefined;
  geom: () => ChartGeom;
  /** After a move/resize: the image's anchor was updated + dirty set; the host marks + persists. */
  onEdit?: (im: SheetImage) => void;
  /** Double-click on an editable image: the host prompts for a replacement file. */
  onReplace?: (im: SheetImage) => void;
  /** True when the active sheet's images can be written back; gates the drag handles + replace. */
  editable?: () => boolean;
}

const STYLE_ID = "sheetedit-image-style";
function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-imagelayer { position:absolute; overflow:hidden; pointer-events:none; z-index:5; }
    .sheetedit-imagelayer-inner { position:absolute; inset:0; }
    .sheetedit-imagebox { position:absolute; }
    .sheetedit-imagebox img { width:100%; height:100%; object-fit:contain; display:block; pointer-events:none; }
    .sheetedit-imagebox.editable { pointer-events:auto; cursor:move; }
    .sheetedit-imagebox.selected { outline:1.5px solid var(--sheetedit-accent,#4c8bf5); outline-offset:1px; }
    .sheetedit-image-resize { position:absolute; right:-5px; bottom:-5px; width:12px; height:12px; border-radius:3px;
      background:var(--sheetedit-accent,#4c8bf5); border:1.5px solid #fff; cursor:nwse-resize; pointer-events:auto; display:none; }
    .sheetedit-imagebox.selected .sheetedit-image-resize { display:block; }
  `;
  document.head.appendChild(s);
}

export function setupImageLayer(deps: ImageLayerDeps): { refresh(): void; teardown(): void } {
  injectStyles();
  const hosts = setupOverlayHosts({
    wrap: deps.wrap,
    panes: deps.panes,
    geom: () => { const g = deps.geom(); return { rnW: g.rnW, headerH: g.headerH, yOfRow: g.yOfRow }; },
    className: "sheetedit-imagelayer",
    innerClassName: "sheetedit-imagelayer-inner",
  });

  let selected: SheetImage | null = null;


  // Commit a dragged/resized pixel rect (content coords) back to the image's cell anchor. For ods,
  // also record the frame's pixel offset from its (unchanged) anchor cell + its size, which the ods
  // writer applies without re-parenting the frame.
  const rectToAnchor = (im: SheetImage, x: number, y: number, w: number, h: number): void => {
    const g = deps.geom();
    const set = (px: number, at: (p: number) => number, of: (i: number) => number): [number, number] => { const i = Math.max(1, at(px)); return [i, Math.max(0, px - of(i))]; };
    const [fc, fco] = set(x, g.colAt, g.xOfCol);
    const [fr, fro] = set(y, g.rowAt, g.yOfRow);
    const [tc, tco] = set(x + w, g.colAt, g.xOfCol);
    const [tr, tro] = set(y + h, g.rowAt, g.yOfRow);
    im.anchor = { fromCol: fc, fromRow: fr, fromColOff: fco, fromRowOff: fro, toCol: tc, toRow: tr, toColOff: tco, toRowOff: tro };
    if (im.odsFrameEl && im.odsAnchorCol != null && im.odsAnchorRow != null) {
      im.odsFrame = { x: x - g.xOfCol(im.odsAnchorCol), y: y - g.yOfRow(im.odsAnchorRow), w, h };
    }
    im.dirty = true;
  };

  const attachDrag = (box: HTMLElement, handle: HTMLElement, im: SheetImage): void => {
    const start = (e: PointerEvent, mode: "move" | "resize"): void => {
      e.preventDefault();
      e.stopPropagation();
      select(im);
      const sx = e.clientX, sy = e.clientY;
      const x0 = parseFloat(box.style.left) || 0, y0 = parseFloat(box.style.top) || 0;
      const w0 = box.offsetWidth, h0 = box.offsetHeight;
      const onMove = (ev: PointerEvent): void => {
        const dx = ev.clientX - sx, dy = ev.clientY - sy;
        if (mode === "move") { box.style.left = `${Math.max(0, x0 + dx)}px`; box.style.top = `${Math.max(0, y0 + dy)}px`; }
        else { box.style.width = `${Math.max(12, w0 + dx)}px`; box.style.height = `${Math.max(12, h0 + dy)}px`; }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        const x = parseFloat(box.style.left) || 0, y = parseFloat(box.style.top) || 0;
        if (x === x0 && y === y0 && box.offsetWidth === w0 && box.offsetHeight === h0) return; // a plain click, not a drag
        rectToAnchor(im, x, y, box.offsetWidth, box.offsetHeight);
        deps.onEdit?.(im);
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    box.addEventListener("pointerdown", (e) => { if (e.target !== handle) start(e, "move"); });
    handle.addEventListener("pointerdown", (e) => start(e, "resize"));
  };

  const boxes = new Map<SheetImage, HTMLElement>();
  const select = (im: SheetImage | null): void => {
    selected = im;
    for (const [i, b] of boxes) b.classList.toggle("selected", i === im);
  };

  const refresh = (): void => {
    hosts.clear();
    boxes.clear();
    const sheet = deps.getSheet();
    const images = sheet?.images ?? [];
    hosts.setVisible(images.length > 0);
    if (!images.length) { selected = null; return; }
    const editable = deps.editable?.() ?? false;
    const g = deps.geom();
    for (const im of images) {
      const a = im.anchor;
      const x = g.xOfCol(a.fromCol) + a.fromColOff;
      const y = g.yOfRow(a.fromRow) + a.fromRowOff;
      const x2 = g.xOfCol(a.toCol) + a.toColOff;
      const y2 = g.yOfRow(a.toRow) + a.toRowOff;
      const box = document.createElement("div");
      box.className = "sheetedit-imagebox" + (editable ? " editable" : "") + (im === selected ? " selected" : "");
      box.style.left = `${x}px`;
      box.style.top = `${y}px`;
      box.style.width = `${Math.max(1, x2 - x)}px`;
      box.style.height = `${Math.max(1, y2 - y)}px`;
      const img = document.createElement("img");
      img.src = im.dataUri;
      img.alt = "";
      box.appendChild(img);
      if (editable) {
        const handle = document.createElement("div");
        handle.className = "sheetedit-image-resize";
        box.appendChild(handle);
        attachDrag(box, handle, im);
        box.title = deps.onReplace ? "Double-click to replace" : "";
        box.addEventListener("dblclick", (e) => { e.preventDefault(); e.stopPropagation(); deps.onReplace?.(im); });
      }
      hosts.hostFor(im.anchor.fromRow).appendChild(box);
      boxes.set(im, box);
    }
    hosts.layout();
  };


  // Tap on empty grid deselects.
  const onGridDown = (): void => { if (selected) select(null); };
  hosts.onPanes("pointerdown", onGridDown as (e: Event) => void);
  return {
    refresh,
    teardown() { hosts.teardown(); },
  };
}
