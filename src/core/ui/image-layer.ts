import type { Sheet } from "../model";
import type { ChartGeom } from "./chart-overlay";

// A read-only overlay that floats a sheet's embedded pictures over the grid, anchored to cells and
// glued while scrolling (the same layer pattern the chart overlay uses, minus interaction). Images
// are preserved on save; this only renders them.

export interface ImageLayerDeps {
  wrap: HTMLElement;
  gridScroll: HTMLElement;
  getSheet: () => Sheet | undefined;
  geom: () => ChartGeom;
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
    .sheetedit-imagebox img { width:100%; height:100%; object-fit:contain; display:block; }
  `;
  document.head.appendChild(s);
}

export function setupImageLayer(deps: ImageLayerDeps): { refresh(): void; teardown(): void } {
  injectStyles();
  const { wrap, gridScroll } = deps;
  const layer = document.createElement("div");
  layer.className = "sheetedit-imagelayer";
  const inner = document.createElement("div");
  inner.className = "sheetedit-imagelayer-inner";
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
    const images = sheet?.images ?? [];
    layer.style.display = images.length ? "block" : "none";
    if (!images.length) return;
    const g = deps.geom();
    for (const im of images) {
      const a = im.anchor;
      const x = g.xOfCol(a.fromCol) + a.fromColOff;
      const y = g.yOfRow(a.fromRow) + a.fromRowOff;
      const x2 = g.xOfCol(a.toCol) + a.toColOff;
      const y2 = g.yOfRow(a.toRow) + a.toRowOff;
      const box = document.createElement("div");
      box.className = "sheetedit-imagebox";
      box.style.left = `${x}px`;
      box.style.top = `${y}px`;
      box.style.width = `${Math.max(1, x2 - x)}px`;
      box.style.height = `${Math.max(1, y2 - y)}px`;
      const img = document.createElement("img");
      img.src = im.dataUri;
      img.alt = "";
      box.appendChild(img);
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
