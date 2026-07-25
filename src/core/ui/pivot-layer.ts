import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet } from "../model";
import type { ChartGeom } from "./chart-overlay";

// A read-only overlay that outlines each pivot table / data-pilot output range on the grid and
// labels it, so a user can tell the region is a live pivot (not hand-editable cells) and that
// editing its source refreshes it. Pivots themselves render as the cells they materialise; this
// only draws the frame. Same glued-to-the-grid layer pattern as the image overlay.

export interface PivotLayerDeps {
  wrap: HTMLElement;
  /** The grid's viewports: the element plus whether it draws the header / row numbers. */
  panes: () => { el: HTMLElement; header: boolean; rowHeader: boolean }[];
  getSheet: () => Sheet | undefined;
  geom: () => ChartGeom;
  label: (name: string) => string; // localised tooltip/label text
  /** Clicking a pivot's tag opens its actions (refresh / edit) at the given viewport point. */
  onTag?: (pivot: NonNullable<Sheet["pivotTables"]>[number], x: number, y: number) => void;
}

const STYLE_ID = "sheetedit-pivot-style";
function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-pivotlayer { position:absolute; overflow:hidden; pointer-events:none; z-index:4; }
    .sheetedit-pivotlayer-inner { position:absolute; inset:0; }
    .sheetedit-pivotbox { position:absolute; box-sizing:border-box; border:1.5px dashed var(--sheetedit-accent,#3b82f6); border-radius:3px; background:color-mix(in srgb, var(--sheetedit-accent,#3b82f6) 6%, transparent); }
    .sheetedit-pivottag { position:absolute; top:0; left:0; transform:translateY(-100%); pointer-events:auto; font:600 10px/1.4 system-ui,sans-serif; color:#fff; background:var(--sheetedit-accent,#3b82f6); padding:1px 6px; border-radius:3px 3px 0 0; white-space:nowrap; cursor:default; }
  `;
  document.head.appendChild(s);
}

export function setupPivotLayer(deps: PivotLayerDeps): { refresh(): void; teardown(): void } {
  injectStyles();
  const hosts = setupOverlayHosts({
    wrap: deps.wrap,
    panes: deps.panes,
    geom: () => { const g = deps.geom(); return { rnW: g.rnW, headerH: g.headerH, yOfRow: g.yOfRow }; },
    className: "sheetedit-pivotlayer",
    innerClassName: "sheetedit-pivotlayer-inner",
  });

  const refresh = (): void => {
    hosts.clear();
    const sheet = deps.getSheet();
    const pivots = (sheet?.pivotTables ?? []).filter((p) => p.targetRange);
    hosts.setVisible(pivots.length > 0);
    if (!pivots.length) return;
    const g = deps.geom();
    for (const p of pivots) {
      const t = p.targetRange!;
      const x = g.xOfCol(t.c1);
      const y = g.yOfRow(t.r1);
      const x2 = g.xOfCol(t.c2 + 1);
      const y2 = g.yOfRow(t.r2 + 1);
      const box = document.createElement("div");
      box.className = "sheetedit-pivotbox";
      box.style.left = `${x}px`;
      box.style.top = `${y}px`;
      box.style.width = `${Math.max(1, x2 - x)}px`;
      box.style.height = `${Math.max(1, y2 - y)}px`;
      const tag = document.createElement("div");
      tag.className = "sheetedit-pivottag";
      // Normally the tag sits just above the frame; when the pivot starts at the top row the frame
      // is flush with the layer edge and an above-tag would be clipped, so tuck it inside instead.
      if (y < 16) { tag.style.transform = "none"; tag.style.borderRadius = "0 3px 3px 0"; }
      tag.textContent = p.name;
      const parts: string[] = [];
      if (p.rowFields.length) parts.push(`Rows: ${p.rowFields.join(", ")}`);
      if (p.colFields.length) parts.push(`Columns: ${p.colFields.join(", ")}`);
      if (p.dataFields.length) parts.push(`Values: ${p.dataFields.map((d) => d.name).join(", ")}`);
      if (p.sourceSheet) parts.push(`Source: ${p.sourceSheet}`);
      tag.title = `${deps.label(p.name)}\n${parts.join("\n")}`;
      if (deps.onTag) { tag.style.cursor = "pointer"; tag.addEventListener("click", (e) => { e.stopPropagation(); deps.onTag!(p, (e as MouseEvent).clientX, (e as MouseEvent).clientY); }); }
      box.appendChild(tag);
      hosts.hostFor(t.r1).appendChild(box);
    }
    hosts.layout();
  };

  return { refresh, teardown() { hosts.teardown(); } };
}
