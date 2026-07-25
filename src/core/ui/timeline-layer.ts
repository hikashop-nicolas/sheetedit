import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet, SheetTimeline } from "../model";
import type { ChartGeom } from "./chart-overlay";

// The timeline overlay: a date-range filter panel floated over the grid. It shows the selected
// range and the periods the data spans at the timeline's granularity (year / quarter / month /
// day); clicking a period selects it, shift-clicking extends the range, and the clear button
// restores the full bounds. The host re-filters the linked pivots.

export interface TimelineLayerDeps {
  wrap: HTMLElement;
  /** The grid's scroll containers, top-most first (two when a row split is on). */
  panes: () => HTMLElement[];
  getSheet: () => Sheet | undefined;
  geom: () => ChartGeom;
  onChange?: (t: SheetTimeline) => void;
}

const STYLE_ID = "sheetedit-timeline-style";
function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-timelinelayer { position:absolute; overflow:hidden; pointer-events:none; z-index:6; }
    .sheetedit-timelinelayer-inner { position:absolute; inset:0; }
    .sheetedit-timelinebox { position:absolute; pointer-events:auto; display:flex; flex-direction:column;
      background:var(--sheetedit-chrome,#fff); color:var(--sheetedit-text,#1c1f24);
      border:1px solid var(--sheetedit-border,#c8ccd2); border-radius:6px;
      box-shadow:0 2px 10px rgba(0,0,0,.18); font:12px system-ui,sans-serif; overflow:hidden; }
    .sheetedit-timeline-head { display:flex; align-items:baseline; gap:6px; padding:5px 7px;
      border-bottom:1px solid var(--sheetedit-border,#c8ccd2); background:rgba(127,127,127,.08); }
    .sheetedit-timeline-title { font-weight:600; flex:0 0 auto; }
    .sheetedit-timeline-range { flex:1; color:var(--sheetedit-muted,#6b7280); overflow:hidden;
      text-overflow:ellipsis; white-space:nowrap; font-size:11px; }
    .sheetedit-timeline-clear { border:0; background:none; cursor:pointer; color:inherit; opacity:.65;
      font:inherit; padding:1px 4px; border-radius:4px; }
    .sheetedit-timeline-clear:hover { opacity:1; background:rgba(127,127,127,.18); }
    .sheetedit-timeline-periods { flex:1; display:flex; align-items:stretch; gap:2px; padding:5px; overflow-x:auto; }
    .sheetedit-timeline-period { flex:1 0 34px; border:1px solid var(--sheetedit-border,#c8ccd2);
      border-radius:3px; background:transparent; color:inherit; font:inherit; font-size:11px;
      cursor:pointer; padding:2px; opacity:.45; white-space:nowrap; }
    .sheetedit-timeline-period.on { opacity:1; background:var(--sheetedit-accent,#4c8bf5); color:#fff;
      border-color:var(--sheetedit-accent,#4c8bf5); }
  `;
  document.head.appendChild(s);
}

/** ISO date -> [y, m, d] (the parts we need; Excel writes "2024-01-01T00:00:00"). */
function parts(iso: string | undefined): { y: number; m: number; d: number } | null {
  if (!iso) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso.trim());
  return m ? { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) } : null;
}
const pad = (n: number): string => String(n).padStart(2, "0");
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/** The periods a timeline spans at its granularity, each with its own [start, end) ISO range. */
export function timelinePeriods(tl: SheetTimeline): { label: string; start: string; end: string }[] {
  const a = parts(tl.boundStart) ?? parts(tl.startDate);
  const b = parts(tl.boundEnd) ?? parts(tl.endDate);
  if (!a || !b) return [];
  const level = String(tl.level ?? "2");
  const out: { label: string; start: string; end: string }[] = [];
  const iso = (y: number, m: number, d: number): string => `${y}-${pad(m)}-${pad(d)}T00:00:00`;
  if (level === "0") { // years
    for (let y = a.y; y <= b.y && out.length < 200; y++) out.push({ label: String(y), start: iso(y, 1, 1), end: iso(y + 1, 1, 1) });
  } else if (level === "1") { // quarters
    for (let y = a.y; y <= b.y && out.length < 200; y++)
      for (let q = 0; q < 4; q++) {
        const sm = q * 3 + 1;
        if (y === a.y && sm + 2 < a.m) continue;
        if (y === b.y && sm > b.m) break;
        out.push({ label: `Q${q + 1} ${y}`, start: iso(y, sm, 1), end: sm + 3 > 12 ? iso(y + 1, 1, 1) : iso(y, sm + 3, 1) });
      }
  } else if (level === "3") { // days
    let y = a.y, m = a.m, d = a.d;
    while (out.length < 400 && (y < b.y || (y === b.y && (m < b.m || (m === b.m && d <= b.d))))) {
      const next = new Date(Date.UTC(y, m - 1, d + 1));
      out.push({ label: `${d}/${m}`, start: iso(y, m, d), end: iso(next.getUTCFullYear(), next.getUTCMonth() + 1, next.getUTCDate()) });
      y = next.getUTCFullYear(); m = next.getUTCMonth() + 1; d = next.getUTCDate();
    }
  } else { // months (default)
    let y = a.y, m = a.m;
    while (out.length < 240 && (y < b.y || (y === b.y && m <= b.m))) {
      const ny = m === 12 ? y + 1 : y, nm = m === 12 ? 1 : m + 1;
      out.push({ label: `${MONTHS[m - 1]} ${String(y).slice(2)}`, start: iso(y, m, 1), end: iso(ny, nm, 1) });
      y = ny; m = nm;
    }
  }
  return out;
}

export function setupTimelineLayer(deps: TimelineLayerDeps): { refresh(): void; teardown(): void } {
  injectStyles();
  const hosts = setupOverlayHosts({
    wrap: deps.wrap,
    panes: deps.panes,
    geom: () => { const g = deps.geom(); return { rnW: g.rnW, headerH: g.headerH, yOfRow: g.yOfRow }; },
    className: "sheetedit-timelinelayer",
    innerClassName: "sheetedit-timelinelayer-inner",
  });

  const refresh = (): void => {
    hosts.clear();
    const sheet = deps.getSheet();
    const tls = sheet?.timelines ?? [];
    hosts.setVisible(tls.length > 0);
    if (!tls.length) return;
    const g = deps.geom();
    for (const tl of tls) {
      const a = tl.anchor;
      const x = a ? g.xOfCol(a.fromCol) + a.fromColOff : 20;
      const y = a ? g.yOfRow(a.fromRow) + a.fromRowOff : 20;
      const w = a ? Math.max(220, g.xOfCol(a.toCol) + a.toColOff - x) : 320;
      const h = a ? Math.max(70, g.yOfRow(a.toRow) + a.toRowOff - y) : 90;
      const box = document.createElement("div");
      box.className = "sheetedit-timelinebox";
      box.style.cssText += `left:${x}px;top:${y}px;width:${w}px;height:${h}px`;
      box.dataset.timeline = tl.name;

      const head = document.createElement("div");
      head.className = "sheetedit-timeline-head";
      const title = document.createElement("span");
      title.className = "sheetedit-timeline-title";
      title.textContent = tl.caption || tl.sourceName || tl.name;
      const range = document.createElement("span");
      range.className = "sheetedit-timeline-range";
      const ps = parts(tl.startDate), pe = parts(tl.endDate);
      range.textContent = ps && pe ? `${ps.y}-${pad(ps.m)}-${pad(ps.d)} → ${pe.y}-${pad(pe.m)}-${pad(pe.d)}` : "All periods";
      const clear = document.createElement("button");
      clear.type = "button"; clear.className = "sheetedit-timeline-clear"; clear.textContent = "⊗"; clear.title = "Clear filter";
      clear.addEventListener("click", (e) => {
        e.stopPropagation();
        tl.startDate = undefined; tl.endDate = undefined; tl.dirty = true;
        deps.onChange?.(tl);
      });
      head.append(title, range, clear);
      box.appendChild(head);

      const strip = document.createElement("div");
      strip.className = "sheetedit-timeline-periods";
      const periods = timelinePeriods(tl);
      const inSel = (p: { start: string; end: string }): boolean => {
        if (!tl.startDate || !tl.endDate) return true; // no filter = every period lit
        return p.start >= tl.startDate && p.start < tl.endDate;
      };
      periods.forEach((p, i) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-timeline-period" + (inSel(p) ? " on" : "");
        b.textContent = p.label;
        b.title = p.label;
        b.dataset.period = String(i);
        b.addEventListener("click", (e) => {
          if (e.shiftKey && tl.startDate) {
            // Extend the range to cover both ends.
            tl.startDate = p.start < tl.startDate ? p.start : tl.startDate;
            tl.endDate = tl.endDate && p.end > tl.endDate ? p.end : (tl.endDate ?? p.end);
          } else if (tl.startDate === p.start && tl.endDate === p.end) {
            tl.startDate = undefined; tl.endDate = undefined; // clicking the only period clears
          } else { tl.startDate = p.start; tl.endDate = p.end; }
          tl.dirty = true;
          deps.onChange?.(tl);
        });
        strip.appendChild(b);
      });
      box.appendChild(strip);
      hosts.hostFor(tl.anchor?.fromRow ?? 1).appendChild(box);
    }
    hosts.layout();
  };

  return { refresh, teardown() { hosts.teardown(); } };
}
