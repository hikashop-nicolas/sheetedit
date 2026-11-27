// Shared placement for the toolbar menus and popovers: keep them on the part of the screen the
// user can actually see (visual viewport, on-screen keyboard, safe areas, below the editor's top).

export interface Box { left: number; top: number; right: number; bottom: number }

/** The safe-area insets, read off a probe since env() is only visible to CSS. */
function safeInsets(host: HTMLElement): Box {
  const probe = document.createElement("div");
  probe.style.cssText = "position:fixed;visibility:hidden;pointer-events:none;padding:env(safe-area-inset-top) env(safe-area-inset-right) env(safe-area-inset-bottom) env(safe-area-inset-left)";
  host.appendChild(probe);
  const cs = getComputedStyle(probe);
  const px = (v: string): number => parseFloat(v) || 0;
  const out = { top: px(cs.paddingTop), right: px(cs.paddingRight), bottom: px(cs.paddingBottom), left: px(cs.paddingLeft) };
  probe.remove();
  return out;
}

/** The visible area in client coordinates. `floor` (the editor) keeps popups below its top edge,
    which is where a host's own header ends. */
export function visibleBounds(floor?: HTMLElement): Box {
  const vv = window.visualViewport;
  const vw = document.documentElement.clientWidth || window.innerWidth;
  const vh = window.innerHeight;
  const b: Box = vv
    ? { left: vv.offsetLeft, top: vv.offsetTop, right: vv.offsetLeft + vv.width, bottom: vv.offsetTop + vv.height }
    : { left: 0, top: 0, right: vw, bottom: vh };
  const s = safeInsets(floor ?? document.body);
  b.left = Math.max(b.left, s.left);
  b.top = Math.max(b.top, s.top);
  b.right = Math.min(b.right, vw - s.right);
  b.bottom = Math.min(b.bottom, vh - s.bottom);
  if (floor) {
    const f = floor.getBoundingClientRect();
    b.top = Math.max(b.top, f.top);
    b.left = Math.max(b.left, f.left);
    b.right = Math.min(b.right, f.right);
  }
  return b;
}

/** A usable anchor rectangle: the element's own, else a handed-over one (a folded button). */
export function anchorRect(el: HTMLElement, fallback?: Box): Box | null {
  const r = el.getBoundingClientRect();
  if (r.width > 0 || r.height > 0) return r;
  return fallback ?? null;
}

export interface PlaceOptions {
  /** The positioned parent for an absolute popup; omitted for a fixed one. */
  container?: HTMLElement;
  /** Popups stay below this element's top (the editor). */
  floor?: HTMLElement;
  gap?: number;
  /** Space kept free on the right, as the toolbar always did. */
  marginRight?: number;
}

/**
 * Put `el` under `anchor`, inside the visible area: shifted left to fit, and capped in height so it
 * scrolls inside itself rather than running off screen. `el` must already be shown.
 */
export function placePopup(el: HTMLElement, anchor: Box, opts: PlaceOptions = {}): void {
  const gap = opts.gap ?? 2;
  const b = visibleBounds(opts.floor);
  el.style.maxHeight = "";
  const origin = opts.container?.getBoundingClientRect() ?? { left: 0, top: 0 };
  const w = el.offsetWidth;
  let top = Math.max(anchor.bottom + gap, b.top);
  let room = b.bottom - top - 4;
  // Too little room below (a keyboard is up): open above the anchor when that has more.
  const above = anchor.top - gap - b.top;
  if (el.offsetHeight > room && above > room + 40) {
    room = above - 4;
    top = Math.max(b.top, anchor.top - gap - Math.min(el.offsetHeight, room));
  }
  if (el.offsetHeight > room) {
    el.style.maxHeight = `${Math.max(80, Math.floor(room))}px`;
    el.style.overflowY = "auto";
  }
  const left = Math.max(b.left, Math.min(anchor.left, b.right - w - (opts.marginRight ?? 0)));
  el.style.top = `${top - origin.top}px`;
  el.style.left = `${left - origin.left}px`;
}

/** Open a list scrolled to its checked entry, else to the top, without scrolling the page. */
export function scrollToChecked(el: HTMLElement): void {
  el.scrollTop = 0;
  const cur = el.querySelector<HTMLElement>('[aria-checked="true"]');
  if (!cur || el.scrollHeight <= el.clientHeight) return;
  const er = el.getBoundingClientRect();
  const cr = cur.getBoundingClientRect();
  const below = cr.bottom - er.bottom;
  if (below > 0) el.scrollTop = Math.min(el.scrollHeight - el.clientHeight, below + el.clientHeight / 2 - cr.height / 2);
}
