import type { ShapeGeom } from "./model";

// Pure geometry for the polygon-based shapes, shared by the SVG overlay (rendering) and the ODF
// writer (enhanced-path). rect / roundRect / ellipse / line have their own primitives and are not
// listed here.

/** Vertices (in a 0..w by 0..h box) for a polygon shape, or null for a primitive (rect/ellipse/line). */
export function shapePoints(geom: ShapeGeom, w: number, h: number): [number, number][] | null {
  const poly = (pts: [number, number][]): [number, number][] => pts;
  switch (geom) {
    case "diamond":
      return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case "triangle":
      return poly([[w / 2, 0], [w, h], [0, h]]);
    case "parallelogram":
      return poly([[w * 0.25, 0], [w, 0], [w * 0.75, h], [0, h]]);
    case "hexagon":
      return poly([[w * 0.25, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [w * 0.25, h], [0, h / 2]]);
    case "pentagon":
      return regular(5, w, h);
    case "star":
      return star5(w, h);
    case "rightArrow":
      return arrowRight(w, h);
    // The other three directions are the same arrow turned. They used to share the right-pointing
    // one outright, so a left arrow pointed right: worse than a fallback, because it looks meant.
    case "leftArrow":
      return flipX(arrowRight(w, h), w);
    case "downArrow":
      return transpose(arrowRight(h, w));
    case "upArrow":
      return flipY(transpose(arrowRight(h, w)), h);
    case "leftRightArrow":
      return poly([[0, h / 2], [w * 0.25, 0], [w * 0.25, h * 0.3], [w * 0.75, h * 0.3], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [w * 0.75, h * 0.7], [w * 0.25, h * 0.7], [w * 0.25, h]]);
    case "upDownArrow":
      return transpose(shapePoints("leftRightArrow", h, w) ?? []);
    case "notchedArrow": // a right arrow with a V cut out of its tail
      return poly([[0, h * 0.3], [w * 0.6, h * 0.3], [w * 0.6, 0], [w, h / 2], [w * 0.6, h], [w * 0.6, h * 0.7], [0, h * 0.7], [w * 0.12, h / 2]]);
    case "bentArrow": // out to the right, then up, with the head on top
      return poly([[0, h], [0, h * 0.65], [w * 0.68, h * 0.65], [w * 0.68, h * 0.3], [w * 0.5, h * 0.3], [w * 0.79, 0], [w, h * 0.3], [w * 0.86, h * 0.3], [w * 0.86, h]]);
    case "quadArrow":
      return quadArrow(w, h);
    case "chevron":
      return poly([[0, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [0, h], [w * 0.25, h / 2]]);
    case "homePlate":
      return poly([[0, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [0, h]]);
    case "trapezoid":
      return poly([[w * 0.25, 0], [w * 0.75, 0], [w, h], [0, h]]);
    case "trapezoidDown":
      return poly([[0, 0], [w, 0], [w * 0.75, h], [w * 0.25, h]]);
    case "triangleDown":
      return poly([[0, 0], [w, 0], [w / 2, h]]);
    case "plus":
      return poly([[w * 0.35, 0], [w * 0.65, 0], [w * 0.65, h * 0.35], [w, h * 0.35], [w, h * 0.65], [w * 0.65, h * 0.65], [w * 0.65, h], [w * 0.35, h], [w * 0.35, h * 0.65], [0, h * 0.65], [0, h * 0.35], [w * 0.35, h * 0.35]]);
    case "corner":
      return poly([[0, 0], [w * 0.4, 0], [w * 0.4, h * 0.6], [w, h * 0.6], [w, h], [0, h]]);
    case "diagStripe":
      return poly([[0, h * 0.5], [w * 0.5, 0], [w, 0], [0, h]]);
    case "manualInput": // a rectangle with a sloping top edge
      return poly([[0, h * 0.25], [w, 0], [w, h], [0, h]]);
    case "snipRect": // one corner cut off square
      return poly([[0, 0], [w * 0.8, 0], [w, h * 0.2], [w, h], [0, h]]);
    case "heptagon":
      return regular(7, w, h);
    case "octagon":
      return regular(8, w, h);
    case "decagon":
      return regular(10, w, h);
    case "dodecagon":
      return regular(12, w, h);
    default:
      return null;
  }
}

/** A right-pointing arrow: a shaft over the left, a head filling the right. */
function arrowRight(w: number, h: number): [number, number][] {
  return [[0, h * 0.3], [w * 0.6, h * 0.3], [w * 0.6, 0], [w, h / 2], [w * 0.6, h], [w * 0.6, h * 0.7], [0, h * 0.7]];
}

/** One arrow out of each side of a central square: four tips, four heads, a cross of shaft. */
function quadArrow(w: number, h: number): [number, number][] {
  const X = (f: number): number => round(w * f), Y = (f: number): number => round(h * f);
  // a/b bound the shaft, c/d the head's outer edge, e/f where the head meets the shaft.
  const [a, b, c, d, e, f] = [0.4, 0.6, 0.3, 0.7, 0.25, 0.75];
  return [
    [X(0.5), 0], [X(d), Y(e)], [X(b), Y(e)], [X(b), Y(a)], [X(f), Y(a)], [X(f), Y(c)],
    [w, Y(0.5)], [X(f), Y(d)], [X(f), Y(b)], [X(b), Y(b)], [X(b), Y(f)], [X(d), Y(f)],
    [X(0.5), h], [X(c), Y(f)], [X(a), Y(f)], [X(a), Y(b)], [X(e), Y(b)], [X(e), Y(d)],
    [0, Y(0.5)], [X(e), Y(c)], [X(e), Y(a)], [X(a), Y(a)], [X(a), Y(e)], [X(c), Y(e)],
  ];
}

const flipX = (pts: [number, number][], w: number): [number, number][] => pts.map(([x, y]) => [round(w - x), y]);
const flipY = (pts: [number, number][], h: number): [number, number][] => pts.map(([x, y]) => [x, round(h - y)]);
/** Swap the axes: turns a shape built across into the same shape built down. */
const transpose = (pts: [number, number][]): [number, number][] => pts.map(([x, y]) => [y, x]);

// Braces and brackets are open outlines, not closed polygons: a stroked path with no fill, whose
// ends stop where they stop. A shape that groups a run of rows says nothing if it is drawn as the
// rectangle around them, which is what an unrecognised preset falls back to.

/** One "}" (or, mirrored, "{") spanning the full height of a box `w` wide starting at `x`. */
function brace(x: number, w: number, h: number, left: boolean): string {
  // The spine runs down the middle of the band; the tips curve out one way and the point the
  // other. `r` is how far the curves reach, capped so a short brace stays a brace.
  const r = Math.min(w, h / 4);
  const mid = x + w / 2;
  const tip = left ? x + w : x; // where the two ends land
  const point = left ? x : x + w; // where the middle pokes out
  const q = (cx: number, cy: number, ex: number, ey: number): string => `Q ${round(cx)} ${round(cy)} ${round(ex)} ${round(ey)}`;
  return [
    `M ${round(tip)} 0`,
    q(mid, 0, mid, r),
    `L ${round(mid)} ${round(h / 2 - r)}`,
    q(mid, h / 2, point, h / 2),
    q(mid, h / 2, mid, h / 2 + r),
    `L ${round(mid)} ${round(h - r)}`,
    q(mid, h, tip, h),
  ].join(" ");
}

/** One "]" (or, mirrored, "[") spanning the full height of a box `w` wide starting at `x`. */
function bracket(x: number, w: number, h: number, left: boolean): string {
  const tip = left ? x + w : x;
  const spine = left ? x : x + w;
  return `M ${round(tip)} 0 L ${round(spine)} 0 L ${round(spine)} ${round(h)} L ${round(tip)} ${round(h)}`;
}

/**
 * An elbow connector: out from the start, a right-angled turn, across, and back in to the end.
 * `adjust` is where the turn falls along the run, as the file states it (it may sit outside 0..1,
 * which is how Excel draws a connector that doubles back).
 */
function elbow(w: number, h: number, adjust: number): string {
  const at = round(w * adjust);
  return `M 0 0 L ${at} 0 L ${at} ${round(h)} L ${round(w)} ${round(h)}`;
}

// Shapes that need curves, or a hole, or a tail hanging outside their own box: a closed path
// rather than a polygon. A callout drawn as a plain rectangle is the clearest case - it loses the
// tail, and with it the whole point, which is to say WHICH cell the note is about.

/** A rounded-rectangle path, `r` px at the corners. */
function roundRectPath(x: number, y: number, w: number, h: number, r: number): string {
  const rr = Math.min(r, w / 2, h / 2);
  return `M ${round(x + rr)} ${round(y)} H ${round(x + w - rr)} A ${round(rr)} ${round(rr)} 0 0 1 ${round(x + w)} ${round(y + rr)}` +
    ` V ${round(y + h - rr)} A ${round(rr)} ${round(rr)} 0 0 1 ${round(x + w - rr)} ${round(y + h)}` +
    ` H ${round(x + rr)} A ${round(rr)} ${round(rr)} 0 0 1 ${round(x)} ${round(y + h - rr)}` +
    ` V ${round(y + rr)} A ${round(rr)} ${round(rr)} 0 0 1 ${round(x + rr)} ${round(y)} Z`;
}

/** The tail of a callout: a wedge off the bottom-left, pointing where Excel's default points. */
function calloutTail(w: number, h: number): string {
  return `M ${round(w * 0.28)} ${round(h)} L ${round(w * 0.08)} ${round(h * 1.34)} L ${round(w * 0.46)} ${round(h)} Z`;
}

/**
 * A cloud: one closed outline of bumps round an ellipse, not a pile of circles. Each pair of
 * neighbouring points on the ellipse is joined by an arc that bulges outward, so the result is a
 * single scalloped edge that fills and strokes as one shape.
 */
function cloud(w: number, h: number): string {
  const bumps = 9;
  const [cx, cy, rx, ry] = [w / 2, h / 2, w * 0.42, h * 0.4];
  const at = (i: number): [number, number] => {
    const a = (i / bumps) * Math.PI * 2 - Math.PI / 2;
    return [round(cx + rx * Math.cos(a)), round(cy + ry * Math.sin(a))];
  };
  let d = `M ${at(0)[0]} ${at(0)[1]}`;
  for (let i = 1; i <= bumps; i++) {
    const [x, y] = at(i);
    const [px, py] = at(i - 1);
    const r = round((Math.hypot(x - px, y - py) / 2) * 1.25);
    d += ` A ${r} ${r} 0 0 1 ${x} ${y}`;
  }
  return `${d} Z`;
}

/** A pie / arc slice from `from` to `to` degrees (clockwise from east), as a closed wedge. */
function wedge(w: number, h: number, from: number, to: number, close: "centre" | "chord"): string {
  const cx = w / 2, cy = h / 2, rx = w / 2, ry = h / 2;
  const at = (deg: number): [number, number] => [round(cx + rx * Math.cos((deg * Math.PI) / 180)), round(cy + ry * Math.sin((deg * Math.PI) / 180))];
  const [sx, sy] = at(from), [ex, ey] = at(to);
  const big = Math.abs(to - from) > 180 ? 1 : 0;
  const arc = `A ${round(rx)} ${round(ry)} 0 ${big} 1 ${ex} ${ey}`;
  return close === "centre" ? `M ${round(cx)} ${round(cy)} L ${sx} ${sy} ${arc} Z` : `M ${sx} ${sy} ${arc} Z`;
}

/** What a path-drawn shape needs: the closed outline to fill, plus any edges drawn on top of it. */
export interface ShapePath { d: string; lines?: string }

/** A closed path for a shape that curves, has a hole, or reaches outside its box; else null. */
export function shapeFilledPath(geom: ShapeGeom, w: number, h: number): ShapePath | null {
  switch (geom) {
    case "callout": return { d: `${roundRectPath(0, 0, w, h, 0)} ${calloutTail(w, h)}` };
    case "roundCallout": return { d: `${roundRectPath(0, 0, w, h, Math.min(w, h) * 0.15)} ${calloutTail(w, h)}` };
    case "ovalCallout":
      return { d: `M 0 ${round(h / 2)} A ${round(w / 2)} ${round(h / 2)} 0 1 0 ${round(w)} ${round(h / 2)} A ${round(w / 2)} ${round(h / 2)} 0 1 0 0 ${round(h / 2)} Z ${calloutTail(w, h)}` };
    case "pie": return { d: wedge(w, h, -90, 90, "centre") };
    case "chord": return { d: wedge(w, h, -45, 135, "chord") };
    case "donut": {
      // Two circles wound the same way: the inner one becomes the hole under evenodd.
      const [cx, cy, rx, ry] = [w / 2, h / 2, w / 2, h / 2];
      const ring = (fx: number, fy: number): string =>
        `M ${round(cx - rx * fx)} ${round(cy)} a ${round(rx * fx)} ${round(ry * fy)} 0 1 0 ${round(rx * fx * 2)} 0 a ${round(rx * fx)} ${round(ry * fy)} 0 1 0 ${round(-rx * fx * 2)} 0 Z`;
      return { d: `${ring(1, 1)} ${ring(0.5, 0.5)}` };
    }
    case "moon": {
      const r = round(w), rh = round(h / 2);
      return { d: `M ${r} 0 A ${r} ${rh} 0 0 0 ${r} ${round(h)} A ${round(w * 0.55)} ${rh} 0 0 1 ${r} 0 Z` };
    }
    case "teardrop":
      return { d: `M 0 ${round(h / 2)} A ${round(w / 2)} ${round(h / 2)} 0 0 1 ${round(w / 2)} 0 L ${round(w)} 0 L ${round(w)} ${round(h / 2)} A ${round(w / 2)} ${round(h / 2)} 0 1 1 0 ${round(h / 2)} Z` };
    case "frame": case "halfFrame": {
      const t = Math.min(w, h) * 0.16;
      const outer = `M 0 0 H ${round(w)} V ${round(h)} H 0 Z`;
      const inner = geom === "frame"
        ? `M ${round(t)} ${round(t)} V ${round(h - t)} H ${round(w - t)} V ${round(t)} Z`
        : `M ${round(t)} ${round(t)} V ${round(h)} H ${round(w)} V ${round(t)} Z`;
      return { d: `${outer} ${inner}` };
    }
    case "can": {
      const ry = h * 0.12;
      return { d: `M 0 ${round(ry)} A ${round(w / 2)} ${round(ry)} 0 0 1 ${round(w)} ${round(ry)} V ${round(h - ry)} A ${round(w / 2)} ${round(ry)} 0 0 1 0 ${round(h - ry)} Z` };
    }
    case "cube": {
      const d = Math.min(w, h) * 0.25;
      return {
        d: `M 0 ${round(d)} L ${round(d)} 0 H ${round(w)} V ${round(h - d)} L ${round(w - d)} ${round(h)} H 0 Z`,
        // The two edges of the front face, stroked over the fill rather than cut out of it.
        lines: `M 0 ${round(d)} H ${round(w - d)} V ${round(h)} M ${round(w - d)} ${round(d)} L ${round(w)} 0`,
      };
    }
    case "cloud":
      return { d: cloud(w, h) };
    case "wave": {
      const a = h * 0.18;
      return { d: `M 0 ${round(a)} Q ${round(w * 0.25)} ${round(-a)} ${round(w / 2)} ${round(a)} T ${round(w)} ${round(a)}` +
        ` V ${round(h - a)} Q ${round(w * 0.75)} ${round(h + a)} ${round(w / 2)} ${round(h - a)} T 0 ${round(h - a)} Z` };
    }
    case "heart": {
      const [cx, top] = [w / 2, h * 0.28];
      return { d: `M ${round(cx)} ${round(h)} C ${round(-w * 0.1)} ${round(h * 0.55)} ${round(w * 0.06)} 0 ${round(cx)} ${round(top)}` +
        ` C ${round(w * 0.94)} 0 ${round(w * 1.1)} ${round(h * 0.55)} ${round(cx)} ${round(h)} Z` };
    }
    case "lightningBolt":
      return { d: `M ${round(w * 0.42)} 0 L ${round(w * 0.86)} ${round(h * 0.42)} L ${round(w * 0.56)} ${round(h * 0.45)}` +
        ` L ${round(w * 0.82)} ${round(h)} L ${round(w * 0.16)} ${round(h * 0.52)} L ${round(w * 0.46)} ${round(h * 0.48)} Z` };
    case "noSmoking": {
      const [cx, cy, rx, ry] = [w / 2, h / 2, w / 2, h / 2];
      return { d: `M ${round(cx - rx)} ${round(cy)} a ${round(rx)} ${round(ry)} 0 1 0 ${round(rx * 2)} 0 a ${round(rx)} ${round(ry)} 0 1 0 ${round(-rx * 2)} 0 Z` +
        ` M ${round(w * 0.18)} ${round(h * 0.28)} L ${round(w * 0.82)} ${round(h * 0.72)} L ${round(w * 0.72)} ${round(h * 0.82)} L ${round(w * 0.18)} ${round(h * 0.38)} Z` };
    }
    default: return null;
  }
}

/** The stroked path for a brace / bracket geometry, or null for anything else. */
export function shapeOutlinePath(geom: ShapeGeom, w: number, h: number, adjust?: number): string | null {
  // A pair puts one on each edge; a single one gets the whole width, as the presets do.
  const band = Math.min(w / 2, h / 4);
  switch (geom) {
    case "leftBrace": return brace(0, w, h, true);
    case "rightBrace": return brace(0, w, h, false);
    case "bracePair": return `${brace(0, band, h, true)} ${brace(w - band, band, h, false)}`;
    case "leftBracket": return bracket(0, w, h, true);
    case "rightBracket": return bracket(0, w, h, false);
    case "bracketPair": return `${bracket(0, band, h, true)} ${bracket(w - band, band, h, false)}`;
    case "elbow": return elbow(w, h, adjust ?? 0.5);
    // A curved connector: the same journey as an elbow, rounded off.
    case "curve": return `M 0 0 C ${round(w * 0.5)} 0 ${round(w * 0.5)} ${round(h)} ${round(w)} ${round(h)}`;
    case "arc": return `M 0 ${round(h / 2)} A ${round(w / 2)} ${round(h / 2)} 0 0 1 ${round(w)} ${round(h / 2)}`;
    default: return null;
  }
}

/** A regular n-gon inscribed in the box, first vertex at top-centre. */
function regular(n: number, w: number, h: number): [number, number][] {
  const cx = w / 2, cy = h / 2, rx = w / 2, ry = h / 2;
  return Array.from({ length: n }, (_, i) => {
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
    return [round(cx + rx * Math.cos(a)), round(cy + ry * Math.sin(a))] as [number, number];
  });
}

/** A 5-point star (alternating outer / inner radius), first point at top-centre. */
function star5(w: number, h: number): [number, number][] {
  const cx = w / 2, cy = h / 2, rx = w / 2, ry = h / 2, inner = 0.382;
  return Array.from({ length: 10 }, (_, i) => {
    const a = -Math.PI / 2 + (i * Math.PI) / 5;
    const f = i % 2 === 0 ? 1 : inner;
    return [round(cx + rx * f * Math.cos(a)), round(cy + ry * f * Math.sin(a))] as [number, number];
  });
}

const round = (n: number): number => Math.round(n * 100) / 100;

/** The polygon shapes (rendered / written via a path); the rest are primitives or closed paths. */
export const POLY_GEOMS: ShapeGeom[] = [
  "triangle", "triangleDown", "diamond", "parallelogram", "trapezoid", "trapezoidDown",
  "hexagon", "pentagon", "heptagon", "octagon", "decagon", "dodecagon", "star",
  "chevron", "homePlate", "plus", "corner", "diagStripe", "manualInput", "snipRect",
  "rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow",
  "notchedArrow", "bentArrow", "quadArrow",
];
