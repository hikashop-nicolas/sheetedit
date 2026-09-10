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
      return poly([[0, h * 0.3], [w * 0.6, h * 0.3], [w * 0.6, 0], [w, h / 2], [w * 0.6, h], [w * 0.6, h * 0.7], [0, h * 0.7]]);
    default:
      return null;
  }
}

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

/** The polygon shapes (rendered / written via a path); the rest are primitives. */
export const POLY_GEOMS: ShapeGeom[] = ["triangle", "diamond", "parallelogram", "hexagon", "pentagon", "star", "rightArrow"];
