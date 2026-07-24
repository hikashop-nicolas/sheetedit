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
