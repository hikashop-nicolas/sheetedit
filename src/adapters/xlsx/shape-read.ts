import { parseXmlOpt, type Sheet, type ShapeGeom, type ShapeGradient, type SheetShape } from "../../core/model";
import { anchorOf, relMap, resolvePart } from "./chart-read";

// Read the drawing shapes on a worksheet: sheet rels -> drawingN.xml -> each <xdr:sp> (auto shape)
// and <xdr:cxnSp> (connector / line). Produces a geometry + fill/outline/text + cell anchor per
// shape, rendered on an SVG overlay and written back on edit; the drawing part is otherwise verbatim.

const descend = (root: Element, local: string): Element[] => Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);
const kid = (parent: Element, local: string): Element | undefined => Array.from(parent.children).find((c) => c.localName === local);

/** Map an OOXML preset geometry to one we render; unknowns fall back to a rectangle. */
function geomOf(prst: string | null): ShapeGeom {
  switch (prst) {
    case "ellipse": case "oval": return "ellipse";
    case "roundRect": case "round1Rect": case "round2SameRect": return "roundRect";
    case "triangle": case "isoscelesTriangle": case "rtTriangle": return "triangle";
    case "line": case "straightConnector1": return "line";
    case "diamond": return "diamond";
    case "parallelogram": return "parallelogram";
    case "hexagon": return "hexagon";
    case "pentagon": return "pentagon";
    case "star4": case "star5": case "star6": case "star7": case "star8": return "star";
    case "rightArrow": case "leftArrow": case "upArrow": case "downArrow": return "rightArrow";
    default: return "rect";
  }
}

/** A CSS colour from a fill / line element (srgbClr direct, schemeClr via the theme). */
function colorFrom(el: Element | undefined, theme: Record<string, string>): string | undefined {
  if (!el) return undefined;
  const clr = descend(el, "srgbClr")[0] ?? descend(el, "schemeClr")[0];
  if (!clr) return undefined;
  if (clr.localName === "srgbClr") return `#${clr.getAttribute("val")}`;
  return theme[clr.getAttribute("val") ?? ""] || undefined;
}

// --- <xdr:style>: the shape's look by reference to the theme's format scheme -----------------
// A shape Excel's gallery inserted carries no fill or line of its own. Its <xdr:style> names a
// colour and an INDEX into the theme's fillStyleLst / lnStyleLst, and that entry (usually a
// gradient of tints of the named colour) is the actual paint. Ignoring it left those shapes
// unfilled, which on a dark grid means an invisible button with unreadable text.

/** The theme's format scheme: the fill and line recipes a shape's style can point at. */
export interface ShapeStyleScheme {
  fills: Element[];
  lines: Element[];
}

const pct = (el: Element | undefined): number | undefined => {
  const v = el?.getAttribute("val");
  return v == null ? undefined : Number(v) / 100000;
};

/** Apply DrawingML's colour transforms (lum/sat/tint/shade) to a hex colour, through HSL. */
function transformColor(hex: string, el: Element): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) return hex;
  const n = parseInt(m[1]!, 16);
  const r = ((n >> 16) & 0xff) / 255, g = ((n >> 8) & 0xff) / 255, b = (n & 0xff) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0;
  const l0 = (max + min) / 2;
  const d = max - min;
  let s = d === 0 ? 0 : d / (1 - Math.abs(2 * l0 - 1));
  let l = l0;
  if (d !== 0) {
    h = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h *= 60;
    if (h < 0) h += 360;
  }
  const find = (name: string) => Array.from(el.children).find((c) => c.localName === name);
  const satMod = pct(find("satMod")); if (satMod !== undefined) s *= satMod;
  const lumMod = pct(find("lumMod")); if (lumMod !== undefined) l *= lumMod;
  const lumOff = pct(find("lumOff")); if (lumOff !== undefined) l += lumOff;
  const tint = pct(find("tint")); if (tint !== undefined) l = l * tint + (1 - tint);
  const shade = pct(find("shade")); if (shade !== undefined) l *= shade;
  s = Math.max(0, Math.min(1, s));
  l = Math.max(0, Math.min(1, l));
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const mm = l - c / 2;
  const [rr, gg, bb] =
    h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  const hx = (v: number) => Math.round((v + mm) * 255).toString(16).padStart(2, "0");
  return `#${hx(rr)}${hx(gg)}${hx(bb)}`;
}

/** The theme's fillStyleLst / lnStyleLst, in order, so a style ref can index them. */
export function readShapeStyleScheme(file: Uint8Array | undefined): ShapeStyleScheme | undefined {
  const doc = file ? parseXmlOpt(file) : undefined;
  if (!doc) return undefined;
  const fmt = descend(doc.documentElement, "fmtScheme")[0];
  if (!fmt) return undefined;
  const listOf = (name: string): Element[] => {
    const el = Array.from(fmt.children).find((c) => c.localName === name);
    return el ? Array.from(el.children) : [];
  };
  return { fills: listOf("fillStyleLst"), lines: listOf("lnStyleLst") };
}

/** One <a:srgbClr>/<a:schemeClr> resolved, with its transforms applied. `phClr` is the placeholder
    a theme recipe uses for whatever colour the shape's style names. */
function colorOfClr(clr: Element | undefined, theme: Record<string, string>, phClr?: string): string | undefined {
  if (!clr) return undefined;
  const val = clr.getAttribute("val") ?? "";
  const base = clr.localName === "srgbClr" ? `#${val}` : val === "phClr" ? phClr : theme[val];
  return base ? transformColor(base, clr) : undefined;
}

/** A fill or line element resolved to what to paint with: a colour, or a gradient. */
export interface ShapePaint {
  color?: string;
  gradient?: ShapeGradient;
}

/** <a:solidFill> / <a:gradFill> / <a:noFill> -> a paint. */
function paintOf(el: Element | undefined, theme: Record<string, string>, phClr?: string): ShapePaint | undefined {
  if (!el) return undefined;
  if (el.localName === "noFill") return undefined;
  if (el.localName === "gradFill") {
    const gsLst = Array.from(el.children).find((c) => c.localName === "gsLst");
    const stops = (gsLst ? Array.from(gsLst.children) : [])
      .map((gs) => ({ pos: Number(gs.getAttribute("pos") || "0") / 100000, color: colorOfClr(gs.firstElementChild ?? undefined, theme, phClr) }))
      .filter((s): s is { pos: number; color: string } => !!s.color);
    if (!stops.length) return undefined;
    // <a:lin ang> is in 60000ths of a degree, clockwise from east. A <a:path> (radial / from a
    // shape's centre) has no linear angle; render it top to bottom, which is what it mostly reads as.
    const lin = Array.from(el.children).find((c) => c.localName === "lin");
    const angle = lin ? Number(lin.getAttribute("ang") || "0") / 60000 : 90;
    return { color: stops[0]!.color, gradient: stops.length > 1 ? { angle, stops } : undefined };
  }
  // solidFill, or an element that simply wraps a colour (a style ref, an <a:ln>).
  const clr = Array.from(el.children).find((c) => c.localName === "srgbClr" || c.localName === "schemeClr");
  const color = colorOfClr(clr, theme, phClr);
  return color ? { color } : undefined;
}

/** The paint an <a:fillRef>/<a:lnRef> resolves to: its colour, poured into the theme recipe its
    idx names (idx 0 means none; a recipe is usually a gradient of tints of that colour). */
function styleRefPaint(ref: Element | undefined, list: Element[], theme: Record<string, string>): ShapePaint | undefined {
  if (!ref) return undefined;
  const base = colorFrom(ref, theme);
  if (!base) return undefined;
  const idx = Number(ref.getAttribute("idx") || "0");
  if (!idx) return undefined;
  const recipe = list[idx - 1];
  if (!recipe) return { color: base };
  // An <a:ln> recipe holds the fill one level down.
  const target = recipe.localName === "ln" ? Array.from(recipe.children).find((c) => /Fill$/.test(c.localName)) : recipe;
  return paintOf(target, theme, base) ?? { color: base };
}

/** Populate sheet.shapes from the worksheet's drawing parts. */
export function readShapes(
  sheet: Sheet,
  files: Record<string, Uint8Array>,
  path: string,
  theme: Record<string, string> = {},
  styleScheme?: ShapeStyleScheme,
): void {
  const relsPath = path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const drawings = relMap(files, relsPath).byType.filter((r) => /drawing/i.test(r.type) && /drawings\//i.test(r.target)).map((r) => resolvePart("xl/worksheets", r.target));
  const out: SheetShape[] = [];
  for (const drawPath of drawings) {
    const drawDoc = files[drawPath] ? parseXmlOpt(files[drawPath]) : undefined;
    if (!drawDoc) continue;
    const anchorEls = Array.from(drawDoc.documentElement.children).filter((e) => /Anchor$/.test(e.localName));
    anchorEls.forEach((anchorEl, anchorIndex) => {
      const sp = descend(anchorEl, "sp")[0] ?? descend(anchorEl, "cxnSp")[0];
      if (!sp || descend(anchorEl, "pic")[0] || descend(anchorEl, "graphicFrame")[0]) return;
      const anchor = anchorOf(anchorEl);
      if (!anchor) return;
      const spPr = kid(sp, "spPr");
      const prst = spPr ? kid(spPr, "prstGeom")?.getAttribute("prst") ?? null : null;
      const noFill = spPr ? !!kid(spPr, "noFill") : false;
      // The shape's own paint wins; a shape that states none falls back to its <xdr:style> refs.
      const style = kid(sp, "style");
      const ownFill = spPr ? paintOf(kid(spPr, "solidFill") ?? kid(spPr, "gradFill"), theme) : undefined;
      const styleFill = style && styleScheme ? styleRefPaint(kid(style, "fillRef"), styleScheme.fills, theme) : undefined;
      const paint = noFill ? undefined : ownFill ?? styleFill;
      const fill = paint?.color;
      const ln = spPr ? kid(spPr, "ln") : undefined;
      const styleStroke = style && styleScheme ? styleRefPaint(kid(style, "lnRef"), styleScheme.lines, theme) : undefined;
      const stroke = colorFrom(ln ? kid(ln, "solidFill") : undefined, theme) ?? (ln && kid(ln, "noFill") ? undefined : styleStroke?.color);
      const lw = ln?.getAttribute("w");
      const txt = descend(sp, "t").map((t) => t.textContent ?? "").join("");
      const rPr = descend(sp, "r")[0] ? kid(descend(sp, "r")[0], "rPr") : undefined;
      const styleText = style ? colorFrom(kid(style, "fontRef"), theme) : undefined;
      out.push({
        geom: geomOf(prst),
        preset: prst ?? undefined,
        anchor,
        fill,
        ...(paint?.gradient ? { fillGradient: paint.gradient } : {}),
        stroke,
        strokeWidth: lw ? Math.max(1, Math.round(Number(lw) / 9525)) : undefined,
        text: txt || undefined,
        ...(sp.getAttribute("macro") ? { macro: sp.getAttribute("macro")! } : {}),
        textColor: colorFrom(rPr ? kid(rPr, "solidFill") : undefined, theme) ?? styleText,
        drawingPath: drawPath,
        anchorIndex,
      });
    });
  }
  if (out.length) sheet.shapes = out;
}
