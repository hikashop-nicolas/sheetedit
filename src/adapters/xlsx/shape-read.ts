import { parseXmlOpt, type Sheet, type ShapeGeom, type ShapeGradient, type SheetShape } from "../../core/model";
import { anchorOf, relMap, resolvePart } from "./chart-read";

// Read the drawing shapes on a worksheet: sheet rels -> drawingN.xml -> each <xdr:sp> (auto shape)
// and <xdr:cxnSp> (connector / line). Produces a geometry + fill/outline/text + cell anchor per
// shape, rendered on an SVG overlay and written back on edit; the drawing part is otherwise verbatim.

const descend = (root: Element, local: string): Element[] => Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);
const kid = (parent: Element, local: string): Element | undefined => Array.from(parent.children).find((c) => c.localName === local);

// OOXML names 187 preset geometries. Most are the same few drawings under different names, and a
// rectangle is a fair stand-in for a decorative one. It is NOT a fair stand-in for a shape whose
// outline carries the meaning: an arrow that points nowhere, a callout with no tail to say what it
// is about, a connector drawn as the box around it. Those are the ones listed here.
const PRESET_GEOM: Record<string, ShapeGeom> = {
  // --- primitives and polygons ---
  ellipse: "ellipse", oval: "ellipse", circle: "ellipse",
  roundRect: "roundRect", round1Rect: "roundRect", round2SameRect: "roundRect", round2DiagRect: "roundRect",
  snip1Rect: "snipRect", snip2SameRect: "snipRect", snip2DiagRect: "snipRect", snipRoundRect: "snipRect",
  triangle: "triangle", isoscelesTriangle: "triangle", rtTriangle: "triangle",
  diamond: "diamond", parallelogram: "parallelogram", trapezoid: "trapezoid",
  nonIsoscelesTrapezoid: "trapezoid", hexagon: "hexagon", pentagon: "pentagon",
  heptagon: "heptagon", octagon: "octagon", decagon: "decagon", dodecagon: "dodecagon",
  chevron: "chevron", homePlate: "homePlate", plaque: "roundRect", bevel: "rect",
  corner: "corner", diagStripe: "diagStripe", foldedCorner: "snipRect",
  frame: "frame", halfFrame: "halfFrame", can: "can", cube: "cube",
  plus: "plus", mathPlus: "plus", mathMultiply: "plus", mathMinus: "rect", mathEqual: "rect",
  mathDivide: "rect", mathNotEqual: "rect",
  // Every star is a star; the point count is detail this does not model.
  star4: "star", star5: "star", star6: "star", star7: "star", star8: "star", star10: "star",
  star12: "star", star16: "star", star24: "star", star32: "star",
  irregularSeal1: "star", irregularSeal2: "star",
  // --- arrows: the direction IS the message ---
  rightArrow: "rightArrow", leftArrow: "leftArrow", upArrow: "upArrow", downArrow: "downArrow",
  leftRightArrow: "leftRightArrow", upDownArrow: "upDownArrow",
  notchedRightArrow: "notchedArrow", stripedRightArrow: "rightArrow",
  bentArrow: "bentArrow", bentUpArrow: "bentArrow", uturnArrow: "bentArrow",
  curvedRightArrow: "rightArrow", curvedLeftArrow: "leftArrow",
  curvedUpArrow: "upArrow", curvedDownArrow: "downArrow",
  leftUpArrow: "bentArrow", leftRightUpArrow: "upArrow", swooshArrow: "rightArrow",
  quadArrow: "quadArrow", quadArrowCallout: "quadArrow",
  circularArrow: "arc", leftCircularArrow: "arc", leftRightCircularArrow: "arc",
  // --- callouts: a box (or oval) with a tail pointing at the cell it talks about ---
  wedgeRectCallout: "callout", wedgeRoundRectCallout: "roundCallout", wedgeEllipseCallout: "ovalCallout",
  cloudCallout: "ovalCallout",
  callout1: "callout", callout2: "callout", callout3: "callout",
  borderCallout1: "callout", borderCallout2: "callout", borderCallout3: "callout",
  accentCallout1: "callout", accentCallout2: "callout", accentCallout3: "callout",
  accentBorderCallout1: "callout", accentBorderCallout2: "callout", accentBorderCallout3: "callout",
  leftArrowCallout: "leftArrow", rightArrowCallout: "rightArrow",
  upArrowCallout: "upArrow", downArrowCallout: "downArrow",
  leftRightArrowCallout: "leftRightArrow", upDownArrowCallout: "upDownArrow",
  // --- connectors ---
  line: "line", lineInv: "line", straightConnector1: "line",
  bentConnector2: "elbow", bentConnector3: "elbow", bentConnector4: "elbow", bentConnector5: "elbow",
  curvedConnector2: "curve", curvedConnector3: "curve", curvedConnector4: "curve", curvedConnector5: "curve",
  arc: "arc",
  // --- braces and brackets ---
  leftBrace: "leftBrace", rightBrace: "rightBrace", bracePair: "bracePair",
  leftBracket: "leftBracket", rightBracket: "rightBracket", bracketPair: "bracketPair",
  // --- curved and hollow ---
  pie: "pie", pieWedge: "pie", chord: "chord", blockArc: "donut", donut: "donut",
  moon: "moon", teardrop: "teardrop", cloud: "cloud", wave: "wave", doubleWave: "wave",
  heart: "heart", lightningBolt: "lightningBolt", noSmoking: "noSmoking",
  sun: "star", smileyFace: "ellipse",
  // --- flowchart symbols, which are these same drawings by another name ---
  flowChartDecision: "diamond", flowChartSort: "diamond",
  flowChartConnector: "ellipse", flowChartOr: "ellipse", flowChartSummingJunction: "ellipse",
  flowChartTerminator: "roundRect", flowChartAlternateProcess: "roundRect", flowChartDelay: "roundRect",
  flowChartInputOutput: "parallelogram", flowChartPreparation: "hexagon",
  flowChartExtract: "triangle", flowChartMerge: "triangleDown",
  flowChartManualOperation: "trapezoidDown", flowChartManualInput: "manualInput",
  flowChartPunchedCard: "snipRect", flowChartOffpageConnector: "homePlate",
  flowChartMagneticDisk: "can", flowChartMagneticDrum: "can", flowChartOnlineStorage: "can",
  flowChartDocument: "flowDocument", flowChartMultidocument: "flowMultidocument",
  flowChartPunchedTape: "flowPunchedTape", flowChartCollate: "flowCollate",
  flowChartOfflineStorage: "flowOfflineStorage", flowChartMagneticTape: "flowMagneticTape",
  flowChartDisplay: "flowDisplay", flowChartPredefinedProcess: "flowPredefined",
  flowChartInternalStorage: "flowStorage",
  // A process box really is a rectangle; saying so beats falling through to one.
  flowChartProcess: "rect",
  // --- banners, scrolls and the rest ---
  ribbon: "ribbon", ribbon2: "ribbon", leftRightRibbon: "ribbon",
  ellipseRibbon: "ribbon", ellipseRibbon2: "ribbon",
  horizontalScroll: "scrollH", verticalScroll: "scrollV",
  gear6: "gear6", gear9: "gear9", funnel: "funnel",
  chartPlus: "chartPlus", chartStar: "chartStar", chartX: "chartX",
  cornerTabs: "tabs", squareTabs: "tabs", plaqueTabs: "tabs",
  // Action buttons are bevelled rounded rectangles carrying an icon. The icon is a picture on
  // top rather than the geometry, so the shape is what we draw.
  actionButtonBlank: "roundRect", actionButtonHome: "roundRect", actionButtonHelp: "roundRect",
  actionButtonInformation: "roundRect", actionButtonForwardNext: "roundRect",
  actionButtonBackPrevious: "roundRect", actionButtonEnd: "roundRect",
  actionButtonBeginning: "roundRect", actionButtonReturn: "roundRect",
  actionButtonDocument: "roundRect", actionButtonSound: "roundRect", actionButtonMovie: "roundRect",
};

/** Map an OOXML preset geometry to one we render; unknowns fall back to a rectangle. */
function geomOf(prst: string | null): ShapeGeom {
  return (prst && PRESET_GEOM[prst]) || "rect";
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
  /** 0..1 from <a:alpha>, when the fill is see-through. Kept apart from the colour so anything
      wanting a plain hex (the property editor, the writer) still gets one. */
  opacity?: number;
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
  if (!color) return undefined;
  // <a:alpha val="68000"> is a 68% opaque fill: a watermark laid over the grid has to let the
  // cells under it show through, or it is just a grey slab covering them.
  const alpha = clr ? pct(Array.from(clr.children).find((c) => c.localName === "alpha")) : undefined;
  return alpha != null && alpha < 1 ? { color, opacity: alpha } : { color };
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

/**
 * Where a grouped shape sits inside its group's box, as fractions of it.
 *
 * The group states a child coordinate space (<a:chOff>/<a:chExt>) and its own box (<a:off>/
 * <a:ext>); a child's own xfrm is in that child space. The ratio is all the renderer needs, and
 * it needs no column widths to work it out.
 */
function withinGroup(grp: Element, sp: Element): { x: number; y: number; w: number; h: number } | undefined {
  const gx = descend(grp, "grpSpPr")[0] && kid(descend(grp, "grpSpPr")[0]!, "xfrm");
  const cx = sp.getElementsByTagName("*");
  const own = Array.from(cx).find((e) => e.localName === "xfrm");
  if (!gx || !own) return undefined;
  const val = (el: Element | undefined, a: string): number => Number(el?.getAttribute(a) ?? NaN);
  const chOff = kid(gx, "chOff"), chExt = kid(gx, "chExt");
  const [bx, by] = [val(chOff, "x"), val(chOff, "y")];
  const [bw, bh] = [val(chExt, "cx"), val(chExt, "cy")];
  const off = kid(own, "off"), ext = kid(own, "ext");
  const [ox, oy, ow, oh] = [val(off, "x"), val(off, "y"), val(ext, "cx"), val(ext, "cy")];
  if (![bx, by, bw, bh, ox, oy, ow, oh].every(Number.isFinite) || bw <= 0 || bh <= 0) return undefined;
  return { x: (ox - bx) / bw, y: (oy - by) / bh, w: ow / bw, h: oh / bh };
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
      if (descend(anchorEl, "pic")[0] || descend(anchorEl, "graphicFrame")[0]) return;
      const anchor = anchorOf(anchorEl);
      if (!anchor) return;
      // A group holds several shapes in one anchor, each placed in the group's own child
      // coordinate space. Reading the first one and stopping drew a two-shape group as one shape.
      const grp = descend(anchorEl, "grpSp")[0];
      const shapes = grp
        ? Array.from(grp.children).filter((c) => c.localName === "sp" || c.localName === "cxnSp")
        : [descend(anchorEl, "sp")[0] ?? descend(anchorEl, "cxnSp")[0]].filter(Boolean);
      for (const sp of shapes as Element[]) {
      const within = grp ? withinGroup(grp, sp) : undefined;
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
      // Text: one line per <a:p>. Joining every <a:t> in the shape ran the paragraphs together,
      // so a two-line label came out as one line that then had to be clipped.
      const body = descend(sp, "txBody")[0];
      const paras = body ? Array.from(body.children).filter((c) => c.localName === "p") : [];
      const txt = paras
        .map((para) => descend(para, "t").map((t) => t.textContent ?? "").join(""))
        .join("\n")
        .replace(/\n+$/, "");
      const bodyPr = body ? kid(body, "bodyPr") : undefined;
      const anchor2 = bodyPr?.getAttribute("anchor");
      const algn = paras.map((para) => kid(para, "pPr")?.getAttribute("algn")).find(Boolean);
      const xfrm = spPr ? kid(spPr, "xfrm") : undefined;
      // <a:avLst><a:gd name="adj1" fmla="val 112661"/>: where an elbow turns, in 1000ths of a
      // percent of the run between its ends.
      // <a:ext cx cy> is the shape's size before any rotation, in EMU (1px = 9525).
      const extEl = xfrm ? kid(xfrm, "ext") : undefined;
      const extW = Number(extEl?.getAttribute("cx") ?? 0) / 9525;
      const extH = Number(extEl?.getAttribute("cy") ?? 0) / 9525;
      const ext = extW > 0 && extH > 0 ? { w: extW, h: extH } : undefined;
      const gd = spPr ? descend(spPr, "gd").find((g) => g.getAttribute("name") === "adj1") : undefined;
      const adjVal = Number(/val\s+(-?\d+)/.exec(gd?.getAttribute("fmla") ?? "")?.[1] ?? NaN);
      const adj = Number.isFinite(adjVal) ? adjVal / 100000 : undefined;
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
        ...(algn === "ctr" ? { textAlign: "center" as const } : algn === "r" ? { textAlign: "right" as const } : algn === "l" || algn === "just" ? { textAlign: "left" as const } : {}),
        ...(anchor2 === "t" ? { textValign: "top" as const } : anchor2 === "b" ? { textValign: "bottom" as const } : anchor2 === "ctr" ? { textValign: "middle" as const } : {}),
        ...(ln && kid(ln, "headEnd")?.getAttribute("type") ? { headEnd: kid(ln, "headEnd")!.getAttribute("type")! } : {}),
        ...(ln && kid(ln, "tailEnd")?.getAttribute("type") ? { tailEnd: kid(ln, "tailEnd")!.getAttribute("type")! } : {}),
        ...(xfrm?.getAttribute("flipH") === "1" ? { flipH: true } : {}),
        ...(xfrm?.getAttribute("flipV") === "1" ? { flipV: true } : {}),
        ...(Number(xfrm?.getAttribute("rot") ?? "0") ? { rotation: Number(xfrm!.getAttribute("rot")) / 60000 } : {}),
        ...(ext ? { extent: ext } : {}),
        ...(ln && kid(ln, "prstDash")?.getAttribute("val") ? { dash: kid(ln, "prstDash")!.getAttribute("val")! } : {}),
        ...(paint?.opacity != null ? { fillOpacity: paint.opacity } : {}),
        ...(adj != null ? { adjust: adj } : {}),
        ...(sp.getAttribute("macro") ? { macro: sp.getAttribute("macro")! } : {}),
        textColor: colorFrom(rPr ? kid(rPr, "solidFill") : undefined, theme) ?? styleText,
        ...(Number(rPr?.getAttribute("sz") ?? 0) ? { textSize: Number(rPr!.getAttribute("sz")) / 100 } : {}),
        ...(within ? { within } : {}),
        drawingPath: drawPath,
        // A grouped shape is drawn but not written back: its place is inside the group, and the
        // anchor index names the group, not the child.
        ...(grp ? {} : { anchorIndex }),
      });
      }
    });
  }
  if (out.length) sheet.shapes = out;
}
