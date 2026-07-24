import { parseXmlOpt, type Sheet, type ShapeGeom, type SheetShape } from "../../core/model";
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

/** Populate sheet.shapes from the worksheet's drawing parts. */
export function readShapes(sheet: Sheet, files: Record<string, Uint8Array>, path: string, theme: Record<string, string> = {}): void {
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
      const fill = noFill ? undefined : colorFrom(spPr ? kid(spPr, "solidFill") : undefined, theme);
      const ln = spPr ? kid(spPr, "ln") : undefined;
      const stroke = colorFrom(ln ? kid(ln, "solidFill") : undefined, theme);
      const lw = ln?.getAttribute("w");
      const txt = descend(sp, "t").map((t) => t.textContent ?? "").join("");
      const rPr = descend(sp, "r")[0] ? kid(descend(sp, "r")[0], "rPr") : undefined;
      out.push({
        geom: geomOf(prst),
        preset: prst ?? undefined,
        anchor,
        fill,
        stroke: stroke ?? (ln && kid(ln, "noFill") ? undefined : stroke),
        strokeWidth: lw ? Math.max(1, Math.round(Number(lw) / 9525)) : undefined,
        text: txt || undefined,
        textColor: colorFrom(rPr ? kid(rPr, "solidFill") : undefined, theme),
        drawingPath: drawPath,
        anchorIndex,
      });
    });
  }
  if (out.length) sheet.shapes = out;
}
