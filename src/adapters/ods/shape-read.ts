import type { ShapeGeom, Sheet, SheetShape, Workbook } from "../../core/model";
import { anchorOf, attrByLocal as A, descend } from "./chart-read";

// Read drawing shapes from an ODS: <draw:rect> / <draw:ellipse> / <draw:line> / <draw:custom-shape>
// anchored in a table-cell. Fill / outline come from the referenced graphic style; text from the
// shape's <text:p>. Produces a geometry + style + cell anchor per shape, rendered on the SVG overlay
// and written back on edit (the frame stays otherwise verbatim).

const GEOM: Record<string, ShapeGeom> = { rect: "rect", ellipse: "ellipse", circle: "ellipse", line: "line", "custom-shape": "rect" };

/** Map graphic style-name -> resolved fill / stroke, following parent-style-name chains. */
function graphicStyles(doc: Document): Map<string, { fill?: string; stroke?: string; strokeWidth?: number }> {
  const raw = new Map<string, Element>();
  for (const st of Array.from(doc.getElementsByTagName("*"))) {
    if (st.localName === "style" && st.getAttribute("style:family") === "graphic") {
      const name = st.getAttribute("style:name");
      if (name) raw.set(name, st);
    }
  }
  const resolved = new Map<string, { fill?: string; stroke?: string; strokeWidth?: number }>();
  const of = (name: string, seen = new Set<string>()): { fill?: string; stroke?: string; strokeWidth?: number } => {
    if (resolved.has(name)) return resolved.get(name)!;
    const el = raw.get(name);
    if (!el || seen.has(name)) return {};
    seen.add(name);
    const parent = el.getAttribute("style:parent-style-name");
    const base = parent ? { ...of(parent, seen) } : {};
    const gp = descend(el, "graphic-properties")[0];
    if (gp) {
      const fillKind = A(gp, "fill");
      if (fillKind === "none") base.fill = undefined;
      else { const fc = A(gp, "fill-color"); if (fc) base.fill = fc; }
      const strokeKind = A(gp, "stroke");
      if (strokeKind === "none") base.stroke = undefined;
      else { const sc = A(gp, "stroke-color"); if (sc) base.stroke = sc; }
      const sw = A(gp, "stroke-width");
      if (sw) { const px = odsLenPx(sw); if (px) base.strokeWidth = Math.max(1, Math.round(px)); }
    }
    resolved.set(name, base);
    return base;
  };
  for (const name of raw.keys()) of(name);
  return resolved;
}

function odsLenPx(v: string | null): number {
  if (!v) return 0;
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(v.trim());
  if (!m) return 0;
  const n = Number(m[1]);
  return m[2] === "cm" ? n * 37.795 : m[2] === "mm" ? n * 3.7795 : m[2] === "in" ? n * 96 : m[2] === "pt" ? n * (96 / 72) : m[2] === "pc" ? n * 16 : n;
}

/** Populate each sheet's shapes from the ODS content.xml drawing objects. */
export function readOdsShapes(wb: Workbook): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  const styles = graphicStyles(doc);
  for (const el of Array.from(doc.getElementsByTagName("*"))) {
    if (!(el.localName in GEOM)) continue;
    if (el.namespaceURI && !/drawing/i.test(el.namespaceURI)) continue;
    // The frame's sheet: nearest ancestor table:table.
    let t: Element | null = el.parentElement;
    while (t && t.localName !== "table") t = t.parentElement;
    const sheet: Sheet | undefined = t ? wb.sheets.find((s) => s.tableEl === t) : undefined;
    if (!sheet || !t) continue;
    const anchor = anchorOf(el, t);
    const style = styles.get(A(el, "style-name") ?? "") ?? {};
    const text = descend(el, "p").map((p) => p.textContent ?? "").join("\n").trim() || undefined;
    const shape: SheetShape = {
      geom: GEOM[el.localName]!,
      anchor,
      fill: style.fill,
      stroke: style.stroke,
      strokeWidth: style.strokeWidth,
      text: shape2d(el.localName) ? text : undefined,
      odsShapeEl: el,
      odsAnchorCol: anchor.fromCol,
      odsAnchorRow: anchor.fromRow,
    };
    (sheet.shapes ??= []).push(shape);
  }
}

const shape2d = (local: string): boolean => local !== "line";
