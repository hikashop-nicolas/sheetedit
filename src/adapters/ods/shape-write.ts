import { type SheetShape, type Workbook } from "../../core/model";
import { ODS } from "./shared";
import { ensureOdsAutoStyles, internOdsStyle } from "./styles";
import { POLY_GEOMS, shapePoints } from "../../core/shape-geom";

// Author drawing shapes on an ods sheet. All shape writes happen on the live contentDoc before
// writeOds serializes it: an existing shape's element is patched in place (svg position + style +
// text); a brand-new shape is appended to the table's <table:shapes> container (a non-row child
// writeOds keeps verbatim and that survives re-saves), anchored by absolute svg:x/y.

const DRAW = ODS.draw, SVG = ODS.svg, TEXT = ODS.text;
const COL_W = 96, ROW_H = 24, PX_CM = 96 / 2.54;
const cm = (px: number): string => `${(px / PX_CM).toFixed(3)}cm`;
// geom -> ODF element: primitives keep their own tag; polygon shapes use a custom-shape.
const drawLocal = (geom: string): string => (geom === "ellipse" ? "ellipse" : geom === "line" ? "line" : POLY_GEOMS.includes(geom as never) ? "custom-shape" : "rect");
// our geom -> ODF enhanced-geometry draw:type (matched by the reader).
const ODF_TYPE: Record<string, string> = { diamond: "diamond", parallelogram: "parallelogram", hexagon: "hexagon", pentagon: "pentagon", star: "star5", rightArrow: "right-arrow", triangle: "isosceles-triangle" };

/** Add the geometry detail for a shape element: an enhanced-geometry for a custom-shape, or a
    corner radius for a rounded rectangle. Removes any prior enhanced-geometry first (a restyle). */
function decorateGeom(doc: Document, el: Element, sh: SheetShape): void {
  for (const c of Array.from(el.children)) if (c.localName === "enhanced-geometry") el.removeChild(c);
  if (el.localName === "custom-shape") {
    const pts = shapePoints(sh.geom, 1000, 1000) ?? [];
    const path = pts.length ? `M ${pts.map(([x, y]) => `${Math.round(x)} ${Math.round(y)}`).join(" L ")} Z N` : "";
    const eg = doc.createElementNS(DRAW, "draw:enhanced-geometry");
    eg.setAttributeNS(DRAW, "draw:type", ODF_TYPE[sh.geom] ?? "non-primitive");
    eg.setAttributeNS(SVG, "svg:viewBox", "0 0 1000 1000");
    eg.setAttributeNS(DRAW, "draw:enhanced-path", path);
    el.appendChild(eg);
  } else if (sh.geom === "roundRect") {
    el.setAttributeNS(DRAW, "draw:corner-radius", cm(6));
  }
}

/** A graphic style (fill / stroke) for a shape, interned into the doc's automatic styles. */
function graphicStyle(doc: Document, sh: SheetShape): string {
  const st = doc.createElementNS(ODS.style, "style:style");
  const gp = doc.createElementNS(ODS.style, "style:graphic-properties");
  if (sh.fill) { gp.setAttributeNS(DRAW, "draw:fill", "solid"); gp.setAttributeNS(DRAW, "draw:fill-color", sh.fill); }
  else gp.setAttributeNS(DRAW, "draw:fill", "none");
  if (sh.stroke) { gp.setAttributeNS(DRAW, "draw:stroke", "solid"); gp.setAttributeNS(SVG, "svg:stroke-width", cm(sh.strokeWidth ?? 1)); gp.setAttributeNS(SVG, "svg:stroke-color", sh.stroke); }
  else gp.setAttributeNS(DRAW, "draw:stroke", "none");
  st.appendChild(gp);
  return internOdsStyle(doc, ensureOdsAutoStyles(doc), "graphic", "gr", st);
}

function setText(doc: Document, el: Element, sh: SheetShape): void {
  for (const p of Array.from(el.children)) if (p.localName === "p") el.removeChild(p);
  if (sh.text && sh.geom !== "line") { const p = doc.createElementNS(TEXT, "text:p"); p.textContent = sh.text; el.appendChild(p); }
}

/** Position an existing (offset from its anchor cell, via odsFrame) or new (absolute) shape element. */
function setGeom(el: Element, sh: SheetShape, x: number, y: number, w: number, h: number): void {
  if (sh.geom === "line") {
    el.setAttributeNS(SVG, "svg:x1", cm(x)); el.setAttributeNS(SVG, "svg:y1", cm(y));
    el.setAttributeNS(SVG, "svg:x2", cm(x + Math.max(1, w))); el.setAttributeNS(SVG, "svg:y2", cm(y + Math.max(1, h)));
  } else {
    el.setAttributeNS(SVG, "svg:x", cm(x)); el.setAttributeNS(SVG, "svg:y", cm(y));
    el.setAttributeNS(SVG, "svg:width", cm(Math.max(1, w))); el.setAttributeNS(SVG, "svg:height", cm(Math.max(1, h)));
    for (const a of ["end-cell-address", "end-x", "end-y"]) el.removeAttributeNS(ODS.table, a);
  }
}

/** The <table:shapes> container for a table (created as its first child if absent). */
function shapesContainer(doc: Document, table: Element): Element {
  const found = Array.from(table.children).find((c) => c.localName === "shapes");
  if (found) return found;
  const el = doc.createElementNS(ODS.table, "table:shapes");
  table.insertBefore(el, table.firstChild); // before columns / rows, per the ODF content model
  return el;
}

/** Persist authored / moved / resized / restyled shapes. Call BEFORE writeOds serializes content.xml. */
export function writeOdsShapes(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    const doc = sheet.tableEl?.ownerDocument;
    for (const sh of sheet.shapes ?? []) {
      if (!sh.dirty && !sh.created) continue;
      if (sh.created) {
        if (!doc || !sheet.tableEl) continue;
        const a = sh.anchor;
        const el = doc.createElementNS(DRAW, `draw:${drawLocal(sh.geom)}`);
        el.setAttributeNS(DRAW, "draw:style-name", graphicStyle(doc, sh));
        setGeom(el, sh, (a.fromCol - 1) * COL_W + a.fromColOff, (a.fromRow - 1) * ROW_H + a.fromRowOff, (a.toCol - a.fromCol) * COL_W + a.toColOff - a.fromColOff, (a.toRow - a.fromRow) * ROW_H + a.toRowOff - a.fromRowOff);
        decorateGeom(doc, el, sh);
        setText(doc, el, sh);
        shapesContainer(doc, sheet.tableEl).appendChild(el);
        sh.odsShapeEl = el; sh.odsAnchorCol = 1; sh.odsAnchorRow = 1; sh.created = false;
      } else if (sh.odsShapeEl) {
        const d = sh.odsShapeEl.ownerDocument;
        if (sh.odsFrame) setGeom(sh.odsShapeEl, sh, sh.odsFrame.x, sh.odsFrame.y, sh.odsFrame.w, sh.odsFrame.h);
        if (d) sh.odsShapeEl.setAttributeNS(DRAW, "draw:style-name", graphicStyle(d, sh));
        if (d) setText(d, sh.odsShapeEl, sh);
      }
      sh.dirty = false;
    }
  }
}
