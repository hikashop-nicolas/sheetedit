import type { Cell, CellStyle, Sheet, StyleChange, Workbook } from "../../core/model";
import { formatNumber, mergeCellStyle } from "../../core/model";
import { isDateFmt, isTimeOnlyFmt } from "../../core/dates";
import { xmlOf } from "../xlsx/shared";
import { ODS } from "./shared";
// ---------------------------------------------------------------------------
// ods styles: automatic-style interning for user style changes
// ---------------------------------------------------------------------------

export function ensureOdsAutoStyles(doc: Document): Element {
  let el = doc.getElementsByTagName("office:automatic-styles")[0] as Element | undefined;
  if (!el) {
    el = doc.createElementNS(ODS.office, "office:automatic-styles");
    const body = doc.getElementsByTagName("office:body")[0];
    doc.documentElement.insertBefore(el, body ?? null);
  }
  return el;
}

export function findOdsStyleByName(doc: Document, name: string): Element | undefined {
  for (const s of Array.from(doc.getElementsByTagName("style:style")))
    if (s.getAttribute("style:name") === name) return s;
  return undefined;
}

// Add a built style element to <office:automatic-styles>, reusing an existing one with
// the same family + serialized properties. Returns the (existing or new) style name.
export function internOdsStyle(doc: Document, autoStyles: Element, family: string, prefix: string, styleEl: Element): string {
  styleEl.setAttributeNS(ODS.style, "style:family", family);
  const sig = Array.from(styleEl.children).map(xmlOf).join("");
  for (const ex of Array.from(autoStyles.children)) {
    if (ex.localName !== "style" || ex.getAttribute("style:family") !== family) continue;
    if (Array.from(ex.children).map(xmlOf).join("") === sig) return ex.getAttribute("style:name")!;
  }
  const used = new Set(Array.from(doc.getElementsByTagName("style:style")).map((s) => s.getAttribute("style:name")));
  let n = 1;
  while (used.has(prefix + n)) n++;
  const name = prefix + n;
  styleEl.setAttributeNS(ODS.style, "style:name", name);
  autoStyles.appendChild(styleEl);
  return name;
}

export const odsSetOrRemove = (el: Element, qn: string, v: string | undefined, ns: string = ODS.fo) => {
  if (v == null) el.removeAttribute(qn);
  else el.setAttributeNS(ns, qn, v);
};

// Apply a resolved CellStyle onto an ods cell style element (cloned from the original
// so number formats / parents survive), creating the property children as needed.
export function applyCellStyleToOds(doc: Document, st: Element, cs: CellStyle): void {
  const child = (tag: string): Element => {
    const ex = st.getElementsByTagName(tag)[0];
    if (ex) return ex;
    const el = doc.createElementNS(ODS.style, tag);
    st.appendChild(el);
    return el;
  };
  const cp = child("style:table-cell-properties");
  odsSetOrRemove(cp, "fo:background-color", cs.bg);
  cp.removeAttribute("fo:border"); // use per-side so partial borders are exact
  const bv = (c?: string) => (c ? `0.5pt solid ${c}` : undefined);
  odsSetOrRemove(cp, "fo:border-top", bv(cs.borders?.top));
  odsSetOrRemove(cp, "fo:border-right", bv(cs.borders?.right));
  odsSetOrRemove(cp, "fo:border-bottom", bv(cs.borders?.bottom));
  odsSetOrRemove(cp, "fo:border-left", bv(cs.borders?.left));
  odsSetOrRemove(cp, "fo:wrap-option", cs.wrap ? "wrap" : undefined);
  odsSetOrRemove(cp, "style:vertical-align", cs.valign, ODS.style);
  const tp = child("style:text-properties");
  odsSetOrRemove(tp, "fo:font-weight", cs.bold ? "bold" : undefined);
  odsSetOrRemove(tp, "fo:font-style", cs.italic ? "italic" : undefined);
  // Map the CSS underline flavour back to ODF: dotted/dashed/wavy as the line style, double
  // as text-underline-type. Absent flavour = plain solid.
  const uStyle = !cs.underline
    ? undefined
    : cs.underlineStyle === "dotted"
      ? "dotted"
      : cs.underlineStyle === "dashed"
        ? "dash"
        : cs.underlineStyle === "wavy"
          ? "wave"
          : "solid";
  odsSetOrRemove(tp, "style:text-underline-style", uStyle, ODS.style);
  odsSetOrRemove(tp, "style:text-underline-width", cs.underline ? "auto" : undefined, ODS.style);
  odsSetOrRemove(tp, "style:text-underline-color", cs.underline ? "font-color" : undefined, ODS.style);
  odsSetOrRemove(tp, "style:text-underline-type", cs.underline && cs.underlineStyle === "double" ? "double" : undefined, ODS.style);
  odsSetOrRemove(tp, "style:text-line-through-style", cs.strike ? "solid" : undefined, ODS.style);
  odsSetOrRemove(tp, "fo:font-size", cs.fontSize ? `${cs.fontSize}pt` : undefined);
  if (cs.fontFamily) tp.removeAttribute("style:font-name"); // fo:font-family must win over a cloned font-name
  odsSetOrRemove(tp, "fo:font-family", cs.fontFamily);
  odsSetOrRemove(tp, "fo:color", cs.color);
  const pp = child("style:paragraph-properties");
  odsSetOrRemove(pp, "fo:text-align", cs.align === "center" ? "center" : cs.align === "right" ? "end" : cs.align === "left" ? "start" : undefined);
  // Drop property children that ended up empty so dedup stays tight.
  for (const el of [cp, tp, pp]) if (el.attributes.length === 0 && el.children.length === 0) st.removeChild(el);
}

/**
 * Number format for an ods cell. The format lives in the model (grid display)
 * and maps to the cell's ODF value type on save (date/time/percentage/currency);
 * consumer apps then render it with their default format for that type.
 */
export function setOdsCellNumFmt(_wb: Workbook, _sheet: Sheet, cell: Cell, fmt: string | number | undefined, currency?: string): void {
  cell.numFmt = fmt;
  cell.numFmtDirty = false;
  cell.odsCurrency = undefined;
  if (fmt == null) cell.odsValueType = undefined;
  else if (isDateFmt(fmt)) cell.odsValueType = isTimeOnlyFmt(fmt) ? "time" : "date";
  else if (typeof fmt === "string" && fmt.includes("%")) cell.odsValueType = "percentage";
  else if (currency) {
    cell.odsValueType = "currency";
    cell.odsCurrency = currency;
  } else cell.odsValueType = undefined;
  cell.display = cell.kind === "n" && fmt != null ? formatNumber(fmt, cell.value) ?? undefined : undefined;
  cell.edited = true;
}

export function setOdsCellStyle(wb: Workbook, _sheet: Sheet, cell: Cell, change: StyleChange): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  const autoStyles = ensureOdsAutoStyles(doc);
  const desired = mergeCellStyle(cell.cellStyle ?? {}, change);
  const orig = cell.style ? findOdsStyleByName(doc, cell.style) : undefined;
  const st = orig
    ? (orig.cloneNode(true) as Element)
    : doc.createElementNS(ODS.style, "style:style");
  st.removeAttribute("style:name");
  applyCellStyleToOds(doc, st, desired);
  cell.style = internOdsStyle(doc, autoStyles, "table-cell", "ce", st);
  cell.cellStyle = desired;
  cell.edited = true;
}

// Build (or reuse) a table-column style of the given width and return its name.
export function odsColStyle(doc: Document, autoStyles: Element, px: number): string {
  const st = doc.createElementNS(ODS.style, "style:style");
  const p = doc.createElementNS(ODS.style, "style:table-column-properties");
  p.setAttributeNS(ODS.fo, "fo:break-before", "auto");
  p.setAttributeNS(ODS.style, "style:column-width", `${(px / 96) * 2.54}cm`);
  st.appendChild(p);
  return internOdsStyle(doc, autoStyles, "table-column", "co", st);
}

// Set one column's width (px), splitting the <table:table-column> run that covers it.
