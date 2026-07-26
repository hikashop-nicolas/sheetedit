import { parseXmlOpt, serializeXml, type Workbook } from "../../core/model";
import { THEME_SLOTS, type WorkbookTheme } from "../../core/theme";

// Switching a workbook's theme rewrites xl/theme/theme1.xml and nothing else: the cells keep their
// `<color theme="N">` references, so the new palette reaches them the same way the old one did.
//
// The part is patched in place, because <a:fmtScheme> (the effect/fill/line styles behind chart and
// shape presets) is large, unmodelled here, and would be lost by regenerating the file.

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const THEME_PART = "xl/theme/theme1.xml";

const hexOf = (css: string): string => css.replace("#", "").toUpperCase();
const childByLocal = (parent: Element, local: string): Element | undefined =>
  Array.from(parent.children).find((e) => e.localName === local);
const findByLocal = (doc: Document, local: string): Element | undefined =>
  Array.from(doc.getElementsByTagName("*")).find((e) => e.localName === local);

/** Set one clrScheme slot to a literal colour, replacing whatever kind of element was there. */
function setSlot(doc: Document, scheme: Element, slot: string, css: string): void {
  let holder = childByLocal(scheme, slot);
  if (!holder) {
    holder = doc.createElementNS(A, `a:${slot}`);
    scheme.appendChild(holder);
  }
  // dk1/lt1 are usually a sysClr; keep that shape and just refresh the resolved value, so a theme
  // that means "the window text colour" is not silently frozen to one literal.
  const existing = holder.firstElementChild;
  if (existing?.localName === "sysClr") {
    existing.setAttribute("lastClr", hexOf(css));
    return;
  }
  while (holder.firstChild) holder.removeChild(holder.firstChild);
  const c = doc.createElementNS(A, "a:srgbClr");
  c.setAttribute("val", hexOf(css));
  holder.appendChild(c);
}

/** Point a <a:majorFont> / <a:minorFont> at a typeface. */
function setSchemeFont(doc: Document, holder: Element | undefined, typeface: string | undefined): void {
  if (!holder || !typeface) return;
  let latin = childByLocal(holder, "latin");
  if (!latin) {
    latin = doc.createElementNS(A, "a:latin");
    holder.insertBefore(latin, holder.firstChild);
  }
  latin.setAttribute("typeface", typeface);
}

/** A complete theme part, for a workbook that shipped without one. */
function blankTheme(theme: WorkbookTheme): string {
  const slot = (name: string): string => `<a:${name}><a:srgbClr val="${hexOf(theme.colors[name as keyof typeof theme.colors])}"/></a:${name}>`;
  // fmtScheme is required by the schema; these are the minimum one-entry lists Excel accepts.
  const fill = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  const line = `<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">${fill}<a:prstDash val="solid"/></a:ln>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${A}" name="${theme.name}"><a:themeElements>` +
    `<a:clrScheme name="${theme.name}">${THEME_SLOTS.map(slot).join("")}</a:clrScheme>` +
    `<a:fontScheme name="${theme.name}">` +
    `<a:majorFont><a:latin typeface="${theme.majorFont ?? "Calibri Light"}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="${theme.minorFont ?? "Calibri"}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>` +
    `</a:fontScheme>` +
    `<a:fmtScheme name="${theme.name}">` +
    `<a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line}${line}${line}</a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst>` +
    `</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

/** Register a newly created theme part in the content types and the workbook relationships. */
function registerThemePart(wb: Workbook): void {
  const ct = wb.files["[Content_Types].xml"] ? parseXmlOpt(wb.files["[Content_Types].xml"]) : undefined;
  if (ct?.documentElement && !ct.documentElement.outerHTML.includes("/xl/theme/theme1.xml")) {
    const ov = ct.createElementNS(CT_NS, "Override");
    ov.setAttribute("PartName", "/xl/theme/theme1.xml");
    ov.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml");
    ct.documentElement.appendChild(ov);
    wb.files["[Content_Types].xml"] = serializeXml(ct);
  }
  const relsPath = "xl/_rels/workbook.xml.rels";
  const rels = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (rels?.documentElement) {
    const already = Array.from(rels.documentElement.children).some((r) => (r.getAttribute("Target") ?? "").includes("theme/theme1.xml"));
    if (!already) {
      const ids = Array.from(rels.documentElement.children).map((r) => Number((r.getAttribute("Id") ?? "").replace(/\D/g, "")) || 0);
      const rel = rels.createElementNS(REL_NS, "Relationship");
      rel.setAttribute("Id", `rId${Math.max(0, ...ids) + 1}`);
      rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
      rel.setAttribute("Target", "theme/theme1.xml");
      rels.documentElement.appendChild(rel);
      wb.files[relsPath] = serializeXml(rels);
    }
  }
}

/** Write the workbook's theme back into xl/theme/theme1.xml. */
export function writeXlsxTheme(wb: Workbook): void {
  if (!wb.themeDirty || !wb.theme) return;
  wb.themeDirty = false;
  const theme = wb.theme;
  const existing = wb.files[THEME_PART];
  const doc = existing ? parseXmlOpt(existing) : undefined;
  if (!doc) {
    wb.files[THEME_PART] = new TextEncoder().encode(blankTheme(theme));
    registerThemePart(wb);
    return;
  }
  const scheme = findByLocal(doc, "clrScheme");
  if (scheme) {
    scheme.setAttribute("name", theme.name);
    for (const slot of THEME_SLOTS) setSlot(doc, scheme, slot, theme.colors[slot]);
  }
  const fontScheme = findByLocal(doc, "fontScheme");
  if (fontScheme) {
    fontScheme.setAttribute("name", theme.name);
    setSchemeFont(doc, childByLocal(fontScheme, "majorFont"), theme.majorFont);
    setSchemeFont(doc, childByLocal(fontScheme, "minorFont"), theme.minorFont);
  }
  wb.files[THEME_PART] = serializeXml(doc);
}
