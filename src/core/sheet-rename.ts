import { parseXmlOpt, serializeXml, type Workbook } from "./model";
import { quoteSheetName } from "./names";
import { renameSheetInFormula } from "./sheet-span";

// A sheet rename reaches past the cell formulas: conditional formats, validations, charts,
// sparklines and pivot sources name the sheet too, in the model and in the file parts kept
// verbatim for the save. This rewrites all of them; cell formulas are handled by the caller.

/** The ODF form of a reference text ($Sheet1.A1, 'My sheet'.A1:.B2) with `from` renamed to `to`. */
export function renameSheetInOdf(text: string, from: string, to: string): string {
  if (!text.toLowerCase().includes(from.toLowerCase().replace(/'/g, "''"))) return text;
  const low = from.toLowerCase();
  const re = /(?<![\p{L}\p{N}_.'$])(\$?)('(?:[^']|'')+'|[^\s.'"[\]:;,()$!<>=+\-*/&^%{}#]+)\.(?=\$?[A-Za-z]{1,3}(?![A-Za-z])|\$?\d)/gu;
  return text.replace(re, (m, abs: string, name: string) => {
    const bare = name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
    return bare.toLowerCase() === low ? `${abs}${quoteSheetName(to)}.` : m;
  });
}

const XLSX_FORMULA_TAGS = new Set(["formula", "formula1", "formula2", "f"]);

/** Rewrite the formula texts of an xlsx part outside its cell data; true when one changed. */
function renameInXlsxDoc(doc: Document, rename: (f: string) => string): boolean {
  let changed = false;
  const walk = (el: Element): void => {
    for (const ch of Array.from(el.children)) {
      if (ch.localName === "sheetData") continue; // cell formulas come from the model
      if (XLSX_FORMULA_TAGS.has(ch.localName) && ch.children.length === 0) {
        const t = ch.textContent ?? "";
        const next = rename(t);
        if (next !== t) {
          ch.textContent = next;
          changed = true;
        }
      } else walk(ch);
    }
  };
  if (doc.documentElement) walk(doc.documentElement);
  return changed;
}

/** Rewrite every attribute of an ODF part that holds a reference; true when one changed. */
function renameInOdfDoc(doc: Document, from: string, to: string): boolean {
  let changed = false;
  const all = doc.getElementsByTagName("*");
  for (let i = 0; i < all.length; i++) {
    const el = all[i]!;
    for (const a of Array.from(el.attributes)) {
      if (!a.value.includes(".") || a.localName === "name") continue;
      const next = renameSheetInOdf(a.value, from, to);
      if (next !== a.value) {
        el.setAttributeNS(a.namespaceURI, a.name, next);
        changed = true;
      }
    }
  }
  return changed;
}

/** Rewrite the ref of every chart data reference ({ ref, cache }) in a chart model. */
function renameChartRefs(node: unknown, rename: (f: string) => string, seen = new Set<object>()): boolean {
  if (!node || typeof node !== "object" || seen.has(node)) return false;
  seen.add(node);
  let changed = false;
  const o = node as Record<string, unknown>;
  if (typeof o.ref === "string") {
    const next = rename(o.ref);
    if (next !== o.ref) {
      o.ref = next;
      changed = true;
    }
  }
  for (const [k, v] of Object.entries(o)) if (k !== "cache" && typeof v === "object") changed = renameChartRefs(v, rename, seen) || changed;
  return changed;
}

/** Apply a sheet rename everywhere outside cell formulas and defined names. */
export function renameSheetInParts(wb: Workbook, from: string, to: string): void {
  const rename = (f: string): string => renameSheetInFormula(f, from, to);
  const low = from.toLowerCase();
  for (const sheet of wb.sheets) {
    let touched = false;
    for (const cf of sheet.condFormats ?? [])
      for (const rule of cf.rules)
        if (rule.formulas) {
          const next = rule.formulas.map(rename);
          if (next.some((f, i) => f !== rule.formulas![i])) {
            rule.formulas = next;
            touched = true;
          }
        }
    for (const dv of sheet.validations ?? [])
      for (const k of ["rangeRef", "formula1", "formula2"] as const) {
        const f = dv[k];
        if (f != null && rename(f) !== f) {
          dv[k] = rename(f);
          touched = true;
        }
      }
    for (const sp of sheet.sparklines ?? []) {
      const next = rename(sp.dataRef);
      if (next !== sp.dataRef) {
        sp.dataRef = next;
        touched = true;
      }
    }
    for (const chart of sheet.charts ?? []) if (renameChartRefs(chart, rename)) touched = true;
    for (const p of sheet.pivotTables ?? []) {
      if (p.sourceSheet?.toLowerCase() === low) p.sourceSheet = to;
      if (p.hostSheet?.toLowerCase() === low) p.hostSheet = to;
    }
    if (sheet.doc && renameInXlsxDoc(sheet.doc, rename)) touched = true;
    if (touched) {
      sheet.layoutDirty = true;
      if (wb.kind === "ods") sheet.odsDirty = true;
    }
  }
  if (wb.kind === "xlsx") {
    for (const path of Object.keys(wb.files)) {
      const chart = /^xl\/charts\/chart(?:Ex)?[^/]*\.xml$/i.test(path);
      const pivot = /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/i.test(path);
      if (!chart && !pivot) continue;
      const doc = parseXmlOpt(wb.files[path]!);
      if (!doc) continue;
      let changed = chart && renameInXlsxDoc(doc, rename);
      if (pivot)
        for (const src of Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "worksheetSource"))
          if (src.getAttribute("sheet")?.toLowerCase() === low) {
            src.setAttribute("sheet", to);
            changed = true;
          }
      if (changed) wb.files[path] = serializeXml(doc);
    }
  } else if (wb.kind === "ods") {
    if (wb.contentDoc) renameInOdfDoc(wb.contentDoc, from, to);
    for (const path of Object.keys(wb.files)) {
      if (path !== "styles.xml" && !/^[^/]+\/content\.xml$/.test(path)) continue; // embedded charts
      const doc = parseXmlOpt(wb.files[path]!);
      if (doc && renameInOdfDoc(doc, from, to)) wb.files[path] = serializeXml(doc);
    }
  }
}
