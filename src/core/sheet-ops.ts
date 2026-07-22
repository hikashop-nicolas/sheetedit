import { parseXmlOpt, serializeXml, type Workbook } from "./model";
import { createWorksheet, uniqueSheetName } from "../adapters/xlsx/sheet-create";

// Add / rename / delete / reorder sheets, for both .xlsx and .ods. Each op mutates the model
// (wb.sheets) and the file immediately (xlsx: xl/workbook.xml + rels + content-types; ods: the
// table:table elements in content.xml), so the normal write path just serializes what's there.

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ODS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

export const sheetsEditable = (wb: Workbook): boolean => wb.kind === "xlsx" || wb.kind === "ods";

const sheetEls = (wbDoc: Document): Element[] =>
  Array.from(wbDoc.getElementsByTagName("sheet")).filter((e) => e.localName === "sheet");

const ridOf = (el: Element): string | null => el.getAttribute("r:id") ?? el.getAttributeNS(REL_NS, "id");

/** A name not already used by another sheet (optionally ignoring the sheet at `except`). */
function freeName(wb: Workbook, base: string, except = -1): string {
  const clean = uniqueSheetName(wb, base); // handles length / illegal chars / a first collision
  const taken = new Set(wb.sheets.filter((_, i) => i !== except).map((s) => s.name.toLowerCase()));
  if (!taken.has(clean.toLowerCase())) return clean;
  const stem = clean.replace(/ \(\d+\)$/, "").slice(0, 28);
  for (let i = 2; ; i++) if (!taken.has(`${stem} (${i})`.toLowerCase())) return `${stem} (${i})`;
}

function createOdsTable(wb: Workbook, name: string): void {
  const doc = wb.contentDoc!;
  const anyTable = wb.sheets.find((s) => s.tableEl)?.tableEl;
  const parent = anyTable?.parentNode;
  if (!parent) throw new Error("no spreadsheet body to add a sheet to");
  const table = doc.createElementNS(ODS_TABLE, "table:table");
  table.setAttributeNS(ODS_TABLE, "table:name", name);
  table.appendChild(doc.createElementNS(ODS_TABLE, "table:table-column"));
  const row = doc.createElementNS(ODS_TABLE, "table:table-row");
  row.appendChild(doc.createElementNS(ODS_TABLE, "table:table-cell"));
  table.appendChild(row);
  // Insert after the last existing table (before table:named-expressions etc.).
  const last = anyTable ? Array.from(parent.childNodes).filter((n) => (n as Element).localName === "table").pop() as Element : null;
  parent.insertBefore(table, last ? last.nextSibling : null);
  wb.sheets.push({ name, cells: new Map(), maxRow: 0, maxCol: 0, tableEl: table, odsDirty: true });
}

/** Add a new empty sheet; returns its index. */
export function addSheet(wb: Workbook, base = "Sheet"): number {
  const name = freeName(wb, base);
  if (wb.kind === "xlsx") createWorksheet(wb, name);
  else if (wb.kind === "ods") createOdsTable(wb, name);
  else throw new Error("this file type has no sheets");
  return wb.sheets.length - 1;
}

/** Rename the sheet at `index`; returns the applied (deduped) name. */
export function renameSheet(wb: Workbook, index: number, name: string): string {
  const clean = freeName(wb, name.trim() || "Sheet", index);
  const sheet = wb.sheets[index];
  if (!sheet) throw new Error("no such sheet");
  sheet.name = clean;
  if (wb.kind === "xlsx") {
    const doc = parseXmlOpt(wb.files["xl/workbook.xml"]);
    const el = doc && sheetEls(doc)[index];
    if (doc && el) { el.setAttribute("name", clean); wb.files["xl/workbook.xml"] = serializeXml(doc); }
  } else if (wb.kind === "ods") {
    sheet.tableEl?.setAttributeNS(ODS_TABLE, "table:name", clean);
  }
  return clean;
}

/** Delete the sheet at `index` (a workbook must keep at least one). */
export function deleteSheet(wb: Workbook, index: number): void {
  if (wb.sheets.length <= 1) throw new Error("a workbook must keep at least one sheet");
  const sheet = wb.sheets[index];
  if (!sheet) throw new Error("no such sheet");
  if (wb.kind === "xlsx") {
    const wbDoc = parseXmlOpt(wb.files["xl/workbook.xml"]);
    const el = wbDoc && sheetEls(wbDoc)[index];
    if (wbDoc && el) {
      const rid = ridOf(el);
      el.parentNode?.removeChild(el);
      wb.files["xl/workbook.xml"] = serializeXml(wbDoc);
      const relsDoc = parseXmlOpt(wb.files["xl/_rels/workbook.xml.rels"]);
      if (relsDoc && rid) {
        for (const r of Array.from(relsDoc.getElementsByTagName("Relationship")))
          if (r.getAttribute("Id") === rid) r.parentNode?.removeChild(r);
        wb.files["xl/_rels/workbook.xml.rels"] = serializeXml(relsDoc);
      }
    }
    if (sheet.path) {
      delete wb.files[sheet.path];
      delete wb.files[sheet.path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels")];
      const ct = parseXmlOpt(wb.files["[Content_Types].xml"]);
      if (ct) {
        for (const o of Array.from(ct.getElementsByTagName("Override")))
          if (o.getAttribute("PartName") === `/${sheet.path}`) o.parentNode?.removeChild(o);
        wb.files["[Content_Types].xml"] = serializeXml(ct);
      }
    }
  } else if (wb.kind === "ods") {
    sheet.tableEl?.parentNode?.removeChild(sheet.tableEl);
  }
  wb.sheets.splice(index, 1);
}

/** Move the sheet at `from` to index `to`. */
export function moveSheet(wb: Workbook, from: number, to: number): void {
  if (from === to || !wb.sheets[from] || to < 0 || to >= wb.sheets.length) return;
  if (wb.kind === "xlsx") {
    const wbDoc = parseXmlOpt(wb.files["xl/workbook.xml"]);
    const container = wbDoc?.getElementsByTagName("sheets")[0];
    if (wbDoc && container) {
      const els = sheetEls(wbDoc);
      const [moved] = els.splice(from, 1);
      els.splice(to, 0, moved);
      els.forEach((e) => container.appendChild(e)); // appendChild moves; re-append in new order
      wb.files["xl/workbook.xml"] = serializeXml(wbDoc);
    }
  } else if (wb.kind === "ods") {
    const el = wb.sheets[from].tableEl;
    const parent = el?.parentNode;
    if (el && parent) {
      el.parentNode?.removeChild(el);
      const remaining = Array.from(parent.childNodes).filter((n) => (n as Element).localName === "table") as Element[];
      const ref = remaining[to] ?? (remaining.length ? remaining[remaining.length - 1].nextSibling : null);
      parent.insertBefore(el, ref);
    }
  }
  const [s] = wb.sheets.splice(from, 1);
  wb.sheets.splice(to, 0, s);
}
