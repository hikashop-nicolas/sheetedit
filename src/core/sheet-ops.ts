import { parseXmlOpt, serializeXml, type Cell, type Sheet, type Workbook } from "./model";
import { createWorksheet, uniqueSheetName } from "../adapters/xlsx/sheet-create";
import { setOdsSheetHidden } from "../adapters/ods/styles";

// Add / rename / delete / reorder sheets, for both .xlsx and .ods. Each op mutates the model
// (wb.sheets) and the file immediately (xlsx: xl/workbook.xml + rels + content-types; ods: the
// table:table elements in content.xml), so the normal write path just serializes what's there.

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ODS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

export const sheetsEditable = (wb: Workbook): boolean => wb.kind === "xlsx" || wb.kind === "ods";

/**
 * Give every sheet a stable id, and keep the ones already assigned.
 *
 * Sheets read from the file take their id from their position, because every peer read the
 * same file and so agrees without being told. Sheets added afterwards take a random one:
 * there is no shared file to derive it from, and two people each adding a sheet must end up
 * with two sheets rather than one.
 */
export function assignSheetIds(wb: Workbook): void {
  wb.sheets.forEach((sheet, index) => {
    sheet.cid ??= `s${index}`;
  });
}

/**
 * Give every picture a stable id, from its sheet and its place in that sheet.
 *
 * Safe to derive rather than randomise because images cannot be added or removed here:
 * a session can move, resize or replace one, and all three keep it in place in the list.
 */
export function assignImageIds(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    sheet.images?.forEach((im, i) => {
      im.cid ??= `${sheet.cid ?? sheet.name}-i${i}`;
    });
  }
}

/** Unique across peers, for a sheet that was not in the file. */
export function newSheetId(): string {
  const salt = new Uint8Array(4);
  crypto.getRandomValues(salt);
  return `s-${[...salt].map((b) => b.toString(16).padStart(2, "0")).join("")}`;
}

/** The sheet with this id, or undefined. Addressing by id survives a rename. */
export const sheetById = (wb: Workbook, cid: string): Sheet | undefined =>
  wb.sheets.find((s) => s.cid === cid);

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
  wb.sheets.push({ name, cid: newSheetId(), cells: new Map(), maxRow: 0, maxCol: 0, tableEl: table, odsDirty: true });
}

/** Add a new empty sheet; returns its index. */
export function addSheet(wb: Workbook, base = "Sheet"): number {
  const name = freeName(wb, base);
  if (wb.kind === "xlsx") createWorksheet(wb, name);
  else if (wb.kind === "ods") createOdsTable(wb, name);
  else throw new Error("this file type has no sheets");
  return wb.sheets.length - 1;
}

/**
 * Duplicate a sheet's contents into a new one and return its index. Cells (values, formulas,
 * number formats and their style index, which points at the workbook-wide pool either way), column
 * widths, row heights, merges and print setup come across.
 *
 * What does NOT: the drawing layer (charts, images, shapes, controls) and anything else living in
 * its own part, because each of those needs its part copied and re-registered, and a half-copied
 * drawing is worse than an absent one. A copy is a copy of the grid.
 */
export function copySheet(wb: Workbook, index: number): number {
  const src = wb.sheets[index];
  if (!src) throw new Error("no such sheet");
  const at = addSheet(wb, src.name);
  const dst = wb.sheets[at];
  if (!dst) throw new Error("the copy could not be created");
  for (const cell of src.cells.values()) {
    const copy: Cell = {
      row: cell.row, col: cell.col, value: cell.value, kind: cell.kind,
      display: cell.display, formula: cell.formula, numFmt: cell.numFmt, style: cell.style,
      cellStyle: cell.cellStyle, edited: true, fDirty: cell.formula != null, numFmtDirty: cell.numFmt != null,
    };
    dst.cells.set(`${cell.row}:${cell.col}`, copy);
  }
  dst.maxRow = src.maxRow;
  dst.maxCol = src.maxCol;
  if (src.colWidths) dst.colWidths = new Map(src.colWidths);
  if (src.rowHeights) dst.rowHeights = new Map(src.rowHeights);
  if (src.hiddenRows) dst.hiddenRows = new Set(src.hiddenRows);
  if (src.hiddenCols) dst.hiddenCols = new Set(src.hiddenCols);
  if (src.merges) dst.merges = src.merges.map((m) => ({ ...m }));
  if (src.printSetup) dst.printSetup = { ...src.printSetup };
  if (src.autoFilter) dst.autoFilter = { ...src.autoFilter };
  dst.outlineDirty = true;
  if (wb.kind === "ods") dst.odsDirty = true;
  return at;
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

/** How many sheets the user can actually see, which is what gates hiding another one. */
export const visibleSheetCount = (wb: Workbook): number => wb.sheets.filter((s) => !s.visibility).length;

/**
 * Show or hide a sheet. Both formats state this on the sheet's own element, so it is written
 * straight through rather than deferred to the save path.
 *
 * "Very hidden" is xlsx-only and cannot be undone from a UI, only by a macro, which is the whole
 * point of it; ODF has no equivalent and stores it as an ordinary hidden sheet.
 */
export function setSheetVisibility(wb: Workbook, index: number, visibility: Sheet["visibility"]): void {
  const sheet = wb.sheets[index];
  if (!sheet) throw new Error("no such sheet");
  if (visibility && !sheet.visibility && visibleSheetCount(wb) <= 1) {
    throw new Error("a workbook must keep at least one visible sheet");
  }
  sheet.visibility = visibility;
  sheet.visibilityDirty = true;
  if (wb.kind === "xlsx") {
    const doc = parseXmlOpt(wb.files["xl/workbook.xml"]);
    const el = doc && sheetEls(doc)[index];
    if (!doc || !el) return;
    if (!visibility) el.removeAttribute("state");
    else el.setAttribute("state", visibility === "veryHidden" ? "veryHidden" : "hidden");
    wb.files["xl/workbook.xml"] = serializeXml(doc);
  } else if (wb.kind === "ods") {
    const table = sheet.tableEl;
    if (table && wb.contentDoc) setOdsSheetHidden(wb.contentDoc, table, !!visibility);
  }
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
