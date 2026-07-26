import { strFromU8, strToU8 } from "fflate";
import type { MValue } from "mlang";
import { colToLetters, getCell, parseA1Ref, parseXmlOpt, serializeXml, type Sheet, type Workbook } from "../../core/model";
import { setCellInput } from "../../core/workbook";
import { createWorksheet, uniqueSheetName } from "./sheet-create";

// Excel table (ListObject) helpers for the Power Query integration: list the workbook's
// tables, expose one as an mlang table value (Excel.CurrentWorkbook), and write a query
// result back into its range (resizing the table part's @ref). Pure model/zip level, no
// DOM; mlang is referenced as types only so the editor bundle stays free of it until the
// query panel lazy-loads the engine.

export interface WorkbookTable {
  /** The table part path, e.g. "xl/tables/table1.xml". */
  path: string;
  name: string;
  displayName: string;
  sheetIndex: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  headerRows: number; // 0 or 1
}

type MTable = Extract<MValue, { kind: "table" }>;

/** Cheap sniff: does this workbook embed Power Query definitions (a DataMashup item)?
    Real Excel writes the item as UTF-16 LE (BOM FF FE); synthetic/other producers use
    UTF-8 - decode accordingly (mirrors mlang/qdeff's decodeOoxmlText, duplicated here so
    the base bundle needs no mlang import). */
export function workbookHasQueries(files: Record<string, Uint8Array>): boolean {
  for (const [path, data] of Object.entries(files)) {
    if (!/^customXml\/item\d+\.xml$/i.test(path)) continue;
    const xml =
      data.length >= 2 && data[0] === 0xff && data[1] === 0xfe
        ? new TextDecoder("utf-16le").decode(data.subarray(2))
        : strFromU8(data);
    if (xml.includes("DataMashup")) return true;
  }
  return false;
}

/** Query names whose connection is flagged "Refresh data when opening the file"
    (refreshOnLoad in xl/connections.xml). Excel names such connections "Query - <Name>". */
export function refreshOnLoadQueries(files: Record<string, Uint8Array>): string[] {
  const data = files["xl/connections.xml"];
  if (!data) return [];
  const xml = strFromU8(data);
  const out: string[] = [];
  const re = /<connection\b[^>]*>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const tag = m[0];
    if (!/\brefreshOnLoad="(?:1|true)"/i.test(tag)) continue;
    const name = /\bname="([^"]*)"/.exec(tag)?.[1];
    if (name) out.push(name.replace(/^Query\s*-\s*/, ""));
  }
  return out;
}

function parseRef(ref: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const [a, b] = ref.split(":");
  const p1 = parseA1Ref(a ?? "");
  const p2 = b ? parseA1Ref(b) : p1;
  if (!p1 || !p2) return null;
  return { r1: p1.row, c1: p1.col, r2: p2.row, c2: p2.col };
}

/** All Excel tables in the workbook, resolved to their sheet via the worksheet rels. */
export function listWorkbookTables(wb: Workbook): WorkbookTable[] {
  // Map each table part to the sheet whose rels reference it.
  const sheetOfTable = new Map<string, number>();
  wb.sheets.forEach((sheet, si) => {
    if (!sheet.path) return;
    const relsPath = sheet.path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
    const rels = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]!) : undefined;
    if (!rels) return;
    for (const rel of Array.from(rels.getElementsByTagName("Relationship"))) {
      const target = rel.getAttribute("Target") ?? "";
      if (!/tables\/table[^/]*\.xml$/i.test(target)) continue;
      sheetOfTable.set(target.replace(/^(\.\.\/|\/xl\/)+/, "xl/").replace(/^tables\//, "xl/tables/"), si);
      // Normalize "../tables/table1.xml" -> "xl/tables/table1.xml".
      const norm = target.replace(/^\.\.\//, "xl/");
      sheetOfTable.set(norm, si);
    }
  });

  const out: WorkbookTable[] = [];
  for (const [path, data] of Object.entries(wb.files)) {
    if (!/^xl\/tables\/[^/]+\.xml$/i.test(path)) continue;
    const doc = parseXmlOpt(data);
    const root = doc?.documentElement;
    if (!root || root.localName !== "table") continue;
    const ref = parseRef(root.getAttribute("ref") ?? "");
    if (!ref) continue;
    const headerAttr = root.getAttribute("headerRowCount");
    out.push({
      path,
      name: root.getAttribute("name") ?? root.getAttribute("displayName") ?? path,
      displayName: root.getAttribute("displayName") ?? root.getAttribute("name") ?? path,
      sheetIndex: sheetOfTable.get(path) ?? 0,
      ...ref,
      headerRows: headerAttr === null ? 1 : Number(headerAttr) > 0 ? 1 : 0,
    });
  }
  return out;
}

/** Excel names a query's load-to table after the query, with invalid chars as "_". */
export function tableForQuery(tables: WorkbookTable[], queryName: string): WorkbookTable | null {
  const sanitized = queryName.replace(/[^\p{L}\p{N}_.]/gu, "_");
  return (
    tables.find((t) => t.displayName === queryName) ??
    tables.find((t) => t.name === queryName) ??
    tables.find((t) => t.displayName === sanitized || t.name === sanitized) ??
    null
  );
}

const cellValueOf = (sheet: Sheet, r: number, c: number): MValue => {
  const cell = getCell(sheet, r, c);
  if (!cell || cell.value === "") return { kind: "null" };
  if (cell.kind === "n") {
    const n = Number(cell.value);
    if (Number.isFinite(n)) return { kind: "number", value: n };
  }
  if (cell.kind === "b") return { kind: "logical", value: cell.value === "TRUE" || cell.value === "1" };
  return { kind: "text", value: cell.value };
};

/** The mlang table value for one Excel table (what Excel.CurrentWorkbook exposes). */
export function tableValue(wb: Workbook, tbl: WorkbookTable): MTable {
  const sheet = wb.sheets[tbl.sheetIndex];
  if (!sheet) return { kind: "table", columns: [], rows: [] };
  const columns: string[] = [];
  for (let c = tbl.c1; c <= tbl.c2; c++) {
    if (tbl.headerRows > 0) {
      const v = cellValueOf(sheet, tbl.r1, c);
      columns.push(v.kind === "text" ? v.value : v.kind === "number" ? String(v.value) : `Column${c - tbl.c1 + 1}`);
    } else {
      columns.push(`Column${c - tbl.c1 + 1}`);
    }
  }
  const rows: MValue[][] = [];
  for (let r = tbl.r1 + tbl.headerRows; r <= tbl.r2; r++) {
    const row: MValue[] = [];
    for (let c = tbl.c1; c <= tbl.c2; c++) row.push(cellValueOf(sheet, r, c));
    rows.push(row);
  }
  return { kind: "table", columns, rows };
}

const p2 = (n: number): string => String(n).padStart(2, "0");
const hms = (secs: number): string => `${p2(Math.floor(secs / 3600))}:${p2(Math.floor((secs % 3600) / 60))}:${p2(Math.floor(secs % 60))}`;

// Refreshed values become cell text. Temporal values are written in ISO form (unambiguous
// and sortable) rather than Excel serials, since this is a display-oriented table editor.
const rawFor = (v: MValue): string => {
  switch (v.kind) {
    case "null": return "";
    case "number": return String(v.value);
    case "text": return v.value;
    case "logical": return v.value ? "TRUE" : "FALSE";
    case "date": return `${String(v.y).padStart(4, "0")}-${p2(v.m)}-${p2(v.d)}`;
    case "time": return hms(v.secs);
    case "datetime": return `${String(v.y).padStart(4, "0")}-${p2(v.m)}-${p2(v.d)} ${hms(v.secs)}`;
    case "duration": return `${Math.trunc(v.secs / 86400)}.${hms(((v.secs % 86400) + 86400) % 86400)}`;
    default: return "";
  }
};

/** The new bottom-right of the table once `result` is loaded (>=1 data row, per Excel). */
function newExtent(tbl: WorkbookTable, result: MTable): { r2: number; c2: number } {
  return {
    r2: tbl.r1 + tbl.headerRows + Math.max(1, result.rows.length) - 1,
    c2: tbl.c1 + Math.max(1, result.columns.length) - 1,
  };
}

/** Every cell the write-back may touch (old range ∪ new range), for undo recording. */
export function touchedPositions(tbl: WorkbookTable, result: MTable): { r: number; c: number }[] {
  const ext = newExtent(tbl, result);
  const out: { r: number; c: number }[] = [];
  for (let r = tbl.r1; r <= Math.max(tbl.r2, ext.r2); r++)
    for (let c = tbl.c1; c <= Math.max(tbl.c2, ext.c2); c++) out.push({ r, c });
  return out;
}

/** Write a query result into its table: header + rows, clear stale cells, resize @ref. */
export function applyQueryResult(wb: Workbook, tbl: WorkbookTable, result: MTable): { rows: number } {
  const sheet = wb.sheets[tbl.sheetIndex];
  if (!sheet) throw new Error(`No sheet for table ${tbl.displayName}`);
  const ext = newExtent(tbl, result);

  if (tbl.headerRows > 0) {
    result.columns.forEach((name, j) => setCellInput(sheet, tbl.r1, tbl.c1 + j, name));
  }
  result.rows.forEach((row, i) => {
    const r = tbl.r1 + tbl.headerRows + i;
    result.columns.forEach((_, j) => setCellInput(sheet, r, tbl.c1 + j, rawFor(row[j] ?? { kind: "null" })));
  });
  // Clear anything the old range covered that the new result no longer does.
  for (let r = tbl.r1 + tbl.headerRows; r <= Math.max(tbl.r2, ext.r2); r++) {
    for (let c = tbl.c1; c <= Math.max(tbl.c2, ext.c2); c++) {
      const inNew = r <= ext.r2 && c <= ext.c2 && r - tbl.r1 - tbl.headerRows < result.rows.length;
      if (!inNew) setCellInput(sheet, r, c, "");
    }
  }

  // Resize the table part (and its autoFilter) to the new extent.
  const doc = parseXmlOpt(wb.files[tbl.path]!);
  if (doc) {
    const newRef = `${colToLetters(tbl.c1)}${tbl.r1}:${colToLetters(ext.c2)}${ext.r2}`;
    doc.documentElement.setAttribute("ref", newRef);
    const af = doc.documentElement.getElementsByTagName("autoFilter")[0];
    if (af) af.setAttribute("ref", newRef);
    wb.files[tbl.path] = serializeXml(doc);
  }
  tbl.r2 = ext.r2;
  tbl.c2 = ext.c2;
  return { rows: result.rows.length };
}

/** Load a query result onto a brand-new sheet (header row + data), for a query with no existing
    destination table. Returns the new sheet's index. Values are written as plain cells (not a
    live ListObject); a merchant can save and reopen the workbook to see the data. */
export function loadResultToNewSheet(wb: Workbook, queryName: string, result: MTable): { sheetIndex: number; rows: number } {
  const sheet = createWorksheet(wb, uniqueSheetName(wb, queryName));
  result.columns.forEach((name, c) => setCellInput(sheet, 1, c + 1, name));
  result.rows.forEach((row, r) => {
    result.columns.forEach((_, c) => setCellInput(sheet, r + 2, c + 1, rawFor(row[c] ?? { kind: "null" })));
  });
  return { sheetIndex: wb.sheets.length - 1, rows: result.rows.length };
}

// --- creating a table ------------------------------------------------------------------------
// ListObjects.Add from a macro. Everything Excel looks for has to be written: the table part, a
// worksheet relationship, the <tableParts> entry on the sheet, and a content-type override. The
// column names come from the header row, since a table's columns are named, not positional.

const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

const xmlEscape = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** A workbook-unique table name: Excel refuses duplicates and they break structured references. */
export function uniqueTableName(wb: Workbook, base: string): string {
  const taken = new Set(listWorkbookTables(wb).flatMap((t) => [t.name.toLowerCase(), t.displayName.toLowerCase()]));
  const clean = (base || "Table").replace(/[^A-Za-z0-9_.]/g, "_").replace(/^[^A-Za-z_]/, "_") || "Table";
  if (!taken.has(clean.toLowerCase())) return clean;
  for (let i = 2; ; i++) if (!taken.has(`${clean}${i}`.toLowerCase())) return `${clean}${i}`;
}

/**
 * Create a table (ListObject) over `rect` on `sheet`. With `hasHeaders`, the first row names the
 * columns; without, a header row is INSERTED, because a table part must name its columns and
 * Excel's own "my table has no headers" does exactly that.
 */
export function createTable(
  wb: Workbook,
  sheetIndex: number,
  rect: { r1: number; c1: number; r2: number; c2: number },
  opts: { name?: string; hasHeaders?: boolean; style?: string } = {},
): WorkbookTable {
  const sheet = wb.sheets[sheetIndex];
  if (!sheet) throw new Error("no such sheet");
  const hasHeaders = opts.hasHeaders !== false;
  if (!hasHeaders) {
    // Push the body down and write generated names, so the range the table covers still has one.
    for (let r = rect.r2; r >= rect.r1; r--)
      for (let c = rect.c1; c <= rect.c2; c++)
        setCellInput(sheet, r + 1, c, getCell(sheet, r, c)?.value ?? "");
    for (let c = rect.c1; c <= rect.c2; c++) setCellInput(sheet, rect.r1, c, `Column${c - rect.c1 + 1}`);
    rect = { ...rect, r2: rect.r2 + 1 };
  }
  const name = uniqueTableName(wb, opts.name ?? "Table1");
  let n = 1;
  while (wb.files[`xl/tables/table${n}.xml`]) n++;
  const path = `xl/tables/table${n}.xml`;

  // Column names must be unique inside the table, which is what Excel enforces on creation.
  const seen = new Set<string>();
  const columns: string[] = [];
  for (let c = rect.c1; c <= rect.c2; c++) {
    let base = String(getCell(sheet, rect.r1, c)?.value ?? "").trim() || `Column${c - rect.c1 + 1}`;
    let candidate = base, i = 2;
    while (seen.has(candidate.toLowerCase())) candidate = `${base}${i++}`;
    seen.add(candidate.toLowerCase());
    columns.push(candidate);
    if (candidate !== String(getCell(sheet, rect.r1, c)?.value ?? "")) setCellInput(sheet, rect.r1, c, candidate);
  }

  const ref = `${colToLetters(rect.c1)}${rect.r1}:${colToLetters(rect.c2)}${rect.r2}`;
  const cols = columns.map((cname, i) => `<tableColumn id="${i + 1}" name="${xmlEscape(cname)}"/>`).join("");
  const styleEl = opts.style
    ? `<tableStyleInfo name="${xmlEscape(opts.style)}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>`
    : "";
  wb.files[path] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<table xmlns="${SS_MAIN}" id="${n}" name="${xmlEscape(name)}" displayName="${xmlEscape(name)}" ref="${ref}" totalsRowShown="0">` +
      `<autoFilter ref="${ref}"/><tableColumns count="${columns.length}">${cols}</tableColumns>${styleEl}</table>`,
  );

  // The worksheet relationship, and the <tableParts> entry that points at it.
  const relsPath = (sheet.path ?? "").replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const relsDoc =
    parseXmlOpt(wb.files[relsPath]) ??
    parseXmlOpt(strToU8(`<Relationships xmlns="${PKG_REL_NS}"/>`))!;
  const rids = new Set(Array.from(relsDoc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let rn = 1;
  while (rids.has(`rId${rn}`)) rn++;
  const rid = `rId${rn}`;
  const relEl = relsDoc.createElementNS(PKG_REL_NS, "Relationship");
  relEl.setAttribute("Id", rid);
  relEl.setAttribute("Type", `${REL_NS}/table`);
  relEl.setAttribute("Target", `../tables/table${n}.xml`);
  relsDoc.documentElement.appendChild(relEl);
  wb.files[relsPath] = serializeXml(relsDoc);

  const doc = sheet.doc;
  if (doc) {
    let parts = doc.getElementsByTagName("tableParts")[0];
    if (!parts) {
      parts = doc.createElementNS(SS_MAIN, "tableParts");
      doc.documentElement.appendChild(parts); // <tableParts> is last in the schema
    }
    const part = doc.createElementNS(SS_MAIN, "tablePart");
    part.setAttribute("r:id", rid);
    parts.appendChild(part);
    parts.setAttribute("count", String(parts.children.length));
    sheet.layoutDirty = true;
  }

  const ct = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (ct && ct.documentElement.localName === "Types") {
    const ov = ct.createElementNS(CT_NS, "Override");
    ov.setAttribute("PartName", `/${path}`);
    ov.setAttribute("ContentType", CT_TABLE);
    ct.documentElement.appendChild(ov);
    wb.files["[Content_Types].xml"] = serializeXml(ct);
  }

  return { path, name, displayName: name, sheetIndex, r1: rect.r1, c1: rect.c1, r2: rect.r2, c2: rect.c2, headerRows: 1 };
}

/** Rename a table in its part, keeping name and displayName in step. */
export function renameTable(wb: Workbook, tbl: WorkbookTable, name: string): void {
  const doc = parseXmlOpt(wb.files[tbl.path]);
  if (!doc) return;
  const unique = uniqueTableName(wb, name);
  doc.documentElement.setAttribute("name", unique);
  doc.documentElement.setAttribute("displayName", unique);
  wb.files[tbl.path] = serializeXml(doc);
  tbl.name = unique;
  tbl.displayName = unique;
}

/** Read or set a table's style name (<tableStyleInfo name>). */
export function tableStyle(wb: Workbook, tbl: WorkbookTable, name?: string): string {
  const doc = parseXmlOpt(wb.files[tbl.path]);
  if (!doc) return "";
  let info = doc.getElementsByTagName("tableStyleInfo")[0];
  if (name === undefined) return info?.getAttribute("name") ?? "";
  if (!info) {
    info = doc.createElementNS(SS_MAIN, "tableStyleInfo");
    info.setAttribute("showRowStripes", "1");
    doc.documentElement.appendChild(info);
  }
  info.setAttribute("name", name);
  wb.files[tbl.path] = serializeXml(doc);
  return name;
}
