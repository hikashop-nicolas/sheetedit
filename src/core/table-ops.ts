import { colToLetters, getCell, parseXmlOpt, serializeXml, type Sheet, type Workbook } from "./model";

// Named data ranges: an xlsx table (ListObject) and an ODF named database range.
//
// The two formats call them different things and store them differently, and they are the
// same idea: a rectangle of data with a name, a header row and named columns. Everything
// that consumes one in this editor wants that idea and not the storage: structured formula
// references (Table1[Sales]), table slicers, Power Query's Excel.CurrentWorkbook, VBA's
// ListObjects, pivot sources.
//
// Until now they could only be read. Every one of those features was therefore reachable
// only on a file authored somewhere else, and the slicer button told people to "select a
// cell inside a table first" with no way to make one. That is what this fixes.

/** A named data range, in the shape both formats reduce to. */
export interface TableDef {
  /** Stable identity for collaboration; the name is not it, because a rename is an edit. */
  cid: string;
  name: string;
  /** The sheet's collaboration id. */
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  /** Whether the first row holds the column names. */
  headerRow: boolean;
  /** Column names, in order. Structured references resolve through these. */
  columns: string[];
}

export const tablesAuthorable = (wb: Workbook): boolean => wb.kind === "xlsx" || wb.kind === "ods";

/** An id for a table created during a session; unique across peers. */
export const newTableId = (): string => {
  const salt = new Uint8Array(4);
  crypto.getRandomValues(salt);
  return `tb-${[...salt].map((b) => b.toString(16).padStart(2, "0")).join("")}`;
};

/** A name nothing else is using. Table names must be unique within a workbook. */
export function freeTableName(wb: Workbook, base = "Table"): string {
  const taken = new Set((wb.tables ?? []).map((t) => t.name.toLowerCase()));
  for (let n = 1; ; n++) {
    const name = `${base}${n}`;
    if (!taken.has(name.toLowerCase())) return name;
  }
}

/** Column names from the range's first row, falling back to Column1, Column2, ... */
export function columnNames(sheet: Sheet, range: { r1: number; c1: number; c2: number }, headerRow: boolean): string[] {
  const out: string[] = [];
  for (let c = range.c1; c <= range.c2; c++) {
    const text = headerRow ? (getCell(sheet, range.r1, c)?.value ?? "").trim() : "";
    out.push(text || `Column${c - range.c1 + 1}`);
  }
  // Duplicates are not allowed: a structured reference would not know which one it meant.
  const seen = new Set<string>();
  return out.map((name) => {
    if (!seen.has(name.toLowerCase())) {
      seen.add(name.toLowerCase());
      return name;
    }
    for (let n = 2; ; n++) {
      const candidate = `${name}${n}`;
      if (seen.has(candidate.toLowerCase())) continue;
      seen.add(candidate.toLowerCase());
      return candidate;
    }
  });
}

const a1 = (r: number, c: number): string => `${colToLetters(c)}${r}`;
const rangeRef = (t: { r1: number; c1: number; r2: number; c2: number }): string => `${a1(t.r1, t.c1)}:${a1(t.r2, t.c2)}`;

// --- xlsx ------------------------------------------------------------------------------

const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

const freePart = (wb: Workbook, dir: string, stem: string): string => {
  for (let n = 1; ; n++) {
    const path = `${dir}/${stem}${n}.xml`;
    if (!wb.files[path]) return path;
  }
};

/** The next free table id: xlsx numbers tables workbook-wide and rejects a clash. */
function freeTableXmlId(wb: Workbook): number {
  let max = 0;
  for (const [path, bytes] of Object.entries(wb.files)) {
    if (!/^xl\/tables\/table[^/]*\.xml$/i.test(path)) continue;
    const doc = parseXmlOpt(bytes);
    max = Math.max(max, Number(doc?.documentElement.getAttribute("id") ?? "0"));
  }
  return max + 1;
}

function addRel(wb: Workbook, relsPath: string, type: string, target: string): string | null {
  const doc = wb.files[relsPath]
    ? parseXmlOpt(wb.files[relsPath])
    : parseXmlOpt(new TextEncoder().encode(`<?xml version="1.0"?><Relationships xmlns="${PKG_REL}"></Relationships>`));
  if (!doc) return null;
  const used = new Set(Array.from(doc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id") ?? ""));
  let id = "";
  for (let n = 1; ; n++) {
    id = `rId${n}`;
    if (!used.has(id)) break;
  }
  const rel = doc.createElementNS(PKG_REL, "Relationship");
  rel.setAttribute("Id", id);
  rel.setAttribute("Type", type);
  rel.setAttribute("Target", target);
  doc.documentElement.appendChild(rel);
  wb.files[relsPath] = serializeXml(doc);
  return id;
}

const sheetRelsPath = (sheetPath: string): string =>
  sheetPath.replace(/worksheets\/([^/]+\.xml)$/i, "worksheets/_rels/$1.rels");

function tableXml(id: number, def: TableDef): string {
  const ref = rangeRef(def);
  const cols = def.columns
    .map((name, i) => `<tableColumn id="${i + 1}" name="${name.replace(/[<>&"]/g, "")}"/>`)
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<table xmlns="${MAIN}" id="${id}" name="${def.name}" displayName="${def.name}" ref="${ref}" ` +
    `headerRowCount="${def.headerRow ? 1 : 0}" totalsRowShown="0">` +
    `<autoFilter ref="${ref}"/><tableColumns count="${def.columns.length}">${cols}</tableColumns>` +
    `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>` +
    `</table>`
  );
}

function ensureTableContentType(wb: Workbook, path: string): void {
  const raw = wb.files["[Content_Types].xml"];
  const doc = raw ? parseXmlOpt(raw) : undefined;
  if (!doc) return;
  const already = Array.from(doc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === `/${path}`);
  if (already) return;
  const ov = doc.createElementNS(CT_NS, "Override");
  ov.setAttribute("PartName", `/${path}`);
  ov.setAttribute("ContentType", CT_TABLE);
  doc.documentElement.appendChild(ov);
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

/** Point the worksheet at its table parts, which xlsx requires alongside the relationships. */
function syncTableParts(wb: Workbook, sheet: Sheet): void {
  if (!sheet.doc || !sheet.path) return;
  const relsPath = sheetRelsPath(sheet.path);
  const rels = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  const ids = rels
    ? Array.from(rels.getElementsByTagName("Relationship"))
        .filter((r) => /tables\/table[^/]*\.xml$/i.test(r.getAttribute("Target") ?? ""))
        .map((r) => r.getAttribute("Id") ?? "")
    : [];
  const root = sheet.doc.documentElement;
  for (const el of Array.from(root.children)) if (el.localName === "tableParts") root.removeChild(el);
  if (!ids.length) return;
  const parts = sheet.doc.createElementNS(MAIN, "tableParts");
  parts.setAttribute("count", String(ids.length));
  for (const id of ids) {
    const part = sheet.doc.createElementNS(MAIN, "tablePart");
    part.setAttributeNS(REL_NS, "r:id", id);
    parts.appendChild(part);
  }
  root.appendChild(parts); // last: the schema fixes the order and Excel refuses it elsewhere
  sheet.layoutDirty = true;
}

function createXlsxTable(wb: Workbook, sheet: Sheet, def: TableDef): boolean {
  if (!sheet.path || !sheet.doc) return false;
  const path = freePart(wb, "xl/tables", "table");
  wb.files[path] = new TextEncoder().encode(tableXml(freeTableXmlId(wb), def));
  ensureTableContentType(wb, path);
  const rId = addRel(wb, sheetRelsPath(sheet.path), `${REL_NS}/table`, `../tables/${path.replace(/^.*\//, "")}`);
  if (!rId) return false;
  syncTableParts(wb, sheet);
  return true;
}

/** The part backing a table, found by its name. */
function xlsxTablePart(wb: Workbook, name: string): { path: string; doc: Document } | null {
  for (const [path, bytes] of Object.entries(wb.files)) {
    if (!/^xl\/tables\/table[^/]*\.xml$/i.test(path)) continue;
    const doc = parseXmlOpt(bytes);
    if (doc?.documentElement.getAttribute("name") === name) return { path, doc };
  }
  return null;
}

function updateXlsxTable(wb: Workbook, def: TableDef, previousName: string): boolean {
  const found = xlsxTablePart(wb, previousName);
  if (!found) return false;
  const root = found.doc.documentElement;
  const ref = rangeRef(def);
  root.setAttribute("name", def.name);
  root.setAttribute("displayName", def.name);
  root.setAttribute("ref", ref);
  root.setAttribute("headerRowCount", def.headerRow ? "1" : "0");
  for (const af of Array.from(root.children)) if (af.localName === "autoFilter") af.setAttribute("ref", ref);
  for (const cols of Array.from(root.children)) {
    if (cols.localName !== "tableColumns") continue;
    while (cols.firstChild) cols.removeChild(cols.firstChild);
    cols.setAttribute("count", String(def.columns.length));
    def.columns.forEach((name, i) => {
      const col = found.doc.createElementNS(MAIN, "tableColumn");
      col.setAttribute("id", String(i + 1));
      col.setAttribute("name", name);
      cols.appendChild(col);
    });
  }
  wb.files[found.path] = serializeXml(found.doc);
  return true;
}

function deleteXlsxTable(wb: Workbook, sheet: Sheet, name: string): boolean {
  const found = xlsxTablePart(wb, name);
  if (!found || !sheet.path) return false;
  delete wb.files[found.path];

  const relsPath = sheetRelsPath(sheet.path);
  const rels = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (rels) {
    for (const rel of Array.from(rels.getElementsByTagName("Relationship"))) {
      const target = (rel.getAttribute("Target") ?? "").replace(/^.*\//, "");
      if (target === found.path.replace(/^.*\//, "")) rel.parentNode?.removeChild(rel);
    }
    wb.files[relsPath] = serializeXml(rels);
  }
  const ct = wb.files["[Content_Types].xml"] ? parseXmlOpt(wb.files["[Content_Types].xml"]) : undefined;
  if (ct) {
    for (const ov of Array.from(ct.getElementsByTagName("Override"))) {
      if (ov.getAttribute("PartName") === `/${found.path}`) ov.parentNode?.removeChild(ov);
    }
    wb.files["[Content_Types].xml"] = serializeXml(ct);
  }
  syncTableParts(wb, sheet);
  return true;
}

// --- ods -------------------------------------------------------------------------------

const ODS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

const odfTarget = (sheetName: string, t: { r1: number; c1: number; r2: number; c2: number }): string =>
  `${sheetName}.$${colToLetters(t.c1)}$${t.r1}:.$${colToLetters(t.c2)}$${t.r2}`;

/** The <table:database-ranges> container, created if the document has none. */
function odsRangesContainer(wb: Workbook): Element | null {
  const doc = wb.contentDoc;
  if (!doc) return null;
  const spreadsheet = doc.getElementsByTagName("office:spreadsheet")[0];
  if (!spreadsheet) return null;
  let container = Array.from(spreadsheet.children).find((e) => e.localName === "database-ranges");
  if (!container) {
    container = doc.createElementNS(ODS_TABLE, "table:database-ranges");
    const firstTable = spreadsheet.getElementsByTagName("table:table")[0];
    spreadsheet.insertBefore(container, firstTable ?? spreadsheet.firstChild);
  }
  return container;
}

const odsRangeNamed = (wb: Workbook, name: string): Element | undefined =>
  Array.from(wb.contentDoc?.getElementsByTagName("*") ?? [])
    .find((e) => e.localName === "database-range" && (e.getAttribute("table:name") ?? "") === name);

function writeOdsTable(wb: Workbook, sheet: Sheet, def: TableDef, previousName: string): boolean {
  const doc = wb.contentDoc;
  const container = odsRangesContainer(wb);
  if (!doc || !container) return false;
  let dr = odsRangeNamed(wb, previousName);
  if (!dr) {
    dr = doc.createElementNS(ODS_TABLE, "table:database-range");
    container.appendChild(dr);
  }
  dr.setAttributeNS(ODS_TABLE, "table:name", def.name);
  dr.setAttributeNS(ODS_TABLE, "table:target-range-address", odfTarget(sheet.name, def));
  // A named range says whether its first row is labels; the filter buttons are what make an
  // autofilter, and a named range is not one, so they stay off unless the sheet asks.
  dr.setAttributeNS(ODS_TABLE, "table:contains-header", def.headerRow ? "true" : "false");
  sheet.odsDirty = true;
  return true;
}

function deleteOdsTable(wb: Workbook, name: string): boolean {
  const dr = odsRangeNamed(wb, name);
  if (!dr) return false;
  const parent = dr.parentNode;
  parent?.removeChild(dr);
  if (parent && !(parent as Element).children.length) parent.parentNode?.removeChild(parent);
  return true;
}

// --- the operations --------------------------------------------------------------------

/** Create a named data range over `range` on `sheet`. Returns it, or null if unsupported. */
export function createTable(
  wb: Workbook,
  sheet: Sheet,
  range: { r1: number; c1: number; r2: number; c2: number },
  opts: { name?: string; headerRow?: boolean } = {},
): TableDef | null {
  if (!tablesAuthorable(wb)) return null;
  const headerRow = opts.headerRow ?? true;
  const def: TableDef = {
    cid: newTableId(),
    name: opts.name ?? freeTableName(wb),
    sheet: sheet.cid ?? sheet.name,
    ...range,
    headerRow,
    columns: columnNames(sheet, { ...range }, headerRow),
  };
  const ok = wb.kind === "xlsx" ? createXlsxTable(wb, sheet, def) : writeOdsTable(wb, sheet, def, def.name);
  if (!ok) return null;
  (wb.tables ??= []).push(def);
  return def;
}

/** Rename, move or resize an existing one. */
export function updateTable(wb: Workbook, sheet: Sheet, cid: string, next: Partial<TableDef>): boolean {
  const at = (wb.tables ?? []).findIndex((t) => t.cid === cid);
  if (at < 0) return false;
  const before = wb.tables![at];
  const merged: TableDef = { ...before, ...next, cid: before.cid };
  if (next.r1 !== undefined || next.c1 !== undefined || next.c2 !== undefined || next.headerRow !== undefined) {
    merged.columns = next.columns ?? columnNames(sheet, merged, merged.headerRow);
  }
  const ok = wb.kind === "xlsx" ? updateXlsxTable(wb, merged, before.name) : writeOdsTable(wb, sheet, merged, before.name);
  if (!ok) return false;
  wb.tables![at] = merged;
  return true;
}

/** Remove one. The cells stay; only the name and the structure around them go. */
export function deleteTable(wb: Workbook, sheet: Sheet, cid: string): boolean {
  const at = (wb.tables ?? []).findIndex((t) => t.cid === cid);
  if (at < 0) return false;
  const def = wb.tables![at];
  const ok = wb.kind === "xlsx" ? deleteXlsxTable(wb, sheet, def.name) : deleteOdsTable(wb, def.name);
  if (!ok) return false;
  wb.tables!.splice(at, 1);
  return true;
}

/**
 * Fill wb.tables from whatever the file already had.
 *
 * xlsx tables come from their parts; ODF named database ranges come from content.xml, minus
 * the anonymous one LibreOffice writes for an autofilter, which is a filter and not a table.
 * Ids are derived from position, so both peers agree without being told.
 */
export function readTables(wb: Workbook): void {
  if (wb.tables) return;
  const out: TableDef[] = [];

  if (wb.kind === "xlsx") {
    const paths = Object.keys(wb.files)
      .filter((p) => /^xl\/tables\/table[^/]*\.xml$/i.test(p))
      .sort();
    paths.forEach((path, i) => {
      const doc = parseXmlOpt(wb.files[path]);
      const root = doc?.documentElement;
      if (!root) return;
      const ref = root.getAttribute("ref") ?? "";
      const box = parseRef(ref);
      if (!box) return;
      const sheet = sheetOwning(wb, path);
      if (!sheet) return;
      out.push({
        cid: `tb${i}`,
        name: root.getAttribute("name") ?? root.getAttribute("displayName") ?? `Table${i + 1}`,
        sheet: sheet.cid ?? sheet.name,
        ...box,
        headerRow: (root.getAttribute("headerRowCount") ?? "1") !== "0",
        columns: Array.from(root.getElementsByTagName("*"))
          .filter((e) => e.localName === "tableColumn")
          .map((e, n) => e.getAttribute("name") ?? `Column${n + 1}`),
      });
    });
  } else if (wb.kind === "ods") {
    const ranges = Array.from(wb.contentDoc?.getElementsByTagName("*") ?? []).filter(
      (e) => e.localName === "database-range",
    );
    ranges.forEach((dr, i) => {
      const name = dr.getAttribute("table:name") ?? "";
      // LibreOffice writes the autofilter as an anonymous range; that is a filter, not a table.
      if (!name || name.startsWith("__Anonymous")) return;
      const target = dr.getAttribute("table:target-range-address") ?? "";
      const dot = target.indexOf(".");
      if (dot < 0) return;
      const sheetName = target.slice(0, dot);
      const box = parseRef(target.slice(dot + 1).replace(/\$/g, "").replace(/:\./, ":"));
      const sheet = wb.sheets.find((s2) => s2.name === sheetName);
      if (!box || !sheet) return;
      out.push({
        cid: `tb${i}`,
        name,
        sheet: sheet.cid ?? sheet.name,
        ...box,
        headerRow: (dr.getAttribute("table:contains-header") ?? "true") !== "false",
        columns: columnNames(sheet, box, true),
      });
    });
  }
  wb.tables = out;
}

/** "A1:C7" to a box, or null. */
function parseRef(ref: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const parts = ref.split(":");
  const one = cellRef(parts[0] ?? "");
  const two = cellRef(parts[1] ?? parts[0] ?? "");
  if (!one || !two) return null;
  return {
    r1: Math.min(one.r, two.r),
    c1: Math.min(one.c, two.c),
    r2: Math.max(one.r, two.r),
    c2: Math.max(one.c, two.c),
  };
}

function cellRef(ref: string): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref.trim());
  if (!m) return null;
  let c = 0;
  for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]), c };
}

/**
 * The sheet this table part belongs to.
 *
 * Both links are required, and deliberately: a relationship says the part exists, and the
 * worksheet's <tableParts> is what makes it a table OF that sheet. Excel refuses a file with
 * one and not the other, so a reader that accepts it would call a broken workbook fine.
 */
function sheetOwning(wb: Workbook, tablePath: string): Sheet | undefined {
  const file = tablePath.replace(/^.*\//, "");
  return wb.sheets.find((sheet) => {
    if (!sheet.path || !sheet.doc) return false;
    const rels = wb.files[sheetRelsPath(sheet.path)];
    const doc = rels ? parseXmlOpt(rels) : undefined;
    if (!doc) return false;
    const rel = Array.from(doc.getElementsByTagName("Relationship")).find(
      (r) => (r.getAttribute("Target") ?? "").replace(/^.*\//, "") === file,
    );
    if (!rel) return false;
    const id = rel.getAttribute("Id");
    return Array.from(sheet.doc.documentElement.children).some(
      (el) =>
        el.localName === "tableParts" &&
        Array.from(el.children).some((part) => part.getAttributeNS(REL_NS, "id") === id || part.getAttribute("r:id") === id),
    );
  });
}
