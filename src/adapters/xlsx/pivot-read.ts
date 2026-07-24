import type { PivotCacheRef, PivotTableInfo, Sheet, Workbook } from "../../core/model";
import { parseA1Ref, parseXmlOpt, serializeXml } from "../../core/model";

// ---------------------------------------------------------------------------
// xlsx pivot tables (read-only): detect PivotTable / PivotCache parts, model
// them for display, and record their worksheet source so a source-data edit can
// flag the cache for refresh (see markPivotSourcesDirty in the editor). The
// pivot output itself is rendered as the ordinary cells it materialises; its
// parts pass through the writer verbatim.
// ---------------------------------------------------------------------------

/** "A1:D10" (or "A1") -> a 1-based inclusive range, or null. */
function parseRange(ref: string | null): { r1: number; c1: number; r2: number; c2: number } | null {
  if (!ref) return null;
  const [a, b] = ref.replace(/\$/g, "").split(":");
  const p1 = parseA1Ref(a ?? "");
  const p2 = b ? parseA1Ref(b) : p1;
  if (!p1 || !p2) return null;
  return { r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) };
}

/** Resolve a rels file's Id -> Target map (Targets are relative to the part's directory). */
function relMap(files: Record<string, Uint8Array>, relsPath: string): Map<string, string> {
  const m = new Map<string, string>();
  const doc = files[relsPath] ? parseXmlOpt(files[relsPath]!) : undefined;
  if (doc)
    for (const r of Array.from(doc.getElementsByTagName("Relationship"))) {
      const id = r.getAttribute("Id");
      const tgt = r.getAttribute("Target");
      if (id && tgt) m.set(id, tgt);
    }
  return m;
}

/** Normalise a rels Target ("../pivotCache/x.xml") against a base dir ("xl/worksheets") -> a
    package path ("xl/pivotCache/x.xml"). Absolute Targets (leading "/") are taken as-is. */
function resolvePart(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").filter(Boolean);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

interface CacheInfo { part: string; fields: string[]; sourceSheet?: string; source?: { r1: number; c1: number; r2: number; c2: number } }

function readCache(files: Record<string, Uint8Array>, part: string): CacheInfo | undefined {
  const bytes = files[part];
  if (!bytes) return undefined;
  const doc = parseXmlOpt(bytes);
  if (!doc) return undefined;
  const info: CacheInfo = { part, fields: [] };
  for (const cf of Array.from(doc.getElementsByTagName("cacheField"))) info.fields.push(cf.getAttribute("name") ?? "");
  const ws = doc.getElementsByTagName("worksheetSource")[0];
  if (ws) {
    const sheet = ws.getAttribute("sheet") ?? undefined;
    const rng = parseRange(ws.getAttribute("ref"));
    if (sheet && rng) { info.sourceSheet = sheet; info.source = rng; }
  }
  return info;
}

function fieldNames(tableDoc: Document, tag: string, fields: string[]): string[] {
  const container = tableDoc.getElementsByTagName(tag)[0];
  if (!container) return [];
  const out: string[] = [];
  for (const f of Array.from(container.children)) {
    if (f.localName !== "field") continue;
    const x = Number(f.getAttribute("x") ?? "-1");
    if (x >= 0 && x < fields.length) out.push(fields[x]!);
    else if (x === -2) out.push("∑ Values"); // the values placeholder axis
  }
  return out;
}

export function readXlsxPivots(wb: Workbook, files: Record<string, Uint8Array>): void {
  const wbDoc = files["xl/workbook.xml"] ? parseXmlOpt(files["xl/workbook.xml"]!) : undefined;
  if (!wbDoc) return;
  // cacheId -> cache definition part, via <pivotCaches> + the workbook rels.
  const wbRels = relMap(files, "xl/_rels/workbook.xml.rels");
  const cachePartById = new Map<string, string>();
  for (const pc of Array.from(wbDoc.getElementsByTagName("pivotCache"))) {
    const id = pc.getAttribute("cacheId");
    const rid = pc.getAttribute("r:id") ?? pc.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const tgt = rid ? wbRels.get(rid) : undefined;
    if (id && tgt) cachePartById.set(id, resolvePart("xl", tgt));
  }
  const cacheInfo = new Map<string, CacheInfo>(); // part -> parsed cache
  const caches: PivotCacheRef[] = [];
  const seenCache = new Set<string>();

  for (const sheet of wb.sheets) {
    const path = sheet.path;
    if (!path) continue;
    const relsPath = path.replace(/([^/]+)$/, "_rels/$1.rels");
    const rels = relMap(files, relsPath);
    const base = path.replace(/\/[^/]+$/, "");
    const pivots: PivotTableInfo[] = [];
    for (const [, tgt] of rels) {
      if (!/pivotTable[^/]*\.xml$/i.test(tgt)) continue;
      const tablePart = resolvePart(base, tgt);
      const tbytes = files[tablePart];
      if (!tbytes) continue;
      const tdoc = parseXmlOpt(tbytes);
      if (!tdoc) continue;
      const def = tdoc.getElementsByTagName("pivotTableDefinition")[0] ?? tdoc.documentElement;
      const cacheId = def.getAttribute("cacheId");
      const cachePart = cacheId ? cachePartById.get(cacheId) : undefined;
      let cache = cachePart ? cacheInfo.get(cachePart) : undefined;
      if (!cache && cachePart) { cache = readCache(files, cachePart); if (cache) cacheInfo.set(cachePart, cache); }
      const cf = cache?.fields ?? [];
      const dataFields = Array.from(tdoc.getElementsByTagName("dataField")).map((d) => {
        const fld = Number(d.getAttribute("fld") ?? "-1");
        return { name: d.getAttribute("name") ?? cf[fld] ?? "", func: d.getAttribute("subtotal") ?? "sum" };
      });
      const info: PivotTableInfo = {
        name: def.getAttribute("name") ?? "PivotTable",
        rowFields: fieldNames(tdoc, "rowFields", cf),
        colFields: fieldNames(tdoc, "colFields", cf),
        pageFields: Array.from(tdoc.getElementsByTagName("pageField")).map((p) => {
          const n = Number(p.getAttribute("fld") ?? "-1");
          return n >= 0 && n < cf.length ? cf[n]! : "";
        }).filter(Boolean),
        dataFields,
      };
      const loc = tdoc.getElementsByTagName("location")[0];
      const tr = parseRange(loc?.getAttribute("ref") ?? null);
      if (tr) info.targetRange = tr;
      if (cache?.sourceSheet && cache.source) { info.sourceSheet = cache.sourceSheet; info.sourceRange = cache.source; }
      pivots.push(info);
      if (cache?.sourceSheet && cache.source && cache.part && !seenCache.has(cache.part)) {
        seenCache.add(cache.part);
        caches.push({ part: cache.part, sourceSheet: cache.sourceSheet, source: cache.source });
      }
    }
    if (pivots.length) attachPivots(sheet, pivots);
  }
  if (caches.length) wb.pivotCaches = caches;
}

function attachPivots(sheet: Sheet, pivots: PivotTableInfo[]): void {
  sheet.pivotTables = (sheet.pivotTables ?? []).concat(pivots);
}

/** Editing a cell inside a pivot's worksheet source leaves the pivot's cached output stale. Set
    refreshOnLoad on the affected cache definition so Excel recomputes the pivot when the file
    opens (done once per cache; the modified part rides through the writer's verbatim pass). */
export function flagXlsxPivotRefresh(wb: Workbook, sheetName: string, positions: { r: number; c: number }[]): void {
  const caches = wb.pivotCaches;
  if (!caches?.length) return;
  for (const cache of caches) {
    if (cache.refreshFlagged || cache.sourceSheet !== sheetName) continue;
    const s = cache.source;
    const hit = positions.some((p) => p.r >= s.r1 && p.r <= s.r2 && p.c >= s.c1 && p.c <= s.c2);
    if (!hit) continue;
    cache.refreshFlagged = true; // even if the part is missing, don't retry every keystroke
    const bytes = wb.files[cache.part];
    const doc = bytes ? parseXmlOpt(bytes) : undefined;
    const root = doc?.documentElement;
    if (root) {
      root.setAttribute("refreshOnLoad", "1");
      wb.files[cache.part] = serializeXml(doc!);
    }
  }
}
