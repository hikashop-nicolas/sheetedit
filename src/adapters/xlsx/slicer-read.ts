import { getCell, parseXmlOpt, type SheetSlicer, type Workbook } from "../../core/model";
import { anchorOf, relMap, resolvePart } from "./chart-read";
import { listWorkbookTables } from "./tables";

// Slicers: Excel's interactive pivot filters. Three parts are involved -
//   xl/slicers/slicerN.xml       the views  (<x14:slicer name cache caption columnCount .../>)
//   xl/slicerCaches/slicerCacheN.xml  the selection (<x14:slicerCacheDefinition name sourceName>
//                                     with <x14:i x="0" s="1"/> per source item)
//   the sheet's drawing           a graphicFrame carrying an sle:slicer extension, for the anchor
// Item LABELS are not stored on the slicer: `x` indexes the pivot cache field's <sharedItems>, so
// they are resolved from the cache part of a pivot table the slicer drives.

const descend = (root: Element | Document, local: string): Element[] => Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);
const attr = (el: Element, name: string): string | null => {
  const direct = el.getAttribute(name);
  if (direct != null) return direct;
  for (const a of Array.from(el.attributes)) if (a.localName === name) return a.value;
  return null;
};

/** The shared item strings of one cache field, by field name. */
function sharedItemsOf(files: Record<string, Uint8Array>, cachePart: string | undefined, fieldName: string): string[] {
  if (!cachePart || !files[cachePart]) return [];
  const doc = parseXmlOpt(files[cachePart]);
  if (!doc) return [];
  for (const cf of descend(doc, "cacheField")) {
    if ((attr(cf, "name") ?? "") !== fieldName) continue;
    const si = Array.from(cf.getElementsByTagName("*")).find((e) => e.localName === "sharedItems");
    if (!si) return [];
    return Array.from(si.children).map((e) => attr(e, "v") ?? "");
  }
  return [];
}

/** The anchor of the graphicFrame whose sle:slicer extension names this slicer. */
function slicerAnchors(files: Record<string, Uint8Array>, sheetPath: string): Map<string, import("../../core/chart-model").ChartAnchor> {
  const out = new Map<string, import("../../core/chart-model").ChartAnchor>();
  const relsPath = sheetPath.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const drawings = relMap(files, relsPath).byType.filter((r) => /drawing/i.test(r.type) && /drawings\//i.test(r.target)).map((r) => resolvePart("xl/worksheets", r.target));
  for (const dp of drawings) {
    const doc = files[dp] ? parseXmlOpt(files[dp]) : undefined;
    if (!doc) continue;
    for (const anchorEl of Array.from(doc.documentElement.children)) {
      if (!/Anchor$/.test(anchorEl.localName)) continue;
      const sl = descend(anchorEl, "slicer")[0];
      const name = sl ? attr(sl, "name") : null;
      const a = anchorOf(anchorEl);
      if (name && a) out.set(name, a);
    }
  }
  return out;
}

/** Populate sheet.slicers from the worksheet's slicer + slicerCache parts. Runs after the pivots
    have been read, so a slicer can borrow its pivot's cache part to label the items. */
export function readSlicers(wb: Workbook, files: Record<string, Uint8Array>): void {
  // cache name -> parsed cache definition
  type Cache = { path: string; sourceName: string; pivotTables: string[]; sel: Map<number, boolean>;
    kind: "pivot" | "table" | "olap"; tableId?: number; tableCol?: number; olapLabels?: string[] };
  const caches = new Map<string, Cache>();
  for (const path of Object.keys(files)) {
    if (!/^xl\/slicerCaches\/.*\.xml$/i.test(path)) continue;
    const doc = parseXmlOpt(files[path]!);
    const root = doc?.documentElement;
    if (!root || root.localName !== "slicerCacheDefinition") continue;
    const name = attr(root, "name");
    if (!name) continue;
    const sel = new Map<number, boolean>();
    // A tabular cache lists <x14:i x s>; an OLAP one lists <x14:i n c> under its level ranges.
    const tabular = descend(root, "tabular")[0];
    if (tabular) for (const i of descend(tabular, "i")) {
      const x = Number(attr(i, "x") ?? "-1");
      if (x >= 0) sel.set(x, attr(i, "s") === "1" || attr(i, "s") === "true");
    }
    // A table slicer carries an x15:tableSlicerCache in the cache's extLst (tableId + column).
    const tsc = descend(root, "tableSlicerCache")[0];
    const olap = descend(root, "olap")[0];
    let kind: Cache["kind"] = "pivot";
    let olapLabels: string[] | undefined;
    if (tsc) kind = "table";
    else if (olap && !tabular) {
      kind = "olap";
      // OLAP caches carry their own captions, so the items can at least be listed.
      const seen: string[] = [];
      for (const i of descend(olap, "i")) { const c = attr(i, "c") ?? attr(i, "n"); if (c) seen.push(c); }
      const selected = new Set(descend(olap, "selection").map((sn) => attr(sn, "n") ?? "").filter(Boolean));
      olapLabels = seen;
      seen.forEach((lbl, k) => sel.set(k, selected.size === 0 || selected.has(lbl)));
    }
    caches.set(name, {
      path,
      sourceName: attr(root, "sourceName") ?? "",
      pivotTables: descend(root, "pivotTable").map((p) => attr(p, "name") ?? "").filter(Boolean),
      sel, kind,
      tableId: tsc ? Number(attr(tsc, "tableId") ?? "0") : undefined,
      tableCol: tsc ? Number(attr(tsc, "column") ?? "0") : undefined,
      olapLabels,
    });
  }
  if (!caches.size) return;
  // Resolve table slicers against the workbook's tables (matching the table part's @id).
  const tables = listWorkbookTables(wb).map((t) => {
    const doc = wb.files[t.path] ? parseXmlOpt(wb.files[t.path]) : undefined;
    const id = Number(doc?.documentElement.getAttribute("id") ?? "0");
    const cols = doc ? Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "tableColumn").map((e) => Number(e.getAttribute("id") ?? "0")) : [];
    return { ...t, id, cols };
  });

  for (const sheet of wb.sheets) {
    if (!sheet.path) continue;
    const relsPath = sheet.path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
    const parts = relMap(files, relsPath).byType.filter((r) => /slicer(?!Cache)/i.test(r.type) || /slicers\//i.test(r.target)).map((r) => resolvePart("xl/worksheets", r.target));
    const anchors = slicerAnchors(files, sheet.path);
    const out: SheetSlicer[] = [];
    for (const sp of parts) {
      const doc = files[sp] ? parseXmlOpt(files[sp]) : undefined;
      if (!doc) continue;
      for (const sl of descend(doc, "slicer")) {
        const name = attr(sl, "name"), cacheName = attr(sl, "cache");
        if (!name || !cacheName) continue;
        const cache = caches.get(cacheName);
        if (!cache) continue;
        // Item labels come from a different place per kind.
        let labels: string[] = [];
        let table: SheetSlicer["table"];
        if (cache.kind === "olap") {
          labels = cache.olapLabels ?? [];
        } else if (cache.kind === "table") {
          const t = tables.find((x) => x.id === cache.tableId) ?? tables[0];
          if (t) {
            // `column` is a tableColumn @id; map it to a 0-based offset inside the table range.
            const ci = Math.max(0, t.cols.indexOf(cache.tableCol ?? 0));
            table = { sheetIndex: t.sheetIndex, r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2, headerRows: t.headerRows, col: ci };
            const ts = wb.sheets[t.sheetIndex];
            if (ts) {
              const seen = new Set<string>();
              for (let r = t.r1 + t.headerRows; r <= t.r2; r++) {
                const v = getCell(ts, r, t.c1 + ci);
                const s = v ? (v.display ?? v.value) : "";
                if (!seen.has(s)) { seen.add(s); labels.push(s); }
              }
            }
          }
        } else {
          // Pivot: labels are the cache field's sharedItems.
          let cachePart: string | undefined;
          for (const s2 of wb.sheets)
            for (const pt of s2.pivotTables ?? [])
              if (!cachePart && (cache.pivotTables.length === 0 || cache.pivotTables.includes(pt.name))) cachePart = pt.cachePart;
          labels = sharedItemsOf(files, cachePart, cache.sourceName);
        }
        const count = Math.max(labels.length, ...[...cache.sel.keys()].map((k) => k + 1), 0);
        const anySelected = [...cache.sel.values()].some(Boolean);
        const items = Array.from({ length: count }, (_, x) => ({
          x,
          label: labels[x] ?? String(x),
          // A cache with nothing marked selected means "no filter", not "hide everything".
          selected: anySelected ? cache.sel.get(x) === true : true,
        }));
        out.push({
          name,
          cache: cacheName,
          caption: attr(sl, "caption") ?? undefined,
          columnCount: Number(attr(sl, "columnCount") ?? "1") || 1,
          style: attr(sl, "style") ?? undefined,
          kind: cache.kind,
          table,
          sourceName: cache.sourceName,
          pivotTables: cache.pivotTables,
          items,
          anchor: anchors.get(name),
          slicerPath: sp,
          cachePath: cache.path,
        });
      }
    }
    if (out.length) sheet.slicers = out;
  }
}
