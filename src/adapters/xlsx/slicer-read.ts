import { parseXmlOpt, type SheetSlicer, type Workbook } from "../../core/model";
import { anchorOf, relMap, resolvePart } from "./chart-read";

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
  const caches = new Map<string, { path: string; sourceName: string; pivotTables: string[]; sel: Map<number, boolean> }>();
  for (const path of Object.keys(files)) {
    if (!/^xl\/slicerCaches\/.*\.xml$/i.test(path)) continue;
    const doc = parseXmlOpt(files[path]!);
    const root = doc?.documentElement;
    if (!root || root.localName !== "slicerCacheDefinition") continue;
    const name = attr(root, "name");
    if (!name) continue;
    const sel = new Map<number, boolean>();
    for (const i of descend(root, "i")) {
      const x = Number(attr(i, "x") ?? "-1");
      const s = attr(i, "s");
      if (x >= 0) sel.set(x, s === "1" || s === "true");
    }
    caches.set(name, {
      path,
      sourceName: attr(root, "sourceName") ?? "",
      pivotTables: descend(root, "pivotTable").map((p) => attr(p, "name") ?? "").filter(Boolean),
      sel,
    });
  }
  if (!caches.size) return;

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
        // Label the items from the cache field of a pivot this slicer drives.
        let cachePart: string | undefined;
        for (const s2 of wb.sheets)
          for (const pt of s2.pivotTables ?? [])
            if (!cachePart && (cache.pivotTables.length === 0 || cache.pivotTables.includes(pt.name))) cachePart = pt.cachePart;
        const labels = sharedItemsOf(files, cachePart, cache.sourceName);
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
