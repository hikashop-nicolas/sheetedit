import { parseXmlOpt, type SheetTimeline, type Workbook } from "../../core/model";
import { anchorOf, relMap, resolvePart } from "./chart-read";

// Timelines: Excel's date filter for a pivot. Like slicers they come in two parts -
//   xl/timelines/timelineN.xml        the view (<x15:timeline name cache caption level .../>)
//   xl/timelineCaches/timelineCacheN.xml  the state (<x15:state> with <x15:selection startDate
//                                         endDate> and <x15:bounds>)
// anchored by a graphicFrame carrying the timeslicer extension. The selected range filters the
// pivot: rows whose date field falls outside [startDate, endDate) are excluded.

const descend = (root: Element | Document, local: string): Element[] => Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);
const attr = (el: Element, name: string): string | null => {
  const d = el.getAttribute(name);
  if (d != null) return d;
  for (const a of Array.from(el.attributes)) if (a.localName === name) return a.value;
  return null;
};

/** The anchors of timeline graphicFrames on a sheet, by timeline name. */
function timelineAnchors(files: Record<string, Uint8Array>, sheetPath: string): Map<string, NonNullable<SheetTimeline["anchor"]>> {
  const out = new Map<string, NonNullable<SheetTimeline["anchor"]>>();
  const relsPath = sheetPath.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const drawings = relMap(files, relsPath).byType.filter((r) => /drawing/i.test(r.type) && /drawings\//i.test(r.target)).map((r) => resolvePart("xl/worksheets", r.target));
  for (const dp of drawings) {
    const doc = files[dp] ? parseXmlOpt(files[dp]) : undefined;
    if (!doc) continue;
    for (const anchorEl of Array.from(doc.documentElement.children)) {
      if (!/Anchor$/.test(anchorEl.localName)) continue;
      // The 2012 timeslicer extension names the timeline this frame shows.
      const ts = descend(anchorEl, "timeslicer")[0];
      const name = ts ? attr(ts, "name") : null;
      const a = anchorOf(anchorEl);
      if (name && a) out.set(name, a);
    }
  }
  return out;
}

/** Populate sheet.timelines from the worksheet's timeline + timelineCache parts. */
export function readTimelines(wb: Workbook, files: Record<string, Uint8Array>): void {
  // cache name -> state
  const caches = new Map<string, { path: string; sourceName: string; pivotTables: string[]; start?: string; end?: string; boundStart?: string; boundEnd?: string; filterType?: string }>();
  for (const path of Object.keys(files)) {
    if (!/^xl\/timelineCaches\/.*\.xml$/i.test(path)) continue;
    const doc = parseXmlOpt(files[path]!);
    const root = doc?.documentElement;
    if (!root || root.localName !== "timelineCacheDefinition") continue;
    const name = attr(root, "name");
    if (!name) continue;
    const state = descend(root, "state")[0];
    const sel = state ? descend(state, "selection")[0] : undefined;
    const bounds = state ? descend(state, "bounds")[0] : undefined;
    caches.set(name, {
      path,
      sourceName: attr(root, "sourceName") ?? "",
      pivotTables: descend(root, "pivotTable").map((p) => attr(p, "name") ?? "").filter(Boolean),
      start: sel ? attr(sel, "startDate") ?? undefined : undefined,
      end: sel ? attr(sel, "endDate") ?? undefined : undefined,
      boundStart: bounds ? attr(bounds, "startDate") ?? undefined : undefined,
      boundEnd: bounds ? attr(bounds, "endDate") ?? undefined : undefined,
      filterType: state ? attr(state, "filterType") ?? undefined : undefined,
    });
  }
  if (!caches.size) return;

  for (const sheet of wb.sheets) {
    if (!sheet.path) continue;
    const relsPath = sheet.path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
    const parts = relMap(files, relsPath).byType.filter((r) => /timeline(?!Cache)/i.test(r.type) || /timelines\//i.test(r.target)).map((r) => resolvePart("xl/worksheets", r.target));
    const anchors = timelineAnchors(files, sheet.path);
    const out: SheetTimeline[] = [];
    for (const tp of parts) {
      const doc = files[tp] ? parseXmlOpt(files[tp]) : undefined;
      if (!doc) continue;
      for (const tl of descend(doc, "timeline")) {
        const name = attr(tl, "name"), cacheName = attr(tl, "cache");
        if (!name || !cacheName) continue;
        const cache = caches.get(cacheName);
        if (!cache) continue;
        out.push({
          name, cache: cacheName,
          caption: attr(tl, "caption") ?? undefined,
          level: attr(tl, "level") ?? undefined,
          sourceName: cache.sourceName,
          pivotTables: cache.pivotTables,
          startDate: cache.start, endDate: cache.end,
          boundStart: cache.boundStart, boundEnd: cache.boundEnd,
          anchor: anchors.get(name),
          timelinePath: tp, cachePath: cache.path,
        });
      }
    }
    if (out.length) sheet.timelines = out;
  }
}
