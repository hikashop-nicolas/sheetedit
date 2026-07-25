import { parseXmlOpt, serializeXml, type Workbook } from "../../core/model";

// Persist a slicer's selection back into its cache part: each <x14:i x="..."> gets s="1" when the
// item is selected and no s attribute otherwise (the schema default is false). Only the selection
// changes; the rest of the cache part - pivotTables list, sortOrder, crossFilter, extLst - is left
// exactly as it was, so an untouched slicer round-trips byte-for-byte.

const attrNS = (el: Element, name: string): string | null => {
  const direct = el.getAttribute(name);
  if (direct != null) return direct;
  for (const a of Array.from(el.attributes)) if (a.localName === name) return a.value;
  return null;
};

/** Write every dirty slicer's selection into its slicerCache part. */
export function writeXlsxSlicers(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    for (const sl of sheet.slicers ?? []) {
      if (!sl.dirty || !sl.cachePath || !wb.files[sl.cachePath]) continue;
      const doc = parseXmlOpt(wb.files[sl.cachePath]);
      if (!doc) continue;
      const selected = new Map(sl.items.map((i) => [i.x, i.selected]));
      const items = Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "i");
      for (const el of items) {
        const x = Number(attrNS(el, "x") ?? "-1");
        if (x < 0) continue;
        const on = selected.get(x) === true;
        // Keep the prefix the file already uses for the attribute.
        if (on) el.setAttribute("s", "1");
        else { el.removeAttribute("s"); for (const a of Array.from(el.attributes)) if (a.localName === "s") el.removeAttributeNode(a); }
      }
      wb.files[sl.cachePath] = serializeXml(doc);
      sl.dirty = false;
    }
  }
}
