import { parseXmlOpt, serializeXml, type Workbook } from "../../core/model";

// Persist a timeline's selected range into its cache part: <x15:selection startDate endDate> under
// <x15:state>. Clearing the range removes the selection element (Excel then shows every period).
// Everything else in the part is left as it was, so an untouched timeline round-trips unchanged.

const findAll = (root: Element | Document, local: string): Element[] =>
  Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);

export function writeXlsxTimelines(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    for (const tl of sheet.timelines ?? []) {
      if (!tl.dirty || !tl.cachePath || !wb.files[tl.cachePath]) continue;
      const doc = parseXmlOpt(wb.files[tl.cachePath]);
      const state = doc ? findAll(doc, "state")[0] : undefined;
      if (!doc || !state) continue;
      let sel = findAll(state, "selection")[0];
      if (tl.startDate && tl.endDate) {
        if (!sel) {
          // Mirror the state element's namespace/prefix so the new child matches the file.
          sel = doc.createElementNS(state.namespaceURI, state.prefix ? `${state.prefix}:selection` : "selection");
          state.insertBefore(sel, state.firstChild);
        }
        sel.setAttribute("startDate", tl.startDate);
        sel.setAttribute("endDate", tl.endDate);
      } else if (sel) sel.parentNode?.removeChild(sel);
      wb.files[tl.cachePath] = serializeXml(doc);
      tl.dirty = false;
    }
  }
}
