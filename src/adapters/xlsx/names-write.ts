import { firstByLocal, parseXmlOpt, serializeXml, type Workbook } from "../../core/model";
import { SS_MAIN } from "./shared";

// Workbook-level defined names back into workbook.xml. An existing <definedName> is patched in
// place (its comment, hidden flag and so on survive); sheet-scoped and _xlnm names are untouched.

export function writeXlsxDefinedNames(wb: Workbook): void {
  if (!wb.namesDirty) return;
  const file = wb.files["xl/workbook.xml"];
  const doc = file ? parseXmlOpt(file) : null;
  const root = doc?.documentElement;
  if (!doc || !root) return;
  wb.namesDirty = false;
  const ns = root.namespaceURI || SS_MAIN;
  const want = wb.definedNames ?? new Map<string, string>();
  let names = firstByLocal(root, "definedNames");
  const kept = new Set<string>();
  for (const dn of names ? Array.from(names.children) : []) {
    const name = dn.getAttribute("name") ?? "";
    if (dn.localName !== "definedName" || dn.getAttribute("localSheetId") != null || name.startsWith("_xlnm")) continue;
    const value = want.get(name);
    if (value == null || kept.has(name)) names!.removeChild(dn);
    else {
      if (dn.textContent?.trim() !== value) dn.textContent = value;
      kept.add(name);
    }
  }
  for (const [name, value] of want) {
    if (kept.has(name)) continue;
    if (!names) {
      names = doc.createElementNS(ns, "definedNames");
      // CT_Workbook order: definedNames sits after externalReferences and before calcPr.
      const later = ["calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];
      const before = Array.from(root.children).find((e) => later.includes(e.localName)) ?? null;
      root.insertBefore(names, before);
    }
    const dn = doc.createElementNS(ns, "definedName");
    dn.setAttribute("name", name);
    dn.textContent = value;
    names.appendChild(dn);
  }
  if (names && !names.children.length) names.parentNode?.removeChild(names);
  wb.files["xl/workbook.xml"] = serializeXml(doc);
}
