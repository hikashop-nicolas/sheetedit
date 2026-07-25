import { parseXmlOpt, serializeXml, type Sheet, type Workbook } from "../../core/model";
import { SHEET_LOCKS, SHEET_LOCK_DEFAULTS, type ProtectionPassword, type SheetProtection } from "../../core/protection";
import { SS_MAIN } from "./shared";
import { insertWsChild } from "./write";

// Protection is written the way Excel writes it: <sheetProtection sheet="1" .../> on the worksheet
// and <workbookProtection lockStructure="1"/> on the workbook. Each boolean attribute names a
// BLOCKED action and carries its own default, so only the flags that differ from the default are
// emitted, which keeps the element as small as Excel's.
//
// A password hash read from the file is written back verbatim. sheetedit never computes one: the
// hash is a deterrent, not encryption, and inventing one would imply a security guarantee the
// format cannot make.

/** Copy a preserved password hash back onto a protection element. */
function writePassword(el: Element, pw: ProtectionPassword | undefined, legacyAttr: string): void {
  el.removeAttribute(legacyAttr);
  for (const a of ["algorithmName", "hashValue", "saltValue", "spinCount"]) el.removeAttribute(a);
  if (!pw) return;
  if (pw.legacy) el.setAttribute(legacyAttr, pw.legacy);
  if (pw.hash) {
    el.setAttribute("hashValue", pw.hash);
    if (pw.algorithmName) el.setAttribute("algorithmName", pw.algorithmName);
    if (pw.saltValue) el.setAttribute("saltValue", pw.saltValue);
    if (pw.spinCount) el.setAttribute("spinCount", pw.spinCount);
  }
}

/** Write (or clear) one sheet's <sheetProtection> from sheet.protection. */
export function writeXlsxProtection(sheet: Sheet): void {
  const doc = sheet.doc;
  const ws = doc?.documentElement;
  if (!doc || !ws) return;
  const existing = Array.from(ws.children).find((e) => e.localName === "sheetProtection");
  const prot: SheetProtection | undefined = sheet.protection;
  if (!prot?.sheet) {
    if (existing) existing.parentNode?.removeChild(existing);
    sheet.layoutDirty = true;
    return;
  }
  const el = existing ?? doc.createElementNS(ws.namespaceURI || SS_MAIN, "sheetProtection");
  el.setAttribute("sheet", "1");
  for (const flag of SHEET_LOCKS) {
    const v = prot.locks?.[flag] ?? SHEET_LOCK_DEFAULTS[flag];
    if (v === SHEET_LOCK_DEFAULTS[flag]) el.removeAttribute(flag);
    else el.setAttribute(flag, v ? "1" : "0");
  }
  writePassword(el, prot.password, "password");
  if (!existing) insertWsChild(ws, el);
  sheet.layoutDirty = true;
}

/** Write (or clear) <workbookProtection> in xl/workbook.xml. */
export function writeXlsxWorkbookProtection(wb: Workbook): void {
  const file = wb.files["xl/workbook.xml"];
  if (!file) return;
  const doc = parseXmlOpt(file);
  const root = doc?.documentElement;
  if (!doc || !root) return;
  const existing = Array.from(root.children).find((e) => e.localName === "workbookProtection");
  const prot = wb.protection;
  if (!prot?.structure && !prot?.windows) {
    if (!existing) return; // nothing recorded and nothing to clear
    existing.parentNode?.removeChild(existing);
    wb.files["xl/workbook.xml"] = serializeXml(doc);
    return;
  }
  const el = existing ?? doc.createElementNS(root.namespaceURI || SS_MAIN, "workbookProtection");
  if (prot.structure) el.setAttribute("lockStructure", "1");
  else el.removeAttribute("lockStructure");
  if (prot.windows) el.setAttribute("lockWindows", "1");
  else el.removeAttribute("lockWindows");
  writePassword(el, prot.password, "workbookPassword");
  // CT_Workbook order: workbookProtection sits after workbookPr and before bookViews.
  if (!existing) {
    const after = ["bookViews", "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr"];
    const before = Array.from(root.children).find((e) => after.includes(e.localName)) ?? null;
    root.insertBefore(el, before);
  }
  wb.files["xl/workbook.xml"] = serializeXml(doc);
}

/** Persist every sheet whose protection changed, plus the workbook's (called from the save path). */
export function writeXlsxProtections(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    if (!sheet.protectionDirty) continue;
    writeXlsxProtection(sheet);
    sheet.protectionDirty = false;
  }
  if (wb.protectionDirty) {
    writeXlsxWorkbookProtection(wb);
    wb.protectionDirty = false;
  }
}
