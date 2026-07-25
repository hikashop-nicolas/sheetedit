import type { Workbook } from "../../core/model";
import { SHEET_LOCK_DEFAULTS, type SheetLock } from "../../core/protection";
import { ODS } from "./shared";

// ODF states protection on the table itself: table:protected on <table:table>, and the workbook's
// sheet set via table:structure-protected on <office:spreadsheet>. The granular flags live in a
// LibreOffice extension element that must be the FIRST child of the table, and they are stated as
// PERMISSIONS, so each one is the inverse of the model's blocked-action flag.
//
// ODF has no equivalent for the format / sort / autofilter / pivot flags, so those are dropped on an
// ods save. They keep their (blocked) defaults when the file is reopened, which is the safe
// direction: nothing silently becomes editable.

/** loext permission attribute -> the model flag it inverts. */
const ODF_PERMISSIONS: [string, SheetLock][] = [
  ["select-protected-cells", "selectLockedCells"],
  ["select-unprotected-cells", "selectUnlockedCells"],
  ["insert-columns", "insertColumns"],
  ["insert-rows", "insertRows"],
  ["delete-columns", "deleteColumns"],
  ["delete-rows", "deleteRows"],
];

/** Write (or clear) protection on every sheet whose state changed, plus the workbook's. */
export function writeOdsProtections(wb: Workbook): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  for (const sheet of wb.sheets) {
    if (!sheet.protectionDirty) continue;
    const table = sheet.tableEl;
    sheet.protectionDirty = false;
    if (!table) continue;
    const ext = Array.from(table.children).find((e) => e.localName === "table-protection");
    const prot = sheet.protection;
    if (!prot?.sheet) {
      table.removeAttribute("table:protected");
      table.removeAttribute("table:protection-key");
      table.removeAttribute("table:protection-key-digest-algorithm");
      if (ext) table.removeChild(ext);
      continue;
    }
    table.setAttributeNS(ODS.table, "table:protected", "true");
    if (prot.password?.hash) {
      table.setAttributeNS(ODS.table, "table:protection-key", prot.password.hash);
      if (prot.password.algorithmName) table.setAttributeNS(ODS.table, "table:protection-key-digest-algorithm", prot.password.algorithmName);
    }
    const el = ext ?? doc.createElementNS(ODS.loext, "loext:table-protection");
    for (const [perm, flag] of ODF_PERMISSIONS) {
      const blocked = prot.locks?.[flag] ?? SHEET_LOCK_DEFAULTS[flag];
      if (blocked) el.removeAttributeNS(ODS.loext, perm);
      else el.setAttributeNS(ODS.loext, `loext:${perm}`, "true");
    }
    // The extension element is only valid as the table's first child.
    if (!ext) table.insertBefore(el, table.firstChild);
  }
  if (wb.protectionDirty) {
    wb.protectionDirty = false;
    const sheetEl = doc.getElementsByTagName("office:spreadsheet")[0];
    if (sheetEl) {
      if (wb.protection?.structure) {
        sheetEl.setAttributeNS(ODS.table, "table:structure-protected", "true");
        if (wb.protection.password?.hash) sheetEl.setAttributeNS(ODS.table, "table:protection-key", wb.protection.password.hash);
      } else {
        sheetEl.removeAttribute("table:structure-protected");
        sheetEl.removeAttribute("table:protection-key");
      }
    }
  }
}
