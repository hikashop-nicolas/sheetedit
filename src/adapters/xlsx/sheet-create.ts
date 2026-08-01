import { strToU8 } from "fflate";
import { parseXmlOpt, serializeXml, type Sheet, type Workbook } from "../../core/model";
import { newSheetId } from "../../core/sheet-ops";

// Create a brand-new worksheet in an .xlsx workbook, registered everywhere Excel and the
// reader look: the worksheet part, a <sheet> entry in xl/workbook.xml, a relationship in
// xl/_rels/workbook.xml.rels, and a content-type override. Used as a Load-To destination for a
// Power Query result that has no existing table. Pure model/zip level, no DOM chrome.

const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

/** A workbook-unique sheet name (Excel caps names at 31 chars and forbids []:*?/\\). */
export function uniqueSheetName(wb: Workbook, base: string): string {
  const clean = (base || "Sheet").replace(/[[\]:*?/\\]/g, " ").trim().slice(0, 31) || "Sheet";
  const taken = new Set(wb.sheets.map((s) => s.name.toLowerCase()));
  if (!taken.has(clean.toLowerCase())) return clean;
  for (let i = 2; ; i++) {
    const cand = `${clean.slice(0, 28)} (${i})`;
    if (!taken.has(cand.toLowerCase())) return cand;
  }
}

/** Create an empty worksheet and return its model Sheet (already appended to wb.sheets). */
export function createWorksheet(wb: Workbook, name: string): Sheet {
  let n = 1;
  while (wb.files[`xl/worksheets/sheet${n}.xml`]) n++;
  const rel = `worksheets/sheet${n}.xml`;
  const path = `xl/${rel}`;
  const doc = parseXmlOpt(strToU8(`<worksheet xmlns="${SS_MAIN}" xmlns:r="${REL_NS}"><sheetData/></worksheet>`));
  const sheetData = doc?.getElementsByTagName("sheetData")[0];
  if (!doc || !sheetData) throw new Error("could not create the worksheet part");

  const wbDoc = parseXmlOpt(wb.files["xl/workbook.xml"]);
  const relsDoc = parseXmlOpt(wb.files["xl/_rels/workbook.xml.rels"]);
  if (!wbDoc || !relsDoc) throw new Error("workbook.xml or its relationships are missing");

  // Relationship id, unique within the workbook rels.
  const rids = new Set(Array.from(relsDoc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let rn = 1;
  while (rids.has(`rId${rn}`)) rn++;
  const rid = `rId${rn}`;
  const relEl = relsDoc.createElementNS(PKG_REL_NS, "Relationship");
  relEl.setAttribute("Id", rid);
  relEl.setAttribute("Type", `${REL_NS}/worksheet`);
  relEl.setAttribute("Target", rel);
  relsDoc.documentElement.appendChild(relEl);
  wb.files["xl/_rels/workbook.xml.rels"] = serializeXml(relsDoc);

  // <sheet> entry with a unique sheetId.
  const ids = Array.from(wbDoc.getElementsByTagName("sheet")).map((e) => Number(e.getAttribute("sheetId") ?? "0"));
  const sheetId = (ids.length ? Math.max(...ids) : 0) + 1;
  let sheetsEl = wbDoc.getElementsByTagName("sheets")[0];
  if (!sheetsEl) {
    sheetsEl = wbDoc.createElementNS(SS_MAIN, "sheets");
    wbDoc.documentElement.appendChild(sheetsEl);
  }
  const sheetEl = wbDoc.createElementNS(SS_MAIN, "sheet");
  sheetEl.setAttribute("name", name);
  sheetEl.setAttribute("sheetId", String(sheetId));
  sheetEl.setAttribute("r:id", rid);
  sheetsEl.appendChild(sheetEl);
  wb.files["xl/workbook.xml"] = serializeXml(wbDoc);

  // Content-type override so the package declares the new part.
  const ct = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (ct && ct.documentElement.localName === "Types") {
    const ov = ct.createElementNS(CT_NS, "Override");
    ov.setAttribute("PartName", `/${path}`);
    ov.setAttribute("ContentType", CT_WORKSHEET);
    ct.documentElement.appendChild(ov);
    wb.files["[Content_Types].xml"] = serializeXml(ct);
  }

  // A sheet added after the file was read cannot take its id from a position both peers
  // agree on, so it gets a random one; see assignSheetIds.
  const sheet: Sheet = { name, cid: newSheetId(), cells: new Map(), maxRow: 0, maxCol: 0, doc, sheetData, path, layoutDirty: true };
  wb.files[path] = serializeXml(doc);
  wb.sheets.push(sheet);
  return sheet;
}
