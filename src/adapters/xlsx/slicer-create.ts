import { parseXmlOpt, serializeXml, type PivotTableInfo, type Sheet, type SheetSlicer, type Workbook } from "../../core/model";
import { pxToEmu } from "../../core/chart-model";
import { addContentType, addRel, ensureSheetDrawing } from "./chart-write";

// Create a slicer from scratch: the two parts (view + cache), their content types and
// relationships, the two extension-list registrations Excel looks for, and a drawing graphicFrame
// for the anchor. Values follow [MS-XLSX] (CT_Slicer / CT_SlicerCacheDefinition /
// CT_TabularSlicerCacheItem); the extension URIs and relationship types are the documented
// Office 2010 ones, cross-checked against an independent implementation (excelize).

const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SLE = "http://schemas.microsoft.com/office/drawing/2010/slicer";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CT_SLICER = "application/vnd.ms-excel.slicer+xml";
const CT_SLICER_CACHE = "application/vnd.ms-excel.slicerCache+xml";
const REL_SLICER = "http://schemas.microsoft.com/office/2007/relationships/slicer";
const REL_SLICER_CACHE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
// extLst URIs: x14:slicerCaches on the workbook, x14:slicerList on the worksheet.
const EXT_SLICER_CACHES = "{BBE1A952-AA13-448e-AADC-164F8A28A991}";
const EXT_SLICER_LIST = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";

const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
/** A slicer / cache name must be a valid defined name: no spaces or punctuation. */
const safeName = (s: string): string => (s.replace(/[^A-Za-z0-9_]/g, "_").replace(/^(\d)/, "_$1") || "Field");

/** The cacheId of the pivot cache backing this pivot, from workbook.xml's <pivotCaches>. */
function cacheIdOf(wb: Workbook, cachePart: string | undefined): number {
  if (!cachePart) return 1;
  const relsDoc = wb.files["xl/_rels/workbook.xml.rels"] ? parseXmlOpt(wb.files["xl/_rels/workbook.xml.rels"]) : undefined;
  const byId = new Map<string, string>();
  if (relsDoc) for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
    const parts: string[] = [];
    for (const seg of `xl/${r.getAttribute("Target")}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    byId.set(r.getAttribute("Id") ?? "", parts.join("/"));
  }
  const wbDoc = wb.files["xl/workbook.xml"] ? parseXmlOpt(wb.files["xl/workbook.xml"]) : undefined;
  if (wbDoc) for (const pc of Array.from(wbDoc.getElementsByTagName("pivotCache"))) {
    const rid = pc.getAttributeNS(R, "id") ?? pc.getAttribute("r:id") ?? "";
    if (byId.get(rid) === cachePart) return Number(pc.getAttribute("cacheId") || "1");
  }
  return 1;
}

/** Append a child into an element's <extLst>, creating the list and the <ext uri> wrapper. */
function addExt(doc: Document, parent: Element, uri: string, innerXml: string): void {
  let extLst = Array.from(parent.children).find((e) => e.localName === "extLst");
  if (!extLst) {
    extLst = doc.createElementNS(parent.namespaceURI || MAIN, "extLst");
    // extLst is the last child of workbook / worksheet per the schema.
    parent.appendChild(extLst);
  }
  // Reuse an existing ext with this uri when present, so a second slicer joins the same list.
  let ext = Array.from(extLst.children).find((e) => e.getAttribute("uri") === uri);
  if (!ext) {
    const frag = parseXmlOpt(enc(`<ext xmlns="${parent.namespaceURI || MAIN}" xmlns:x14="${X14}" xmlns:r="${R}" uri="${uri}">${innerXml}</ext>`));
    if (!frag) return;
    extLst.appendChild(doc.importNode(frag.documentElement, true));
    return;
  }
  // Merge into the existing x14 container (slicerCaches / slicerList).
  const frag = parseXmlOpt(enc(`<w xmlns:x14="${X14}" xmlns:r="${R}">${innerXml}</w>`));
  const incoming = frag?.documentElement.firstElementChild;
  if (!incoming) return;
  const container = Array.from(ext.children).find((e) => e.localName === incoming.localName);
  if (container) { for (const c of Array.from(incoming.children)) container.appendChild(doc.importNode(c, true)); }
  else ext.appendChild(doc.importNode(incoming, true));
}

/** The graphicFrame anchor Excel uses to place a slicer on the sheet. */
function slicerFrameXml(sl: SheetSlicer, id: number): string {
  const a = sl.anchor!;
  const pt = (tag: string, col: number, colOff: number, row: number, rowOff: number): string =>
    `<xdr:${tag}><xdr:col>${col - 1}</xdr:col><xdr:colOff>${pxToEmu(colOff)}</xdr:colOff><xdr:row>${row - 1}</xdr:row><xdr:rowOff>${pxToEmu(rowOff)}</xdr:rowOff></xdr:${tag}>`;
  return `<xdr:twoCellAnchor editAs="oneCell">${pt("from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff)}${pt("to", a.toCol, a.toColOff, a.toRow, a.toRowOff)}` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(sl.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${SLE}"><sle:slicer xmlns:sle="${SLE}" name="${esc(sl.name)}"/></a:graphicData></a:graphic>` +
    `</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

/**
 * Create a slicer on `sheet` for `field` of `info`'s pivot, with every item selected.
 * Returns the model (already pushed onto sheet.slicers) or null when it cannot be built.
 */
export function createXlsxSlicer(wb: Workbook, sheet: Sheet, info: PivotTableInfo, fieldName: string, items: string[], anchor: SheetSlicer["anchor"]): SheetSlicer | null {
  if (!sheet.path || !items.length) return null;
  const used = new Set(wb.sheets.flatMap((s) => (s.slicers ?? []).map((x) => x.name)));
  let name = safeName(fieldName);
  let n = 1;
  while (used.has(name)) name = `${safeName(fieldName)}${++n}`;
  const cacheName = `Slicer_${name}`;

  // --- the cache part: the selection lives here ---
  let ci = 1;
  while (wb.files[`xl/slicerCaches/slicerCache${ci}.xml`]) ci++;
  const cachePath = `xl/slicerCaches/slicerCache${ci}.xml`;
  const itemsXml = items.map((_, i) => `<x14:i x="${i}" s="1"/>`).join("");
  wb.files[cachePath] = enc(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<x14:slicerCacheDefinition xmlns:x14="${X14}" xmlns:r="${R}" name="${esc(cacheName)}" sourceName="${esc(fieldName)}">` +
      `<x14:pivotTables><x14:pivotTable tabId="1" name="${esc(info.name)}"/></x14:pivotTables>` +
      `<x14:data><x14:tabular pivotCacheId="${cacheIdOf(wb, info.cachePart)}" crossFilter="showItemsWithDataAtTop">` +
        `<x14:items count="${items.length}">${itemsXml}</x14:items>` +
      `</x14:tabular></x14:data>` +
    `</x14:slicerCacheDefinition>`);
  addContentType(wb, cachePath, CT_SLICER_CACHE);
  // The cache is owned by the workbook and registered in its extLst.
  const cacheRid = addRel(wb, "xl/_rels/workbook.xml.rels", REL_SLICER_CACHE, `slicerCaches/slicerCache${ci}.xml`);
  const wbDoc = wb.files["xl/workbook.xml"] ? parseXmlOpt(wb.files["xl/workbook.xml"]) : undefined;
  if (wbDoc) {
    addExt(wbDoc, wbDoc.documentElement, EXT_SLICER_CACHES, `<x14:slicerCaches xmlns:x14="${X14}" xmlns:r="${R}"><x14:slicerCache r:id="${cacheRid}"/></x14:slicerCaches>`);
    wb.files["xl/workbook.xml"] = serializeXml(wbDoc);
  }

  // --- the view part: one <x14:slicers> per sheet, holding this sheet's slicer views ---
  const sheetRels = sheet.path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const relsDoc = wb.files[sheetRels] ? parseXmlOpt(wb.files[sheetRels]) : undefined;
  const existing = relsDoc && Array.from(relsDoc.getElementsByTagName("Relationship")).find((r) => r.getAttribute("Type") === REL_SLICER);
  let slicerPath: string;
  let slicerRid: string | undefined;
  const viewXml = `<x14:slicer name="${esc(name)}" cache="${esc(cacheName)}" caption="${esc(fieldName)}" columnCount="1" rowHeight="234950"/>`;
  if (existing) {
    const parts: string[] = [];
    for (const seg of `xl/worksheets/${existing.getAttribute("Target")}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    slicerPath = parts.join("/");
    const xml = new TextDecoder().decode(wb.files[slicerPath] ?? enc(`<x14:slicers xmlns:x14="${X14}"></x14:slicers>`));
    wb.files[slicerPath] = enc(xml.replace(/<\/x14:slicers>\s*$/, `${viewXml}</x14:slicers>`));
  } else {
    let si = 1;
    while (wb.files[`xl/slicers/slicer${si}.xml`]) si++;
    slicerPath = `xl/slicers/slicer${si}.xml`;
    wb.files[slicerPath] = enc(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<x14:slicers xmlns:x14="${X14}" xmlns:r="${R}">${viewXml}</x14:slicers>`);
    addContentType(wb, slicerPath, CT_SLICER);
    slicerRid = addRel(wb, sheetRels, REL_SLICER, `../slicers/slicer${si}.xml`);
  }

  // --- the drawing anchor ---
  const sl: SheetSlicer = {
    name, cache: cacheName, caption: fieldName, columnCount: 1,
    sourceName: fieldName, pivotTables: [info.name],
    items: items.map((label, x) => ({ x, label, selected: true })),
    anchor, slicerPath, cachePath,
  };
  if (anchor) {
    const drawingPath = ensureSheetDrawing(wb, sheet);
    const xml = new TextDecoder().decode(wb.files[drawingPath]);
    const id = 1000 + (xml.match(/<xdr:(twoCell|oneCell|absolute)Anchor/g)?.length ?? 0);
    wb.files[drawingPath] = enc(xml.replace(/<\/xdr:wsDr>\s*$/, `${slicerFrameXml(sl, id)}</xdr:wsDr>`));
  }
  // The worksheet extLst must be added AFTER any drawing work: ensureSheetDrawing re-parses the
  // sheet from wb.files and replaces sheet.doc, which would drop an earlier in-memory edit.
  if (slicerRid && sheet.doc) {
    addExt(sheet.doc, sheet.doc.documentElement, EXT_SLICER_LIST, `<x14:slicerList xmlns:x14="${X14}" xmlns:r="${R}"><x14:slicer r:id="${slicerRid}"/></x14:slicerList>`);
    sheet.layoutDirty = true;
  }
  (sheet.slicers ??= []).push(sl);
  return sl;
}
