import type { Sheet, Workbook } from "../../core/model";
import { colToLetters, numToStr, parseXmlOpt, serializeXml } from "../../core/model";
import { pivotValueLabel, type AxisNode, type PivotComputed, type PivotSpec } from "../../core/pivot";

// xlsx pivot authoring: emit the pivotCacheDefinition + pivotCacheRecords + pivotTable parts, wire
// them into the package (rels, [Content_Types].xml, workbook <pivotCaches>, the host worksheet's
// rels), and set refreshOnLoad so Excel rebuilds the pivot body from the worksheet source on open.
// The output cells themselves are placed by the caller (format-agnostic). Structure mirrors what
// LibreOffice emits, which is Excel-compatible; the cache/records are correct for a non-refreshing
// reader while the field roles + source drive both apps' rebuild.

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const OREL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function addContentType(wb: Workbook, partPath: string, ct: string): void {
  const doc = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (!doc || doc.documentElement.localName !== "Types") return;
  if (Array.from(doc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === `/${partPath}`)) return;
  const ov = doc.createElementNS(CT_NS, "Override");
  ov.setAttribute("PartName", `/${partPath}`);
  ov.setAttribute("ContentType", ct);
  doc.documentElement.appendChild(ov);
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

function addRel(wb: Workbook, relsPath: string, type: string, target: string): string {
  let doc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (!doc) doc = parseXmlOpt(new TextEncoder().encode(`<Relationships xmlns="${REL_NS}"></Relationships>`))!;
  const ids = new Set(Array.from(doc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let n = 1; while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const rel = doc.createElementNS(REL_NS, "Relationship");
  rel.setAttribute("Id", id); rel.setAttribute("Type", type); rel.setAttribute("Target", target);
  doc.documentElement.appendChild(rel);
  wb.files[relsPath] = serializeXml(doc);
  return id;
}

const bytes = (s: string): Uint8Array => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`);

// Encode a row/column axis into rowItems/colItems. Leaves delta-encode (@r = leading levels
// unchanged from the previous leaf, then an <x> per changed level); subtotal lines emit their full
// prefix with t="default"; the grand line is t="grand". Matches LibreOffice's Excel-compatible form.
function axisXml(axis: AxisNode[]): string {
  let prevLeaf: number[] = [];
  let out = "";
  const xsFrom = (key: number[], start: number): string => {
    let xs = "";
    for (let k = start; k < key.length; k++) xs += key[k] === 0 ? "<x/>" : `<x v="${key[k]}"/>`;
    return xs || "<x/>";
  };
  for (const n of axis) {
    if (n.kind === "grand") { out += `<i t="grand"><x/></i>`; continue; }
    if (n.kind === "subtotal") { out += `<i t="default">${xsFrom(n.key, 0)}</i>`; continue; }
    let r = 0; while (r < n.key.length && r < prevLeaf.length && n.key[r] === prevLeaf[r]) r++;
    out += `<i${r > 0 ? ` r="${r}"` : ""}>${xsFrom(n.key, r)}</i>`;
    prevLeaf = n.key;
  }
  return out;
}

function nextCacheId(wb: Workbook): number {
  const doc = parseXmlOpt(wb.files["xl/workbook.xml"]);
  let max = 0;
  if (doc) for (const pc of Array.from(doc.getElementsByTagName("pivotCache"))) max = Math.max(max, Number(pc.getAttribute("cacheId") || "0"));
  return max + 1;
}

/** Append <pivotCache cacheId r:id> into workbook.xml's <pivotCaches> (created + positioned after
    <calcPr>, before <extLst>, per the CT_Workbook child order). */
function addWorkbookPivotCache(wb: Workbook, cacheId: number, relId: string): void {
  const doc = parseXmlOpt(wb.files["xl/workbook.xml"]);
  if (!doc) return;
  const root = doc.documentElement;
  let caches = Array.from(root.children).find((e) => e.localName === "pivotCaches");
  if (!caches) {
    caches = doc.createElementNS(MAIN, "pivotCaches");
    const extLst = Array.from(root.children).find((e) => e.localName === "extLst");
    const calcPr = Array.from(root.children).find((e) => e.localName === "calcPr");
    root.insertBefore(caches, extLst ?? (calcPr ? calcPr.nextSibling : null));
  }
  const pc = doc.createElementNS(MAIN, "pivotCache");
  pc.setAttribute("cacheId", String(cacheId));
  pc.setAttributeNS(OREL, "r:id", relId);
  caches.appendChild(pc);
  wb.files["xl/workbook.xml"] = serializeXml(doc);
}

/** Emit and wire the pivot parts for a pivot whose output is anchored at (anchorRow, anchorCol) on
    destSheet, sourcing sourceSheetName's data described by spec/computed. */
export function writeXlsxPivotParts(
  wb: Workbook,
  destSheet: Sheet,
  anchor: { row: number; col: number },
  sourceSheetName: string,
  spec: PivotSpec,
  computed: PivotComputed,
): { part: string; cachePart: string } {
  const width = spec.source.c2 - spec.source.c1 + 1;
  const cacheId = nextCacheId(wb);
  // Unique part numbers (cache def + records share n; table uses m).
  let n = 1; while (wb.files[`xl/pivotCache/pivotCacheDefinition${n}.xml`] || wb.files[`xl/pivotCache/pivotCacheRecords${n}.xml`]) n++;
  let m = 1; while (wb.files[`xl/pivotTables/pivotTable${m}.xml`]) m++;
  const cacheDefPath = `xl/pivotCache/pivotCacheDefinition${n}.xml`;
  const recordsPath = `xl/pivotCache/pivotCacheRecords${n}.xml`;
  const tablePath = `xl/pivotTables/pivotTable${m}.xml`;

  const pages = spec.pages ?? [];
  const isRow = (c: number) => spec.rows.includes(c);
  const isCol = (c: number) => spec.cols.includes(c);
  const isPage = (c: number) => pages.some((p) => p.field === c);
  const isGroup = (c: number) => isRow(c) || isCol(c) || isPage(c);
  const isVal = (c: number) => spec.values.some((v) => v.calc == null && v.field === c);
  const sub = !!spec.subtotals;
  // Calculated fields become extra cacheFields (databaseField="0" + a formula), indexed after the
  // source columns; each value field's `fld` resolves to its source column or its calc cacheField.
  let calcN = 0;
  const fldOf = spec.values.map((v) => (v.calc != null ? width + calcN++ : v.field ?? 0));
  const calcFieldsXml = spec.values.filter((v) => v.calc != null).map((v, i) => `<cacheField name="${esc(v.name || `Calc${i + 1}`)}" databaseField="0" numFmtId="0" formula="${esc(v.calc!)}"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"/></cacheField>`).join("");
  const totalFields = width + calcN;
  // Calculated items: synthetic members of a row/column field (a formula over that field's items).
  // They ride in the field's sharedItems + pivotField <item f="1">, with the formula in
  // <calculatedItems>. The item's index within the field = its position after the real items.
  const calcItems = spec.calcItems ?? [];
  const fieldRealCount = (c: number): number => computed.fields[c]!.items.length - calcItems.filter((ci) => ci.field === c).length;
  const calcItemIndex = new Map<(typeof calcItems)[number], number>();
  { const per = new Map<number, number>(); for (const ci of calcItems) { const k = per.get(ci.field) ?? 0; calcItemIndex.set(ci, fieldRealCount(ci.field) + k); per.set(ci.field, k + 1); } }
  // "Show values as": the OOXML showDataAs value + a percent number format / running-total base.
  const showAsAttr = (v: (typeof spec.values)[number]): string => {
    switch (v.showAs) {
      case "percentOfTotal": return ' showDataAs="percentOfTotal" numFmtId="10"';
      case "percentOfCol": return ' showDataAs="percentOfCol" numFmtId="10"';
      case "percentOfRow": return ' showDataAs="percentOfRow" numFmtId="10"';
      case "runningTotal": return ` showDataAs="runTotal" baseField="${spec.rows[0] ?? 0}"`;
      default: return "";
    }
  };

  // --- cacheDefinition: cacheFields (sharedItems for grouping/page fields; type flags otherwise) ---
  let cacheFields = "";
  for (let c = 0; c < width; c++) {
    const f = computed.fields[c]!;
    if (isGroup(c)) {
      const items = f.items.map((it) => (it.num ? `<n v="${numToStr(it.value as number)}"/>` : `<s v="${esc(String(it.value))}"/>`)).join("");
      cacheFields += `<cacheField name="${esc(f.name)}" numFmtId="0"><sharedItems count="${f.items.length}">${items}</sharedItems></cacheField>`;
    } else {
      let hasNum = false, hasStr = false, hasBlank = false;
      for (const rec of computed.records) { const cv = rec.cells[c]!; if (cv.value === null) hasBlank = true; else if (cv.num) hasNum = true; else hasStr = true; }
      const flags = `containsSemiMixedTypes="${hasStr || hasBlank ? 1 : 0}" containsString="${hasStr ? 1 : 0}" containsNumber="${hasNum ? 1 : 0}"${hasNum && !hasStr ? ' containsBlank="' + (hasBlank ? 1 : 0) + '"' : ""}`;
      cacheFields += `<cacheField name="${esc(f.name)}" numFmtId="0"><sharedItems ${flags}/></cacheField>`;
    }
  }
  const cacheDef = `<pivotCacheDefinition xmlns="${MAIN}" xmlns:r="${OREL}" r:id="rId1" refreshOnLoad="1" refreshedBy="sheetedit" recordCount="${computed.records.length}" createdVersion="3">`
    + `<cacheSource type="worksheet"><worksheetSource ref="${rangeRef(spec.source)}" sheet="${esc(sourceSheetName)}"/></cacheSource>`
    + `<cacheFields count="${totalFields}">${cacheFields}${calcFieldsXml}</cacheFields>`
    + (calcItems.length ? `<calculatedItems count="${calcItems.length}">` + calcItems.map((ci) => `<calculatedItem field="${ci.field}" formula="${esc(ci.formula)}"><pivotArea outline="0" fieldPosition="0"><references count="1"><reference field="${ci.field}" count="1"><x v="${calcItemIndex.get(ci)}"/></reference></references></pivotArea></calculatedItem>`).join("") + `</calculatedItems>` : "")
    + `</pivotCacheDefinition>`;

  // --- cacheRecords: <x v> for grouping fields, literal <n>/<s>/<m/> otherwise ---
  let recs = "";
  for (const rec of computed.records) {
    let r = "";
    for (let c = 0; c < width; c++) {
      const cv = rec.cells[c]!;
      const f = computed.fields[c]!;
      if (isGroup(c)) {
        const v = cv.value === null ? "(empty)" : cv.value;
        r += `<x v="${f.indexOf.get((typeof v === "number" ? "n:" : "s:") + v) ?? 0}"/>`;
      } else if (cv.value === null) r += "<m/>";
      else if (cv.num) r += `<n v="${numToStr(cv.value as number)}"/>`;
      else r += `<s v="${esc(String(cv.value))}"/>`;
    }
    recs += `<r>${r}</r>`;
  }
  const records = `<pivotCacheRecords xmlns="${MAIN}" xmlns:r="${OREL}" count="${computed.records.length}">${recs}</pivotCacheRecords>`;

  // --- pivotTable ---
  let pivotFields = "";
  for (let c = 0; c < width; c++) {
    const f = computed.fields[c]!;
    if (isRow(c) || isCol(c)) {
      // With subtotals on, drop the defaultSubtotal="0" (default is on) and add the subtotal item.
      // Items beyond the real count are calculated items, flagged f="1".
      const realN = fieldRealCount(c);
      const items = f.items.map((_, i) => (i >= realN ? `<item x="${i}" f="1"/>` : `<item x="${i}"/>`)).join("") + (sub ? `<item t="default"/>` : "");
      pivotFields += `<pivotField axis="${isRow(c) ? "axisRow" : "axisCol"}" compact="0" showAll="0"${sub ? "" : ' defaultSubtotal="0"'}><items count="${f.items.length + (sub ? 1 : 0)}">${items}</items></pivotField>`;
    } else if (isPage(c)) {
      const items = f.items.map((_, i) => `<item x="${i}"/>`).join("");
      pivotFields += `<pivotField axis="axisPage" compact="0" showAll="0" defaultSubtotal="0"><items count="${f.items.length}">${items}</items></pivotField>`;
    } else if (isVal(c)) pivotFields += `<pivotField dataField="1" compact="0" showAll="0"/>`;
    else pivotFields += `<pivotField compact="0" showAll="0"/>`;
  }
  // A pivotField per calculated field (they sit after the source columns as data fields).
  for (let i = 0; i < calcN; i++) pivotFields += `<pivotField dataField="1" compact="0" showAll="0"/>`;
  const R = spec.rows.length, C = spec.cols.length;
  const rowFields = R ? `<rowFields count="${R}">${spec.rows.map((c) => `<field x="${c}"/>`).join("")}</rowFields>` : "";
  const rowItems = `<rowItems count="${computed.rowAxis.length}">${axisXml(computed.rowAxis)}</rowItems>`;
  const colFields = C ? `<colFields count="${C}">${spec.cols.map((c) => `<field x="${c}"/>`).join("")}</colFields>` : "";
  const colItems = C ? `<colItems count="${computed.colAxis.length}">${axisXml(computed.colAxis)}</colItems>` : `<colItems count="1"><i/></colItems>`;
  const pageFields = pages.length ? `<pageFields count="${pages.length}">${pages.map((p) => `<pageField fld="${p.field}"${p.item != null ? ` item="${p.item}"` : ""} hier="-1"/>`).join("")}</pageFields>` : "";
  const dataFields = `<dataFields count="${spec.values.length}">`
    + spec.values.map((v, vi) => {
      const name = v.calc != null ? (v.name || "Calc") : pivotValueLabel(v.func ?? "sum", computed.fields[v.field ?? 0]!.name);
      const subtotal = v.calc != null ? "sum" : (v.func ?? "sum");
      return `<dataField name="${esc(name)}" fld="${fldOf[vi]}" subtotal="${subtotal}"${showAsAttr(v)}/>`;
    }).join("")
    + `</dataFields>`;
  const loc = { r1: anchor.row, c1: anchor.col, r2: anchor.row + computed.height - 1, c2: anchor.col + computed.width - 1 };
  const pageCounts = pages.length ? ` rowPageCount="1" colPageCount="1"` : "";
  const location = `<location ref="${rangeRef(loc)}" firstHeaderRow="1" firstDataRow="${computed.headerRows}" firstDataCol="${computed.headerCols}"${pageCounts}/>`;
  const table = `<pivotTableDefinition xmlns="${MAIN}" name="${esc(destSheet.name === "" ? "PivotTable" : "PivotTable" + m)}" cacheId="${cacheId}" dataOnRows="0" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0" dataCaption="Values" showDrill="0" useAutoFormatting="0" itemPrintTitles="1" indent="0" outline="1" outlineData="1" compact="1" compactData="1">`
    + location + `<pivotFields count="${totalFields}">${pivotFields}</pivotFields>` + rowFields + rowItems + colFields + colItems + pageFields + dataFields
    + `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;

  // --- write parts + wiring ---
  wb.files[cacheDefPath] = bytes(cacheDef);
  wb.files[recordsPath] = bytes(records);
  wb.files[tablePath] = bytes(table);
  wb.files[`xl/pivotCache/_rels/pivotCacheDefinition${n}.xml.rels`] = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OREL}/pivotCacheRecords" Target="pivotCacheRecords${n}.xml"/></Relationships>`);
  wb.files[`xl/pivotTables/_rels/pivotTable${m}.xml.rels`] = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OREL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${n}.xml"/></Relationships>`);
  addContentType(wb, tablePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
  addContentType(wb, cacheDefPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
  addContentType(wb, recordsPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
  // Host worksheet -> pivotTable rel.
  const wsRels = destSheet.path!.replace(/([^/]+)$/, "_rels/$1.rels");
  addRel(wb, wsRels, `${OREL}/pivotTable`, `../pivotTables/pivotTable${m}.xml`);
  // Workbook -> pivotCacheDefinition rel, then the <pivotCache> entry.
  const wbRelId = addRel(wb, "xl/_rels/workbook.xml.rels", `${OREL}/pivotCacheDefinition`, `pivotCache/pivotCacheDefinition${n}.xml`);
  addWorkbookPivotCache(wb, cacheId, wbRelId);
  // Track the cache so a later edit to the source range re-flags refreshOnLoad (already set here).
  (wb.pivotCaches ??= []).push({ part: cacheDefPath, sourceSheet: sourceSheetName, source: { ...spec.source }, refreshFlagged: true });
  return { part: tablePath, cachePart: cacheDefPath };
}

/** Remove a pivot's parts and package wiring (its pivotTable + backing cache, rels, content-type
    overrides, the host worksheet rel, and the workbook <pivotCache>). Used when editing/deleting an
    authored pivot. The output cells are cleared by the caller. */
export function deleteXlsxPivotParts(wb: Workbook, destSheet: Sheet, part: string, cachePart: string): void {
  const rm = (p: string) => { delete wb.files[p]; };
  rm(part); rm(part.replace(/([^/]+)$/, "_rels/$1.rels"));
  removeContentType(wb, part);
  // Host worksheet rel -> this pivotTable.
  const wsRels = destSheet.path!.replace(/([^/]+)$/, "_rels/$1.rels");
  removeRelByTarget(wb, wsRels, /pivotTable/i);
  // Drop the cache only if no other pivotTable's rels still reference it (parts already removed).
  const cacheFile = cachePart.replace(/.*\//, ""); // pivotCacheDefinitionN.xml
  const otherUsesCache = Object.keys(wb.files).some((k) => /xl\/pivotTables\/_rels\/pivotTable\d+\.xml\.rels$/.test(k) && new TextDecoder().decode(wb.files[k]!).includes(cacheFile));
  if (!otherUsesCache) {
    rm(cachePart); rm(cachePart.replace(/([^/]+)$/, "_rels/$1.rels"));
    const recs = cachePart.replace(/pivotCacheDefinition/, "pivotCacheRecords");
    rm(recs);
    removeContentType(wb, cachePart); removeContentType(wb, recs);
    removeWorkbookPivotCache(wb, cachePart);
    if (wb.pivotCaches) wb.pivotCaches = wb.pivotCaches.filter((c) => c.part !== cachePart);
  }
}

function removeContentType(wb: Workbook, partPath: string): void {
  const doc = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (!doc) return;
  const ov = Array.from(doc.getElementsByTagName("Override")).find((o) => o.getAttribute("PartName") === `/${partPath}`);
  if (ov) { ov.parentNode?.removeChild(ov); wb.files["[Content_Types].xml"] = serializeXml(doc); }
}

function removeRelByTarget(wb: Workbook, relsPath: string, targetRe: RegExp): void {
  const doc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (!doc) return;
  let changed = false;
  for (const r of Array.from(doc.getElementsByTagName("Relationship"))) if (targetRe.test(r.getAttribute("Target") ?? "")) { r.parentNode?.removeChild(r); changed = true; }
  if (changed) wb.files[relsPath] = serializeXml(doc);
}

function removeWorkbookPivotCache(wb: Workbook, cachePart: string): void {
  const relsDoc = parseXmlOpt(wb.files["xl/_rels/workbook.xml.rels"]);
  const target = cachePart.replace(/^xl\//, "");
  let relId: string | undefined;
  if (relsDoc) {
    const rel = Array.from(relsDoc.getElementsByTagName("Relationship")).find((r) => (r.getAttribute("Target") ?? "").replace(/^\.\//, "") === target || (r.getAttribute("Target") ?? "") === cachePart);
    if (rel) { relId = rel.getAttribute("Id") ?? undefined; rel.parentNode?.removeChild(rel); wb.files["xl/_rels/workbook.xml.rels"] = serializeXml(relsDoc); }
  }
  const wbDoc = parseXmlOpt(wb.files["xl/workbook.xml"]);
  if (wbDoc && relId) {
    const pc = Array.from(wbDoc.getElementsByTagName("pivotCache")).find((p) => (p.getAttributeNS(OREL, "id") ?? p.getAttribute("r:id")) === relId);
    if (pc) { const parent = pc.parentNode as Element | null; pc.parentNode?.removeChild(pc); if (parent && parent.localName === "pivotCaches" && !parent.children.length) parent.parentNode?.removeChild(parent); wb.files["xl/workbook.xml"] = serializeXml(wbDoc); }
  }
}

function rangeRef(r: { r1: number; c1: number; r2: number; c2: number }): string {
  return `${colToLetters(r.c1)}${r.r1}:${colToLetters(r.c2)}${r.r2}`;
}
