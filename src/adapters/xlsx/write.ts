import type { Cell, DataValidation, Sheet, Workbook } from "../../core/model";
import { colToLetters, ensureCell, firstByLocal, parseXmlOpt, removeByLocal, serializeXml } from "../../core/model";
import { SS_MAIN, ensureXlsxCellEl } from "./shared";
import { writeXlsxCharts } from "./chart-write";
import { setXlsxCellNumFmt } from "./styles";
// ---------------------------------------------------------------------------
// xlsx write: surgical cell/layout writers and the save pass
// ---------------------------------------------------------------------------

export function writeXlsxCell(sheet: Sheet, cell: Cell, plainFormula = false): void {
  const doc = sheet.doc!;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const c = ensureXlsxCellEl(sheet, cell);
  // When only the cached value changed, keep the original <f> untouched so
  // t="shared"/"array", @si and @ref survive. A formula edit (or a group
  // de-share) rewrites <f> as a plain formula instead.
  const oldF = firstByLocal(c, "f");
  const keepF = cell.formula != null && oldF != null && !cell.fDirty && !plainFormula;
  if (!keepF) removeByLocal(c, "f");
  removeByLocal(c, "v");
  removeByLocal(c, "is");
  const addV = (text: string) => {
    const v = doc.createElementNS(ns, "v");
    v.textContent = text;
    c.appendChild(v);
  };
  if (cell.formula != null) {
    if (!keepF) {
      const f = doc.createElementNS(ns, "f");
      f.textContent = cell.formula;
      c.appendChild(f);
    }
    if (cell.kind === "n") {
      c.removeAttribute("t");
      if (cell.value !== "") addV(cell.value);
    } else if (cell.kind === "b") {
      c.setAttribute("t", "b");
      addV(cell.value === "TRUE" || cell.value === "1" ? "1" : "0");
    } else if (cell.kind === "e") {
      c.setAttribute("t", "e");
      addV(cell.value);
    } else if (cell.kind === "blank" || cell.value === "") {
      c.removeAttribute("t");
    } else {
      c.setAttribute("t", "str");
      addV(cell.value);
    }
    return;
  }
  // literal
  if (cell.value === "" || cell.kind === "blank") {
    c.removeAttribute("t");
  } else if (cell.kind === "n") {
    c.removeAttribute("t");
    addV(cell.value);
  } else if (cell.kind === "b") {
    c.setAttribute("t", "b");
    addV(cell.value === "TRUE" || cell.value === "1" ? "1" : "0");
  } else if (cell.kind === "e") {
    c.setAttribute("t", "e");
    addV(cell.value);
  } else {
    c.setAttribute("t", "inlineStr");
    const is = doc.createElementNS(ns, "is");
    const t = doc.createElementNS(ns, "t");
    t.setAttribute("xml:space", "preserve");
    t.textContent = cell.value;
    is.appendChild(t);
    // Furigana: emit the phonetic guide as <rPh> runs after the base text.
    for (const p of cell.phonetic ?? []) {
      if (!p.reading) continue;
      const rPh = doc.createElementNS(ns, "rPh");
      rPh.setAttribute("sb", String(p.sb));
      rPh.setAttribute("eb", String(p.eb));
      const rt = doc.createElementNS(ns, "t");
      rt.textContent = p.reading;
      rPh.appendChild(rt);
      is.appendChild(rPh);
    }
    if (cell.phonetic?.length) is.appendChild(doc.createElementNS(ns, "phoneticPr")); // hints Excel to show it
    c.appendChild(is);
  }
}

// Set a single column's width (px) in the worksheet's <cols>, creating <cols>/<col>
// as needed and splitting any existing run that covers the column. Keeps colWidths in sync.
export function setXlsxColWidth(sheet: Sheet, col: number, px: number): void {
  if (!sheet.colWidths) sheet.colWidths = new Map();
  sheet.colWidths.set(col, px);
  const doc = sheet.doc;
  if (!doc) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const width = Math.max(0, (px - 5) / 7);
  let colsEl = doc.getElementsByTagName("cols")[0] as Element | undefined;
  if (!colsEl) {
    colsEl = doc.createElementNS(ns, "cols");
    // <cols> must precede <sheetData> per the schema.
    sheet.sheetData?.parentNode?.insertBefore(colsEl, sheet.sheetData);
  }
  // Narrow any run that spans `col` so we can give `col` its own entry.
  for (const c of Array.from(colsEl.children)) {
    if (c.localName !== "col") continue;
    const min = Number(c.getAttribute("min") || "0");
    const max = Number(c.getAttribute("max") || String(min));
    if (col < min || col > max) continue;
    if (min === max) {
      c.setAttribute("width", String(width));
      c.setAttribute("customWidth", "1");
      sheet.layoutDirty = true;
      return;
    }
    // Split: left part [min..col-1], right part [col+1..max], plus the single col.
    if (col > min) {
      const left = c.cloneNode(true) as Element;
      left.setAttribute("min", String(min));
      left.setAttribute("max", String(col - 1));
      colsEl.insertBefore(left, c);
    }
    if (col < max) {
      const right = c.cloneNode(true) as Element;
      right.setAttribute("min", String(col + 1));
      right.setAttribute("max", String(max));
      colsEl.insertBefore(right, c);
    }
    c.setAttribute("min", String(col));
    c.setAttribute("max", String(col));
    c.setAttribute("width", String(width));
    c.setAttribute("customWidth", "1");
    sheet.layoutDirty = true;
    return;
  }
  const colEl = doc.createElementNS(ns, "col");
  colEl.setAttribute("min", String(col));
  colEl.setAttribute("max", String(col));
  colEl.setAttribute("width", String(width));
  colEl.setAttribute("customWidth", "1");
  colsEl.appendChild(colEl);
  sheet.layoutDirty = true;
}

// Set a single row's height (px) on its <row>, creating the row element if absent.
export function setXlsxRowHeight(sheet: Sheet, row: number, px: number): void {
  if (!sheet.rowHeights) sheet.rowHeights = new Map();
  sheet.rowHeights.set(row, px);
  const doc = sheet.doc;
  const sd = sheet.sheetData;
  if (!doc || !sd) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const pt = (px * 3) / 4;
  let rowEl: Element | undefined;
  for (const re of Array.from(sd.children)) {
    if (re.localName === "row" && Number(re.getAttribute("r") || "0") === row) {
      rowEl = re;
      break;
    }
  }
  if (!rowEl) {
    rowEl = doc.createElementNS(ns, "row");
    rowEl.setAttribute("r", String(row));
    // Insert keeping rows in ascending order.
    let next: Element | null = null;
    for (const re of Array.from(sd.children)) {
      if (re.localName === "row" && Number(re.getAttribute("r") || "0") > row) {
        next = re;
        break;
      }
    }
    sd.insertBefore(rowEl, next);
  }
  rowEl.setAttribute("ht", String(pt));
  rowEl.setAttribute("customHeight", "1");
  sheet.layoutDirty = true;
}

/** Find or create a <row r> element in ascending order. */
function ensureRowEl(sheet: Sheet, row: number): Element | undefined {
  const doc = sheet.doc, sd = sheet.sheetData;
  if (!doc || !sd) return undefined;
  for (const re of Array.from(sd.children)) if (re.localName === "row" && Number(re.getAttribute("r") || "0") === row) return re;
  const rowEl = doc.createElementNS(doc.documentElement.namespaceURI || SS_MAIN, "row");
  rowEl.setAttribute("r", String(row));
  let next: Element | null = null;
  for (const re of Array.from(sd.children)) if (re.localName === "row" && Number(re.getAttribute("r") || "0") > row) { next = re; break; }
  sd.insertBefore(rowEl, next);
  return rowEl;
}

/** Set/clear the hidden attribute on a row (used by filtering). */
export function setXlsxRowHidden(sheet: Sheet, row: number, hidden: boolean): void {
  const rowEl = ensureRowEl(sheet, row);
  if (!rowEl) return;
  if (hidden) rowEl.setAttribute("hidden", "1");
  else rowEl.removeAttribute("hidden");
  sheet.layoutDirty = true;
}

/** Set or remove the sheet-level <autoFilter ref>. */
export function setXlsxAutoFilter(sheet: Sheet, ref: string | null): void {
  const doc = sheet.doc;
  if (!doc) return;
  const ws = doc.documentElement;
  let af = ws.getElementsByTagName("autoFilter")[0];
  if (!ref) { if (af) af.parentNode?.removeChild(af); sheet.layoutDirty = true; return; }
  if (!af) {
    af = doc.createElementNS(ws.namespaceURI || SS_MAIN, "autoFilter");
    // autoFilter must sit after sheetData (and mergeCells) per the schema; append near the end.
    const sd = sheet.sheetData;
    const merges = ws.getElementsByTagName("mergeCells")[0];
    ws.insertBefore(af, (merges?.nextSibling ?? sd?.nextSibling) ?? null);
  }
  af.setAttribute("ref", ref);
  sheet.layoutDirty = true;
}

const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
/** Insert a worksheet child before the first "later" element (pageMargins, drawing, ...), keeping
    a roughly schema-valid order; else append. */
function insertWsChild(ws: Element, el: Element): void {
  const later = new Set(["sheetCalcPr", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "drawingHF", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]);
  let before: ChildNode | null = null;
  for (const ch of Array.from(ws.children)) if (later.has(ch.localName)) { before = ch; break; }
  ws.insertBefore(el, before);
}

/** Add or remove a list data validation over the given ranges (1-based inclusive). */
export function setXlsxDataValidation(sheet: Sheet, ranges: { r1: number; c1: number; r2: number; c2: number }[], spec: { values?: string[]; rangeRef?: string; allowBlank?: boolean } | null): void {
  const doc = sheet.doc, ws = doc?.documentElement;
  const ns = ws?.namespaceURI || SS_MAIN;
  const sqref = ranges.map((r) => `${colToLetters(r.c1)}${r.r1}:${colToLetters(r.c2)}${r.r2}`).join(" ");
  if (doc && ws) {
    let dvs = ws.getElementsByTagName("dataValidations")[0];
    if (dvs) for (const dv of Array.from(dvs.getElementsByTagName("dataValidation"))) if (dv.getAttribute("sqref") === sqref) dv.parentNode?.removeChild(dv);
    if (spec) {
      if (!dvs) { dvs = doc.createElementNS(ns, "dataValidations"); insertWsChild(ws, dvs); }
      const dv = doc.createElementNS(ns, "dataValidation");
      dv.setAttribute("type", "list");
      dv.setAttribute("allowBlank", spec.allowBlank ? "1" : "0");
      dv.setAttribute("showInputMessage", "1");
      dv.setAttribute("showErrorMessage", "1");
      dv.setAttribute("sqref", sqref);
      const f1 = doc.createElementNS(ns, "formula1");
      f1.textContent = spec.rangeRef ? spec.rangeRef : `"${(spec.values ?? []).join(",")}"`;
      dv.appendChild(f1);
      dvs.appendChild(dv);
      dvs.setAttribute("count", String(dvs.getElementsByTagName("dataValidation").length));
    } else if (dvs && !dvs.getElementsByTagName("dataValidation").length) dvs.parentNode?.removeChild(dvs);
  }
  // Keep the in-memory validations (drive the dropdown) in sync.
  sheet.validations = (sheet.validations ?? []).filter((v) => v.ranges.map((r) => `${colToLetters(r.c1)}${r.r1}:${colToLetters(r.c2)}${r.r2}`).join(" ") !== sqref);
  if (spec) { const v: DataValidation = { ranges, allowBlank: spec.allowBlank }; if (spec.rangeRef) v.rangeRef = spec.rangeRef; else v.values = spec.values; sheet.validations.push(v); }
  sheet.layoutDirty = true;
}

/** Set or remove a hyperlink on a cell. External links get a TargetMode="External" sheet rel;
    internal links use @location. */
export function setXlsxHyperlink(wb: Workbook, sheet: Sheet, r: number, c: number, link: { href: string; internal?: boolean; tip?: string } | null): void {
  const cell = ensureCell(sheet, r, c);
  if (link) cell.link = link; else delete cell.link;
  const doc = sheet.doc, ws = doc?.documentElement;
  if (!doc || !ws) { sheet.layoutDirty = true; return; }
  const ns = ws.namespaceURI || SS_MAIN;
  const ref = `${colToLetters(c)}${r}`;
  let hls = ws.getElementsByTagName("hyperlinks")[0];
  if (hls) for (const h of Array.from(hls.getElementsByTagName("hyperlink"))) if (h.getAttribute("ref") === ref) h.parentNode?.removeChild(h);
  if (link) {
    if (!hls) { hls = doc.createElementNS(ns, "hyperlinks"); insertWsChild(ws, hls); }
    const h = doc.createElementNS(ns, "hyperlink");
    h.setAttribute("ref", ref);
    if (link.internal) {
      h.setAttribute("location", link.href);
    } else {
      const rid = addExternalRel(wb, sheet, link.href);
      if (!ws.getAttribute("xmlns:r")) ws.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
      h.setAttribute("r:id", rid);
    }
    if (link.tip) h.setAttribute("tooltip", link.tip);
    hls.appendChild(h);
  } else if (hls && !hls.getElementsByTagName("hyperlink").length) hls.parentNode?.removeChild(hls);
  sheet.layoutDirty = true;
}

/** Add an external (TargetMode="External") relationship to the sheet's rels part, returning its id. */
function addExternalRel(wb: Workbook, sheet: Sheet, target: string): string {
  const relsPath = (sheet.path ?? "").replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  let doc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (!doc) doc = parseXmlOpt(new TextEncoder().encode(`<Relationships xmlns="${REL_NS}"></Relationships>`))!;
  const ids = new Set(Array.from(doc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let n = 1; while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const rel = doc.createElementNS(REL_NS, "Relationship");
  rel.setAttribute("Id", id);
  rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
  rel.setAttribute("Target", target);
  rel.setAttribute("TargetMode", "External");
  doc.documentElement.appendChild(rel);
  wb.files[relsPath] = serializeXml(doc);
  return id;
}

// Add or remove a merged range (1-based, inclusive). The top-left cell shows through;
// any cells the merge hides keep their data (so unmerging restores it). Updates the
// worksheet's <mergeCells> element and the in-memory merge list.
export function setXlsxMerge(
  sheet: Sheet,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  merge: boolean,
): void {
  const top = Math.min(r1, r2),
    left = Math.min(c1, c2),
    bottom = Math.max(r1, r2),
    right = Math.max(c1, c2);
  const ref = `${colToLetters(left)}${top}:${colToLetters(right)}${bottom}`;
  const merges = (sheet.merges ??= []);
  const idx = merges.findIndex((m) => m.r1 === top && m.c1 === left && m.r2 === bottom && m.c2 === right);
  if (merge) {
    if (idx === -1) merges.push({ r1: top, c1: left, r2: bottom, c2: right });
  } else if (idx !== -1) {
    merges.splice(idx, 1);
  }

  const doc = sheet.doc;
  if (!doc) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  let mcEl = doc.getElementsByTagName("mergeCells")[0] as Element | undefined;
  if (merge) {
    if (!mcEl) {
      mcEl = doc.createElementNS(ns, "mergeCells");
      // <mergeCells> follows <sheetData> in the schema.
      sheet.sheetData?.parentNode?.insertBefore(mcEl, sheet.sheetData.nextSibling);
    }
    const exists = Array.from(mcEl.children).some((m) => m.getAttribute("ref") === ref);
    if (!exists) {
      const mc = doc.createElementNS(ns, "mergeCell");
      mc.setAttribute("ref", ref);
      mcEl.appendChild(mc);
    }
  } else if (mcEl) {
    for (const m of Array.from(mcEl.children))
      if (m.getAttribute("ref") === ref) mcEl.removeChild(m);
  }
  if (mcEl) {
    if (mcEl.children.length === 0) mcEl.parentNode?.removeChild(mcEl);
    else mcEl.setAttribute("count", String(mcEl.children.length));
  }
  sheet.layoutDirty = true;
}

export function writeXlsx(wb: Workbook): void {
  writeXlsxCharts(wb); // persist created/edited charts (DrawingML parts) before serializing sheets
  for (const sheet of wb.sheets) {
    if (!sheet.doc || !sheet.sheetData) continue;
    // Typed dates/percents adopted a number format in the model; persist it to
    // styles.xml for edited cells (a read-side default on untouched cells stays
    // model-only so their XML is not rewritten).
    for (const cell of sheet.cells.values())
      if (cell.numFmtDirty && cell.edited) setXlsxCellNumFmt(wb, sheet, cell, cell.numFmt);
    // A formula change inside a shared group would leave the other members'
    // @si dangling, so rewrite the whole group as plain formulas (de-share).
    const dirtySi = new Set<string>();
    for (const cell of sheet.cells.values()) if (cell.sharedSi != null && cell.fDirty) dirtySi.add(cell.sharedSi);
    let touched = false;
    for (const cell of sheet.cells.values()) {
      const deshare = cell.sharedSi != null && dirtySi.has(cell.sharedSi);
      if (cell.edited || cell.recomputed || deshare) {
        writeXlsxCell(sheet, cell, deshare);
        if (deshare) cell.sharedSi = undefined;
        touched = true;
      }
    }
    if ((touched || sheet.layoutDirty) && sheet.path) wb.files[sheet.path] = serializeXml(sheet.doc);
  }
  if (wb.stylesDirty && wb.stylesDoc) wb.files["xl/styles.xml"] = serializeXml(wb.stylesDoc);
}
