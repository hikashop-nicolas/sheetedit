import type { Cell, DataValidation, DvSpec, Sheet, Workbook } from "../../core/model";
import { colToLetters, ensureCell, firstByLocal, parseXmlOpt, removeByLocal, serializeXml } from "../../core/model";
import { SS_MAIN, ensureXlsxCellEl } from "./shared";
import { writeXlsxCharts } from "./chart-write";
import { writeXlsxImages } from "./image-write";
import { writeXlsxShapes } from "./shape-write";
import { setXlsxCellNumFmt } from "./styles";
// ---------------------------------------------------------------------------
// xlsx write: surgical cell/layout writers and the save pass
// ---------------------------------------------------------------------------

/** Build an <rPr> for a rich-text run, or null when the run has no formatting. Children follow the
    CT_RPrElt schema order (rFont, b, i, strike, color, sz, u). */
function xlsxRunPr(doc: Document, ns: string, run: import("../../core/model").TextRun): Element | null {
  if (!run.bold && !run.italic && !run.underline && !run.strike && !run.size && !run.color && !run.font) return null;
  const rPr = doc.createElementNS(ns, "rPr");
  const flag = (name: string) => { const e = doc.createElementNS(ns, name); rPr.appendChild(e); };
  if (run.font) { const e = doc.createElementNS(ns, "rFont"); e.setAttribute("val", run.font); rPr.appendChild(e); }
  if (run.bold) flag("b");
  if (run.italic) flag("i");
  if (run.strike) flag("strike");
  if (run.color) { const e = doc.createElementNS(ns, "color"); e.setAttribute("rgb", "FF" + run.color.replace(/^#/, "").toUpperCase()); rPr.appendChild(e); }
  if (run.size) { const e = doc.createElementNS(ns, "sz"); e.setAttribute("val", String(run.size)); rPr.appendChild(e); }
  if (run.underline) { const e = doc.createElementNS(ns, "u"); rPr.appendChild(e); }
  return rPr;
}

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
    if (cell.richRuns?.length) {
      // Rich text: one <r> per run, each with its own <rPr> formatting.
      for (const run of cell.richRuns) {
        const r = doc.createElementNS(ns, "r");
        const rPr = xlsxRunPr(doc, ns, run);
        if (rPr) r.appendChild(rPr);
        const rt = doc.createElementNS(ns, "t");
        rt.setAttribute("xml:space", "preserve");
        rt.textContent = run.text;
        r.appendChild(rt);
        is.appendChild(r);
      }
      c.appendChild(is);
      return;
    }
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
// Canonical CT_Worksheet child order, so inserted elements land in a schema-valid position.
const WS_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "drawingHF", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
/** Insert a worksheet child in canonical order (before the first later-ordered sibling). */
function insertWsChild(ws: Element, el: Element): void {
  const idx = WS_ORDER.indexOf(el.localName);
  let before: ChildNode | null = null;
  for (const ch of Array.from(ws.children)) { const ci = WS_ORDER.indexOf(ch.localName); if (ci !== -1 && ci > idx) { before = ch; break; } }
  ws.insertBefore(el, before);
}

/** Append a differential format (a fill) to styles.xml <dxfs>, returning its index. */
function addDxf(wb: Workbook, fillHex: string): number {
  const doc = wb.stylesDoc;
  if (!doc) return -1;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  let dxfs = doc.getElementsByTagName("dxfs")[0];
  if (!dxfs) {
    dxfs = doc.createElementNS(ns, "dxfs");
    const order = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];
    let before: ChildNode | null = null;
    for (const ch of Array.from(doc.documentElement.children)) { const ci = order.indexOf(ch.localName); if (ci !== -1 && ci > order.indexOf("dxfs")) { before = ch; break; } }
    doc.documentElement.insertBefore(dxfs, before);
  }
  const dxf = doc.createElementNS(ns, "dxf");
  const fill = doc.createElementNS(ns, "fill");
  const pf = doc.createElementNS(ns, "patternFill"); pf.setAttribute("patternType", "solid");
  const bg = doc.createElementNS(ns, "bgColor"); bg.setAttribute("rgb", `FF${fillHex.replace("#", "")}`);
  pf.appendChild(bg); fill.appendChild(pf); dxf.appendChild(fill); dxfs.appendChild(dxf);
  dxfs.setAttribute("count", String(dxfs.getElementsByTagName("dxf").length));
  wb.stylesDirty = true;
  return dxfs.getElementsByTagName("dxf").length - 1;
}

/** The equivalent formula Excel stores alongside a timePeriod cfRule (a is the range's top-left cell). */
function timePeriodFormula(period: string, a: string): string {
  switch (period) {
    case "today": return `FLOOR(${a},1)=TODAY()`;
    case "yesterday": return `FLOOR(${a},1)=TODAY()-1`;
    case "tomorrow": return `FLOOR(${a},1)=TODAY()+1`;
    case "last7Days": return `AND(TODAY()-FLOOR(${a},1)<=6,FLOOR(${a},1)<=TODAY())`;
    case "thisWeek": return `AND(TODAY()-ROUNDDOWN(${a},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${a},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case "lastWeek": return `AND(TODAY()-ROUNDDOWN(${a},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${a},0)<(WEEKDAY(TODAY())+7))`;
    case "nextWeek": return `AND(ROUNDDOWN(${a},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${a},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case "thisMonth": return `AND(MONTH(${a})=MONTH(TODAY()),YEAR(${a})=YEAR(TODAY()))`;
    case "lastMonth": return `OR(AND(MONTH(${a})=MONTH(TODAY())-1,YEAR(${a})=YEAR(TODAY())),AND(MONTH(${a})=12,MONTH(TODAY())=1,YEAR(${a})=YEAR(TODAY())-1))`;
    case "nextMonth": return `OR(AND(MONTH(${a})=MONTH(TODAY())+1,YEAR(${a})=YEAR(TODAY())),AND(MONTH(${a})=1,MONTH(TODAY())=12,YEAR(${a})=YEAR(TODAY())+1))`;
    default: return `FLOOR(${a},1)=TODAY()`;
  }
}

export type CfSpec =
  | { kind: "cellIs"; operator: string; value: string; value2?: string; fill: string }
  | { kind: "text"; operator: "containsText" | "notContainsText" | "beginsWith" | "endsWith"; text: string; fill: string }
  | { kind: "top"; rank: number; percent?: boolean; bottom?: boolean; fill: string }
  | { kind: "average"; below?: boolean; equal?: boolean; fill: string }
  | { kind: "dupUnique"; unique?: boolean; fill: string }
  | { kind: "expression"; formula: string; fill: string }
  | { kind: "timePeriod"; period: string; fill: string }
  | { kind: "colorScale"; colors: string[] }
  | { kind: "dataBar"; color: string }
  | { kind: "iconSet"; set: string; count: number };

/** Add (or, with spec null, clear) a conditional-formatting rule over the given ranges, writing the
    worksheet <conditionalFormatting> (+ a dxf for a highlight fill) and updating sheet.condFormats
    so it renders. */
export function setXlsxCondFormat(wb: Workbook, sheet: Sheet, ranges: { r1: number; c1: number; r2: number; c2: number }[], spec: CfSpec | null): void {
  const doc = sheet.doc, ws = doc?.documentElement;
  const ns = ws?.namespaceURI || SS_MAIN;
  const sqref = ranges.map((r) => `${colToLetters(r.c1)}${r.r1}:${colToLetters(r.c2)}${r.r2}`).join(" ");
  if (doc && ws) for (const cf of Array.from(ws.getElementsByTagName("conditionalFormatting"))) if (cf.getAttribute("sqref") === sqref) cf.parentNode?.removeChild(cf);
  sheet.condFormats = (sheet.condFormats ?? []).filter((cf) => cf.ranges.map((r) => `${colToLetters(r.c1)}${r.r1}:${colToLetters(r.c2)}${r.r2}`).join(" ") !== sqref);
  if (!spec) { sheet.layoutDirty = true; return; }
  const priority = 1 + Math.max(0, ...(sheet.condFormats.flatMap((cf) => cf.rules.map((r) => r.priority))));
  // The cfRule @type differs from our spec.kind for text/top/average/dupUnique.
  const cfType = spec.kind === "text" ? spec.operator : spec.kind === "top" ? "top10" : spec.kind === "average" ? "aboveAverage" : spec.kind === "dupUnique" ? (spec.unique ? "uniqueValues" : "duplicateValues") : spec.kind;
  const rule: import("../../core/model").CfRule = { type: cfType, priority };
  const topLeft = `${colToLetters(ranges[0]!.c1)}${ranges[0]!.r1}`;
  if (doc && ws) {
    const cf = doc.createElementNS(ns, "conditionalFormatting"); cf.setAttribute("sqref", sqref);
    const cr = doc.createElementNS(ns, "cfRule"); cr.setAttribute("type", cfType); cr.setAttribute("priority", String(priority));
    const dxfFill = (hex: string) => { const dxfId = addDxf(wb, hex); if (dxfId >= 0) cr.setAttribute("dxfId", String(dxfId)); rule.dxf = { bg: hex }; };
    const formula = (text: string) => { const f = doc.createElementNS(ns, "formula"); f.textContent = text; cr.appendChild(f); };
    if (spec.kind === "cellIs") {
      cr.setAttribute("operator", spec.operator); dxfFill(spec.fill);
      formula(spec.value); rule.operator = spec.operator; rule.formulas = [spec.value];
      if ((spec.operator === "between" || spec.operator === "notBetween") && spec.value2 != null) { formula(spec.value2); rule.formulas.push(spec.value2); }
    } else if (spec.kind === "text") {
      cr.setAttribute("operator", spec.operator); cr.setAttribute("text", spec.text); dxfFill(spec.fill);
      // Excel also stores the equivalent formula for the text predicate.
      const q = spec.text.replace(/"/g, '""');
      const fx = spec.operator === "beginsWith" ? `LEFT(${topLeft},${spec.text.length})="${q}"` : spec.operator === "endsWith" ? `RIGHT(${topLeft},${spec.text.length})="${q}"` : `${spec.operator === "notContainsText" ? "" : "NOT("}ISERROR(SEARCH("${q}",${topLeft}))${spec.operator === "notContainsText" ? "" : ")"}`;
      formula(fx); rule.operator = spec.operator; rule.text = spec.text;
    } else if (spec.kind === "top") {
      cr.setAttribute("rank", String(spec.rank)); if (spec.percent) cr.setAttribute("percent", "1"); if (spec.bottom) cr.setAttribute("bottom", "1"); dxfFill(spec.fill);
      rule.rank = spec.rank; rule.percent = spec.percent; rule.bottom = spec.bottom;
    } else if (spec.kind === "average") {
      if (spec.below) cr.setAttribute("aboveAverage", "0"); if (spec.equal) cr.setAttribute("equalAverage", "1"); dxfFill(spec.fill);
      rule.aboveAverage = !spec.below; rule.equalAverage = spec.equal;
    } else if (spec.kind === "dupUnique") {
      dxfFill(spec.fill);
    } else if (spec.kind === "expression") {
      dxfFill(spec.fill); formula(spec.formula); rule.formulas = [spec.formula];
    } else if (spec.kind === "timePeriod") {
      cr.setAttribute("timePeriod", spec.period); dxfFill(spec.fill);
      formula(timePeriodFormula(spec.period, topLeft)); rule.timePeriod = spec.period;
    } else if (spec.kind === "colorScale") {
      const cs = doc.createElementNS(ns, "colorScale");
      const cfvoTypes = spec.colors.length >= 3 ? ["min", "percentile", "max"] : ["min", "max"];
      cfvoTypes.forEach((ty) => { const v = doc.createElementNS(ns, "cfvo"); v.setAttribute("type", ty); if (ty === "percentile") v.setAttribute("val", "50"); cs.appendChild(v); });
      spec.colors.forEach((col) => { const co = doc.createElementNS(ns, "color"); co.setAttribute("rgb", `FF${col.replace("#", "")}`); cs.appendChild(co); });
      cr.appendChild(cs);
      rule.colorScale = { cfvo: cfvoTypes.map((ty) => ({ type: ty, val: ty === "percentile" ? 50 : undefined })), colors: spec.colors };
    } else if (spec.kind === "dataBar") {
      const db = doc.createElementNS(ns, "dataBar");
      const lo = doc.createElementNS(ns, "cfvo"); lo.setAttribute("type", "min");
      const hi = doc.createElementNS(ns, "cfvo"); hi.setAttribute("type", "max");
      const co = doc.createElementNS(ns, "color"); co.setAttribute("rgb", `FF${spec.color.replace("#", "")}`);
      db.appendChild(lo); db.appendChild(hi); db.appendChild(co); cr.appendChild(db);
      rule.dataBar = { color: spec.color, min: { type: "min" }, max: { type: "max" } };
    } else {
      const is = doc.createElementNS(ns, "iconSet"); is.setAttribute("iconSet", spec.set);
      const cfvo: { type: string; val?: number; gte?: boolean }[] = [];
      for (let i = 0; i < spec.count; i++) { const v = doc.createElementNS(ns, "cfvo"); const pct = Math.round((i / spec.count) * 100); v.setAttribute("type", "percent"); v.setAttribute("val", String(pct)); is.appendChild(v); cfvo.push({ type: "percent", val: pct, gte: true }); }
      cr.appendChild(is);
      rule.iconSet = { set: spec.set, cfvo };
    }
    cf.appendChild(cr);
    insertWsChild(ws, cf);
  }
  sheet.condFormats.push({ ranges, rules: [rule] });
  sheet.layoutDirty = true;
}

// x14 sparklines live in the worksheet extLst under a fixed-uri <ext>. These namespaces + uri are
// the ones Excel emits; the reader matches by local name, so exact prefixes are not load-bearing.
const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
const SPARK_EXT_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";

type SparkStyle = { type: "line" | "column" | "stacked"; color: string; negColor?: string };
type SparkSpec = SparkStyle & { dataRef: string };
type SparkItem = { host: { r: number; c: number }; dataRef: string };

// Locate the x14 sparkline containers, creating them (create=true) or returning what exists.
function sparkContainers(doc: Document, ws: Element, create: boolean): { extLst?: Element; ext?: Element; groups?: Element } {
  const ns = ws.namespaceURI || SS_MAIN;
  const x14 = (name: string): Element => doc.createElementNS(X14_NS, `x14:${name}`);
  let extLst = firstByLocal(ws, "extLst");
  let ext = extLst && Array.from(extLst.children).find((e) => e.localName === "ext" && e.getAttribute("uri") === SPARK_EXT_URI);
  let groups = ext && Array.from(ext.children).find((e) => e.localName === "sparklineGroups");
  if (create) {
    if (!extLst) { extLst = doc.createElementNS(ns, "extLst"); insertWsChild(ws, extLst); }
    if (!ext) { ext = doc.createElementNS(ns, "ext"); ext.setAttribute("uri", SPARK_EXT_URI); extLst.appendChild(ext); }
    if (!groups) { groups = x14("sparklineGroups"); ext.appendChild(groups); }
  }
  return { extLst, ext, groups };
}

// Remove any sparkline(s) whose location is in hostRefs, dropping now-empty groups.
function dropSparklinesAt(groups: Element | undefined, hostRefs: Set<string>): void {
  if (!groups) return;
  for (const g of Array.from(groups.children)) {
    for (const sp of Array.from(g.getElementsByTagName("*")).filter((e) => e.localName === "sparkline")) {
      const sq = Array.from(sp.children).find((c) => c.localName === "sqref");
      if (hostRefs.has((sq?.textContent ?? "").trim().replace(/\$/g, ""))) sp.parentNode?.removeChild(sp);
    }
    if (!Array.from(g.getElementsByTagName("*")).some((e) => e.localName === "sparkline")) g.parentNode?.removeChild(g);
  }
}

// Build one <x14:sparklineGroup> containing a sparkline per item, appended to groups.
function appendSparkGroup(doc: Document, groups: Element, style: SparkStyle, items: SparkItem[]): void {
  const x14 = (name: string): Element => doc.createElementNS(X14_NS, `x14:${name}`);
  const xm = (name: string, text: string): Element => { const e = doc.createElementNS(XM_NS, `xm:${name}`); e.textContent = text; return e; };
  const group = x14("sparklineGroup");
  if (style.type !== "line") group.setAttribute("type", style.type === "stacked" ? "stacked" : "column");
  group.setAttribute("displayEmptyCellsAs", "gap");
  const cs = x14("colorSeries"); cs.setAttribute("rgb", `FF${style.color.replace("#", "")}`); group.appendChild(cs);
  // Column and win/loss sparklines carry a distinct negative-point colour (default Excel red).
  if (style.type !== "line") { const cn = x14("colorNegative"); cn.setAttribute("rgb", `FF${(style.negColor ?? "#d00000").replace("#", "")}`); group.appendChild(cn); }
  const spks = x14("sparklines");
  for (const it of items) {
    const spk = x14("sparkline");
    spk.appendChild(xm("f", it.dataRef));
    spk.appendChild(xm("sqref", `${colToLetters(it.host.c)}${it.host.r}`));
    spks.appendChild(spk);
  }
  group.appendChild(spks);
  groups.appendChild(group);
}

function cleanupSparkContainers(c: { extLst?: Element; ext?: Element; groups?: Element }): void {
  if (c.groups && !c.groups.children.length) c.groups.parentNode?.removeChild(c.groups);
  if (c.ext && !c.ext.children.length) c.ext.parentNode?.removeChild(c.ext);
  if (c.extLst && !c.extLst.children.length) c.extLst.parentNode?.removeChild(c.extLst);
}

/** Add, replace, or (spec === null) remove the sparkline whose location is the host cell. */
export function setXlsxSparkline(sheet: Sheet, host: { r: number; c: number }, spec: SparkSpec | null): void {
  const doc = sheet.doc, ws = doc?.documentElement;
  const hostRef = `${colToLetters(host.c)}${host.r}`;
  sheet.sparklines = (sheet.sparklines ?? []).filter((s) => !(s.host.r === host.r && s.host.c === host.c));
  if (spec) sheet.sparklines.push({ type: spec.type, color: spec.color, negColor: spec.negColor, host: { r: host.r, c: host.c }, dataRef: spec.dataRef });
  if (!sheet.sparklines.length) sheet.sparklines = undefined;
  if (!doc || !ws) return;
  const c = sparkContainers(doc, ws, !!spec);
  dropSparklinesAt(c.groups, new Set([hostRef]));
  if (spec) appendSparkGroup(doc, c.groups!, spec, [{ host, dataRef: spec.dataRef }]);
  else cleanupSparkContainers(c);
  sheet.layoutDirty = true;
}

/** Author one sparkline group spanning several location cells (each mapped to its own data ref). */
export function setXlsxSparklineGroup(sheet: Sheet, style: SparkStyle, items: SparkItem[]): void {
  if (!items.length) return;
  const doc = sheet.doc, ws = doc?.documentElement;
  const hostRefs = new Set(items.map((it) => `${colToLetters(it.host.c)}${it.host.r}`));
  sheet.sparklines = (sheet.sparklines ?? []).filter((s) => !hostRefs.has(`${colToLetters(s.host.c)}${s.host.r}`));
  for (const it of items) sheet.sparklines.push({ type: style.type, color: style.color, negColor: style.negColor, host: { ...it.host }, dataRef: it.dataRef });
  if (!doc || !ws) return;
  const c = sparkContainers(doc, ws, true);
  dropSparklinesAt(c.groups, hostRefs);
  appendSparkGroup(doc, c.groups!, style, items);
  sheet.layoutDirty = true;
}

/** Add or remove a list data validation over the given ranges (1-based inclusive). */
export function setXlsxDataValidation(sheet: Sheet, ranges: { r1: number; c1: number; r2: number; c2: number }[], spec: DvSpec | null): void {
  const doc = sheet.doc, ws = doc?.documentElement;
  const ns = ws?.namespaceURI || SS_MAIN;
  const sqref = ranges.map((r) => `${colToLetters(r.c1)}${r.r1}:${colToLetters(r.c2)}${r.r2}`).join(" ");
  const type = spec ? (spec.type ?? "list") : "list";
  if (doc && ws) {
    let dvs = ws.getElementsByTagName("dataValidations")[0];
    if (dvs) for (const dv of Array.from(dvs.getElementsByTagName("dataValidation"))) if (dv.getAttribute("sqref") === sqref) dv.parentNode?.removeChild(dv);
    if (spec) {
      if (!dvs) { dvs = doc.createElementNS(ns, "dataValidations"); insertWsChild(ws, dvs); }
      const dv = doc.createElementNS(ns, "dataValidation");
      dv.setAttribute("type", type);
      if (type !== "list" && type !== "custom" && spec.operator) dv.setAttribute("operator", spec.operator);
      dv.setAttribute("allowBlank", spec.allowBlank ? "1" : "0");
      dv.setAttribute("showInputMessage", "1");
      dv.setAttribute("showErrorMessage", "1");
      dv.setAttribute("sqref", sqref);
      const addF = (local: string, text: string): void => { const f = doc.createElementNS(ns, local); f.textContent = text; dv.appendChild(f); };
      if (type === "list") addF("formula1", spec.rangeRef ? spec.rangeRef : `"${(spec.values ?? []).join(",")}"`);
      else {
        if (spec.formula1 != null) addF("formula1", spec.formula1);
        if ((spec.operator === "between" || spec.operator === "notBetween") && spec.formula2 != null) addF("formula2", spec.formula2);
      }
      dvs.appendChild(dv);
      dvs.setAttribute("count", String(dvs.getElementsByTagName("dataValidation").length));
    } else if (dvs && !dvs.getElementsByTagName("dataValidation").length) dvs.parentNode?.removeChild(dvs);
  }
  // Keep the in-memory validations (drive the dropdown / invalid outline) in sync.
  sheet.validations = (sheet.validations ?? []).filter((v) => v.ranges.map((r) => `${colToLetters(r.c1)}${r.r1}:${colToLetters(r.c2)}${r.r2}`).join(" ") !== sqref);
  if (spec) {
    const v: DataValidation = { ranges, allowBlank: spec.allowBlank, type };
    if (type === "list") { if (spec.rangeRef) v.rangeRef = spec.rangeRef; else v.values = spec.values; }
    else { v.operator = spec.operator; v.formula1 = spec.formula1; v.formula2 = spec.formula2; }
    sheet.validations.push(v);
  }
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

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

/** Add a rel to a .rels file (creating it), returning the id; reuses an existing rel to the same
    target+type. */
function addRel(wb: Workbook, relsPath: string, type: string, target: string): string {
  let doc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (!doc) doc = parseXmlOpt(enc(`<Relationships xmlns="${REL_NS}"></Relationships>`))!;
  for (const r of Array.from(doc.getElementsByTagName("Relationship"))) if (r.getAttribute("Type") === type && r.getAttribute("Target") === target) return r.getAttribute("Id")!;
  const ids = new Set(Array.from(doc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let n = 1; while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const rel = doc.createElementNS(REL_NS, "Relationship");
  rel.setAttribute("Id", id); rel.setAttribute("Type", type); rel.setAttribute("Target", target);
  doc.documentElement.appendChild(rel);
  wb.files[relsPath] = serializeXml(doc);
  return id;
}

const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
function addContentType(wb: Workbook, opts: { override?: [string, string]; default?: [string, string] }): void {
  const doc = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (!doc || doc.documentElement.localName !== "Types") return;
  if (opts.override) {
    const [part, ct] = opts.override;
    if (!Array.from(doc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === part)) {
      const ov = doc.createElementNS(CT_NS, "Override"); ov.setAttribute("PartName", part); ov.setAttribute("ContentType", ct); doc.documentElement.appendChild(ov);
    }
  }
  if (opts.default) {
    const [ext, ct] = opts.default;
    if (!Array.from(doc.getElementsByTagName("Default")).some((d) => d.getAttribute("Extension") === ext)) {
      const de = doc.createElementNS(CT_NS, "Default"); de.setAttribute("Extension", ext); de.setAttribute("ContentType", ct); doc.documentElement.insertBefore(de, doc.documentElement.firstChild);
    }
  }
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

/** Add, edit or remove a legacy comment (note) on a cell. Writes the sheet's comments part (and a
    minimal VML drawing + legacyDrawing so Excel shows the marker); updates cell.comments so the
    hover popover renders it. */
export function setXlsxComment(wb: Workbook, sheet: Sheet, r: number, c: number, text: string | null, author = "Author"): void {
  const cell = ensureCell(sheet, r, c);
  const ref = `${colToLetters(c)}${r}`;
  const sheetRels = (sheet.path ?? "").replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  // Find (or create) this sheet's comments part.
  let commentsPath: string | undefined;
  const relsDoc = wb.files[sheetRels] ? parseXmlOpt(wb.files[sheetRels]) : undefined;
  if (relsDoc) for (const rel of Array.from(relsDoc.getElementsByTagName("Relationship"))) if (/\/comments$/i.test(rel.getAttribute("Type") ?? "")) {
    const parts: string[] = []; for (const seg of `xl/worksheets/${rel.getAttribute("Target")}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    commentsPath = parts.join("/");
  }
  if (!commentsPath) {
    let n = 1; while (wb.files[`xl/comments${n}.xml`]) n++;
    commentsPath = `xl/comments${n}.xml`;
    wb.files[commentsPath] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${SS_MAIN}"><authors></authors><commentList></commentList></comments>`);
    addRel(wb, sheetRels, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", `../comments${n}.xml`);
    addContentType(wb, { override: [`/${commentsPath}`, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"] });
  }
  const cdoc = parseXmlOpt(wb.files[commentsPath])!;
  const cns = cdoc.documentElement.namespaceURI || SS_MAIN;
  // Author id.
  let authors = cdoc.getElementsByTagName("authors")[0]!;
  let authorId = Array.from(authors.getElementsByTagName("author")).findIndex((a) => a.textContent === author);
  if (authorId < 0 && text != null) { const a = cdoc.createElementNS(cns, "author"); a.textContent = author; authors.appendChild(a); authorId = authors.getElementsByTagName("author").length - 1; }
  const list = cdoc.getElementsByTagName("commentList")[0]!;
  for (const cm of Array.from(list.getElementsByTagName("comment"))) if (cm.getAttribute("ref") === ref) cm.parentNode?.removeChild(cm);
  if (text != null) {
    const cm = cdoc.createElementNS(cns, "comment"); cm.setAttribute("ref", ref); cm.setAttribute("authorId", String(Math.max(0, authorId)));
    const t = cdoc.createElementNS(cns, "text"); const rr = cdoc.createElementNS(cns, "r"); const tt = cdoc.createElementNS(cns, "t"); tt.textContent = text;
    rr.appendChild(tt); t.appendChild(rr); cm.appendChild(t); list.appendChild(cm);
    cell.comments = [{ author, text }];
  } else delete cell.comments;
  wb.files[commentsPath] = serializeXml(cdoc);
  // A minimal VML drawing so Excel renders the note marker (append one shape per comment).
  ensureVmlComment(wb, sheet, sheetRels, r, c, text != null);
  sheet.layoutDirty = true;
}

function ensureVmlComment(wb: Workbook, sheet: Sheet, sheetRels: string, r: number, c: number, present: boolean): void {
  if (!present) return; // removal leaves the (harmless) shape; Excel hides notes with no comment
  const doc = sheet.doc, ws = doc?.documentElement;
  let vmlPath: string | undefined;
  const relsDoc = wb.files[sheetRels] ? parseXmlOpt(wb.files[sheetRels]) : undefined;
  if (relsDoc) for (const rel of Array.from(relsDoc.getElementsByTagName("Relationship"))) if (/vmlDrawing/i.test(rel.getAttribute("Type") ?? "")) {
    const parts: string[] = []; for (const seg of `xl/worksheets/${rel.getAttribute("Target")}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    vmlPath = parts.join("/");
  }
  const header = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>`;
  const shape = (row: number, col: number, id: number): string => `<v:shape id="_x0000_s${1024 + id}" type="#_x0000_t202" style='position:absolute;margin-left:60pt;margin-top:1pt;width:108pt;height:60pt;z-index:${id};visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1}, 15, ${row}, 2, ${col + 3}, 15, ${row + 4}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`;
  if (!vmlPath) {
    let n = 1; while (wb.files[`xl/drawings/vmlDrawing${n}.vml`]) n++;
    vmlPath = `xl/drawings/vmlDrawing${n}.vml`;
    wb.files[vmlPath] = enc(`${header}${shape(r - 1, c - 1, 1)}</xml>`);
    const rid = addRel(wb, sheetRels, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing", `../drawings/vmlDrawing${n}.vml`);
    addContentType(wb, { default: ["vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"] });
    if (doc && ws && !ws.getElementsByTagName("legacyDrawing").length) {
      if (!ws.getAttribute("xmlns:r")) ws.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
      const ld = doc.createElementNS(ws.namespaceURI || SS_MAIN, "legacyDrawing"); ld.setAttribute("r:id", rid); insertWsChild(ws, ld);
    }
  } else {
    // Append another shape to the existing VML if this cell has none yet.
    const cur = new TextDecoder().decode(wb.files[vmlPath]!);
    if (!cur.includes(`<x:Row>${r - 1}</x:Row><x:Column>${c - 1}</x:Column>`)) {
      const count = (cur.match(/<v:shape /g) || []).length;
      wb.files[vmlPath] = enc(cur.replace(/<\/xml>\s*$/, `${shape(r - 1, c - 1, count + 1)}</xml>`));
    }
  }
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
  writeXlsxImages(wb); // persist moved/resized pictures into their drawing parts
  writeXlsxShapes(wb); // persist authored/moved/resized/restyled shapes into their drawing parts
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
