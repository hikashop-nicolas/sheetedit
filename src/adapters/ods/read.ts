import type { Cell, CellKind, CellStyle, CondFormat, Phonetic, Sheet, TextRun, Workbook } from "../../core/model";
import { formatNumber, getCell, key, noteExtent, numToStr, parseA1Ref, parseXml, parseXmlOpt } from "../../core/model";
import { pivotColumnItems, type PivotFunc } from "../../core/pivot";
import { durationToSerial, isoToSerial } from "../../core/dates";

const ODF_TO_FUNC: Record<string, PivotFunc> = { sum: "sum", count: "count", countnums: "countNums", average: "average", min: "min", max: "max" };
import { readOdsCharts } from "./chart-read";
import { readOdsImages } from "./image-read";
import { readOdsShapes } from "./shape-read";
import { REPEAT_CAP, odfToA1, odsBorderColor, odsCellComments, odsCellLink, odsCellRich, odsCellText, odsColorOf, odsLenToPx } from "./shared";
// ---------------------------------------------------------------------------
// ods read: content.xml parsing (tables, rows, styles)
// ---------------------------------------------------------------------------

export interface OdsStyles {
  cell: Map<string, CellStyle>; // family table-cell -> resolved style
  colW: Map<string, number>; // family table-column -> width px
  rowH: Map<string, number>; // family table-row -> height px
  // family table-cell -> its <style:map> conditional-format entries (standard ODF CF).
  cfMap: Map<string, { cond: string; apply: string }[]>;
  text: Map<string, CellStyle>; // family text -> resolved run style (for in-cell <text:span>)
}

// Parse <style:style> from the given docs (content.xml automatic styles + styles.xml),
// resolving table-cell parent chains, and the column/row dimension styles.
export function parseOdsStyles(docs: Document[]): OdsStyles {
  const raw = new Map<string, { el: Element; parent?: string }>();
  const rawText = new Map<string, { el: Element; parent?: string }>();
  const colW = new Map<string, number>();
  const rowH = new Map<string, number>();
  const cfMap = new Map<string, { cond: string; apply: string }[]>();
  for (const doc of docs) {
    for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
      const name = st.getAttribute("style:name");
      if (!name) continue;
      const family = st.getAttribute("style:family");
      if (family === "table-cell") {
        raw.set(name, { el: st, parent: st.getAttribute("style:parent-style-name") || undefined });
        const maps = Array.from(st.getElementsByTagName("style:map"));
        if (maps.length) cfMap.set(name, maps.map((m) => ({ cond: m.getAttribute("style:condition") ?? "", apply: m.getAttribute("style:apply-style-name") ?? "" })).filter((x) => x.cond && x.apply));
      } else if (family === "text") {
        rawText.set(name, { el: st, parent: st.getAttribute("style:parent-style-name") || undefined });
      } else if (family === "table-column") {
        const p = st.getElementsByTagName("style:table-column-properties")[0];
        const w = odsLenToPx(p?.getAttribute("style:column-width"));
        if (w) colW.set(name, w);
      } else if (family === "table-row") {
        const p = st.getElementsByTagName("style:table-row-properties")[0];
        const h = odsLenToPx(p?.getAttribute("style:row-height"));
        if (h) rowH.set(name, h);
      }
    }
  }
  const ownStyle = (el: Element): CellStyle => {
    const s: CellStyle = {};
    const cp = el.getElementsByTagName("style:table-cell-properties")[0];
    if (cp) {
      const bg = odsColorOf(cp.getAttribute("fo:background-color"));
      if (bg) s.bg = bg;
      const all = cp.getAttribute("fo:border");
      const sides: CellStyle["borders"] = {};
      const top = odsBorderColor(cp.getAttribute("fo:border-top") ?? all);
      const right = odsBorderColor(cp.getAttribute("fo:border-right") ?? all);
      const bottom = odsBorderColor(cp.getAttribute("fo:border-bottom") ?? all);
      const left = odsBorderColor(cp.getAttribute("fo:border-left") ?? all);
      if (top) sides.top = top;
      if (right) sides.right = right;
      if (bottom) sides.bottom = bottom;
      if (left) sides.left = left;
      if (top || right || bottom || left) s.borders = sides;
      if (cp.getAttribute("fo:wrap-option") === "wrap") s.wrap = true;
      const va = cp.getAttribute("style:vertical-align");
      if (va === "top" || va === "middle" || va === "bottom") s.valign = va;
    }
    const tp = el.getElementsByTagName("style:text-properties")[0];
    if (tp) {
      if (tp.getAttribute("fo:font-weight") === "bold") s.bold = true;
      if (tp.getAttribute("fo:font-style") === "italic") s.italic = true;
      const us = tp.getAttribute("style:text-underline-style");
      if (us && us !== "none") {
        s.underline = true;
        // ODF underline style/type -> CSS text-decoration-style (double via text-underline-type).
        const flavour =
          tp.getAttribute("style:text-underline-type") === "double"
            ? "double"
            : us === "dotted"
              ? "dotted"
              : /dash/.test(us)
                ? "dashed"
                : us === "wave"
                  ? "wavy"
                  : undefined;
        if (flavour) s.underlineStyle = flavour;
      }
      const ls = tp.getAttribute("style:text-line-through-style");
      if (ls && ls !== "none") s.strike = true;
      const fs = parseFloat(tp.getAttribute("fo:font-size") ?? "");
      if (fs && (tp.getAttribute("fo:font-size") ?? "").endsWith("pt")) s.fontSize = fs;
      const ff = tp.getAttribute("fo:font-family") ?? tp.getAttribute("style:font-name");
      if (ff) s.fontFamily = ff.replace(/^['"]|['"]$/g, "");
      const col = odsColorOf(tp.getAttribute("fo:color"));
      if (col) s.color = col;
    }
    const pp = el.getElementsByTagName("style:paragraph-properties")[0];
    const ta = pp?.getAttribute("fo:text-align");
    if (ta === "center") s.align = "center";
    else if (ta === "end" || ta === "right") s.align = "right";
    else if (ta === "start" || ta === "left") s.align = "left";
    return s;
  };
  const cell = new Map<string, CellStyle>();
  const resolve = (name: string, depth = 0): CellStyle => {
    const cached = cell.get(name);
    if (cached) return cached;
    const entry = raw.get(name);
    if (!entry || depth > 8) return {};
    const base = entry.parent ? resolve(entry.parent, depth + 1) : {};
    const merged = { ...base, ...ownStyle(entry.el) };
    cell.set(name, merged);
    return merged;
  };
  for (const name of raw.keys()) resolve(name);
  const text = new Map<string, CellStyle>();
  const resolveText = (name: string, depth = 0): CellStyle => {
    const cached = text.get(name);
    if (cached) return cached;
    const entry = rawText.get(name);
    if (!entry || depth > 8) return {};
    const base = entry.parent ? resolveText(entry.parent, depth + 1) : {};
    const merged = { ...base, ...ownStyle(entry.el) };
    text.set(name, merged);
    return merged;
  };
  for (const name of rawText.keys()) resolveText(name);
  return { cell, colW, rowH, cfMap, text };
}

// Frozen panes live in settings.xml (ODF view settings), keyed by sheet name. SplitMode 2 =
// "frozen" (1 = split, 0 = none); the SplitPosition is the count of frozen columns / rows.
function readOdsFreeze(files: Record<string, Uint8Array>): Map<string, { rows: number; cols: number }> {
  const out = new Map<string, { rows: number; cols: number }>();
  const f = files["settings.xml"];
  if (!f) return out;
  const doc = parseXmlOpt(f);
  if (!doc) return out;
  const tablesMaps = Array.from(doc.getElementsByTagName("config:config-item-map-named")).filter(
    (m) => m.getAttribute("config:name") === "Tables",
  );
  for (const tables of tablesMaps) {
    for (const entry of Array.from(tables.children)) {
      if (entry.localName !== "config-item-map-entry") continue;
      const name = entry.getAttribute("config:name");
      if (!name) continue;
      const item = (key: string): number => {
        for (const ci of Array.from(entry.children))
          if (ci.localName === "config-item" && ci.getAttribute("config:name") === key) return Number(ci.textContent || "0");
        return 0;
      };
      const cols = item("HorizontalSplitMode") === 2 ? Math.max(0, Math.floor(item("HorizontalSplitPosition"))) : 0;
      const rows = item("VerticalSplitMode") === 2 ? Math.max(0, Math.floor(item("VerticalSplitPosition"))) : 0;
      if (rows > 0 || cols > 0) out.set(name, { rows, cols });
    }
  }
  return out;
}

// One ODF cell address ("$Sheet1.$A$1" or ".$A$1") -> A1 ("Sheet1!A1" or "A1").
function odfRefToA1(part: string): string {
  const dot = part.indexOf(".");
  const sheet = dot > 0 ? part.slice(0, dot).replace(/^\$/, "").replace(/^'|'$/g, "") : "";
  const cell = part.slice(dot + 1).replace(/\$/g, "");
  return sheet ? `${sheet}!${cell}` : cell;
}
// An ODF cell-range-address -> an A1 reference ("Sheet1!A1:B2"); the end cell drops its sheet.
function odfAddrToA1(addr: string): string {
  const parts = addr.split(":");
  const first = odfRefToA1(parts[0]!);
  if (parts.length === 1) return first;
  const second = odfRefToA1(parts[1]!);
  return `${first}:${second.includes("!") ? second.split("!")[1] : second}`;
}

export function readOds(files: Record<string, Uint8Array>): Workbook {
  const contentFile = files["content.xml"];
  if (!contentFile) throw new Error("not an .ods: content.xml missing");
  const contentDoc = parseXml(contentFile);
  const docs = [contentDoc];
  const odsStylesDoc = files["styles.xml"] ? parseXmlOpt(files["styles.xml"]) : undefined;
  if (odsStylesDoc) docs.push(odsStylesDoc);
  const styles = parseOdsStyles(docs);
  const freezeByName = readOdsFreeze(files);
  const wb: Workbook = { kind: "ods", sheets: [], files, contentDoc, contentPath: "content.xml" };
  // Defined names: <table:named-range table:name="X" table:cell-range-address="$Sheet1.$A$1:.$B$2"/>.
  // Convert the ODF address to an A1 reference ("Sheet1!A1:B2") for recalc.
  const definedNames = new Map<string, string>();
  for (const nr of Array.from(contentDoc.getElementsByTagName("table:named-range"))) {
    const name = nr.getAttribute("table:name");
    const addr = nr.getAttribute("table:cell-range-address");
    if (name && addr) definedNames.set(name, odfAddrToA1(addr));
  }
  if (definedNames.size) wb.definedNames = definedNames;
  const validationDefs = parseOdsValidations(contentDoc);
  const condFormats = parseOdsCondFormats(contentDoc);
  for (const table of Array.from(contentDoc.getElementsByTagName("table:table"))) {
    const name = table.getAttribute("table:name") ?? `Sheet${wb.sheets.length + 1}`;
    const sheet: Sheet = { name, cells: new Map(), maxRow: 0, maxCol: 0, tableEl: table };
    const fz = freezeByName.get(name);
    if (fz) sheet.freeze = fz;
    readOdsTable(sheet, table, styles);
    buildOdsValidations(sheet, validationDefs);
    // Conditional formats whose target sheet is this one (empty sheet name = single-sheet target).
    const cfs = condFormats.get(name) ?? (wb.sheets.length === 0 ? condFormats.get("") : undefined);
    if (cfs?.length) sheet.condFormats = cfs;
    buildOdsStyleMapCf(sheet, styles); // standard <style:map> CF (the interoperable mechanism)
    wb.sheets.push(sheet);
  }
  readOdsCharts(wb, files);
  readOdsImages(wb, files);
  readOdsShapes(wb);
  readOdsSparklines(wb);
  readOdsAutoFilter(wb);
  readOdsPivots(wb);
  return wb;
}

// Parse the document-level <table:content-validation> definitions (name -> spec).
type DvType = NonNullable<import("../../core/model").DataValidation["type"]>;
type DvOp = NonNullable<import("../../core/model").DataValidation["operator"]>;
type OdsValidationDef = { values?: string[]; rangeRef?: string; allowBlank?: boolean; type?: DvType; operator?: DvOp; formula1?: string; formula2?: string };

/** Parse a typed (non-list) ODF content-validation condition into type / operator / operands. */
function parseTypedCond(cond: string): { type: DvType; operator?: DvOp; formula1?: string; formula2?: string } | null {
  const custom = /is-true-formula\((.*)\)\s*$/.exec(cond);
  if (custom) return { type: "custom", formula1: custom[1]!.trim() };
  const isTextLen = /cell-content-text-length/.test(cond);
  const fn = isTextLen ? "cell-content-text-length" : "cell-content";
  let operator: DvOp | undefined, f1: string | undefined, f2: string | undefined;
  const between = new RegExp(`${fn}-is-(not-)?between\\(([^,;]+)[,;]([^)]+)\\)`).exec(cond);
  if (between) { operator = between[1] ? "notBetween" : "between"; f1 = between[2]!.trim(); f2 = between[3]!.trim(); }
  else {
    const cmp = new RegExp(`${fn}\\(\\)\\s*(>=|<=|<>|!=|>|<|=)\\s*([^\\s)]+)`).exec(cond);
    const OPS: Record<string, DvOp> = { ">=": "greaterThanOrEqual", "<=": "lessThanOrEqual", "<>": "notEqual", "!=": "notEqual", ">": "greaterThan", "<": "lessThan", "=": "equal" };
    if (cmp) { operator = OPS[cmp[1]!]; f1 = cmp[2]!.trim(); }
  }
  const TYPE: Record<string, DvType> = { "whole-number": "whole", "decimal-number": "decimal", date: "date", time: "time" };
  const tm = /cell-content-is-(whole-number|decimal-number|date|time)\(\)/.exec(cond);
  const type = isTextLen ? "textLength" : (tm ? TYPE[tm[1]!] : undefined);
  return type && operator ? { type, operator, formula1: f1, formula2: f2 } : null;
}

function parseOdsValidations(doc: Document): Map<string, OdsValidationDef> {
  const map = new Map<string, OdsValidationDef>();
  for (const cv of Array.from(doc.getElementsByTagName("table:content-validation"))) {
    const name = cv.getAttribute("table:name");
    if (!name) continue;
    const cond = cv.getAttribute("table:condition") ?? "";
    const allowBlank = cv.getAttribute("table:allow-empty-cell") !== "false";
    const list = /cell-content-is-in-list\((.*)\)\s*$/.exec(cond);
    if (list) {
      const arg = list[1]!.trim();
      if (arg.startsWith("[")) map.set(name, { rangeRef: odfAddrToA1(arg.replace(/^\[/, "").replace(/\]$/, "")), allowBlank, type: "list" });
      else { const values = [...arg.matchAll(/"((?:[^"]|"")*)"/g)].map((q) => q[1]!.replace(/""/g, '"')); map.set(name, { values, allowBlank, type: "list" }); }
      continue;
    }
    const typed = parseTypedCond(cond);
    map.set(name, typed ? { ...typed, allowBlank } : { allowBlank });
  }
  return map;
}

// Group a sheet's cells by their content-validation-name into sheet.validations (one per name).
function buildOdsValidations(sheet: Sheet, defs: Map<string, OdsValidationDef>): void {
  if (!defs.size) return;
  const byName = new Map<string, { r: number; c: number }[]>();
  for (const cell of sheet.cells.values()) {
    if (!cell.odsValidationName) continue;
    (byName.get(cell.odsValidationName) ?? byName.set(cell.odsValidationName, []).get(cell.odsValidationName)!).push({ r: cell.row, c: cell.col });
  }
  const out: NonNullable<Sheet["validations"]> = [];
  for (const [name, cells] of byName) {
    const def = defs.get(name);
    // list validations drive a dropdown; typed ones drive the invalid outline. Conditions we could
    // not parse are preserved via the untouched block + the cell's name but surface no UI.
    if (!def || (!def.values && !def.rangeRef && !def.type)) continue;
    const ranges = cells.map((p) => ({ r1: p.r, c1: p.c, r2: p.r, c2: p.c }));
    if (def.type && def.type !== "list") out.push({ ranges, allowBlank: def.allowBlank, type: def.type, operator: def.operator, formula1: def.formula1, formula2: def.formula2 });
    else out.push({ ranges, values: def.values, rangeRef: def.rangeRef, allowBlank: def.allowBlank, type: "list" });
  }
  if (out.length) sheet.validations = out;
}

// A standard ODF <style:map> condition ("cell-content()>5", "cell-content-is-between(1,10)")
// -> operator + operands. is-true-formula and other forms are not rendered (returned null).
function parseStyleMapCondition(cond: string): { operator: string; formulas: string[] } | null {
  const c = cond.trim();
  const btw = /^cell-content-is-between\((.+),(.+)\)$/i.exec(c);
  if (btw) return { operator: "between", formulas: [btw[1]!.trim(), btw[2]!.trim()] };
  const nbtw = /^cell-content-is-not-between\((.+),(.+)\)$/i.exec(c);
  if (nbtw) return { operator: "notBetween", formulas: [nbtw[1]!.trim(), nbtw[2]!.trim()] };
  const m = /^cell-content\(\)\s*(<=|>=|!=|<>|<|>|=)\s*(.+)$/.exec(c);
  if (!m) return null;
  const op = { "<": "lessThan", ">": "greaterThan", "<=": "lessThanOrEqual", ">=": "greaterThanOrEqual", "=": "equal", "!=": "notEqual", "<>": "notEqual" }[m[1]!]!;
  return { operator: op, formulas: [m[2]!.trim()] };
}

// Standard ODF conditional formatting: cells whose base style carries <style:map> entries.
// Grouped by style (each producing a CondFormat over its cells), the map's apply-style resolved
// to a dxf. This is the interoperable mechanism; calcext (below) only adds colour scales etc.
function buildOdsStyleMapCf(sheet: Sheet, styles: OdsStyles): void {
  if (!styles.cfMap.size) return;
  const byStyle = new Map<string, { r: number; c: number }[]>();
  for (const cell of sheet.cells.values())
    if (cell.style && styles.cfMap.has(cell.style)) (byStyle.get(cell.style) ?? byStyle.set(cell.style, []).get(cell.style)!).push({ r: cell.row, c: cell.col });
  const out: NonNullable<Sheet["condFormats"]> = sheet.condFormats ? [...sheet.condFormats] : [];
  for (const [styleName, cells] of byStyle) {
    const rules: CondFormat["rules"] = [];
    let priority = 1;
    for (const m of styles.cfMap.get(styleName)!) {
      const parsed = parseStyleMapCondition(m.cond);
      const cs = styles.cell.get(m.apply);
      if (parsed && cs) rules.push({ type: "cellIs", priority: priority++, operator: parsed.operator, formulas: parsed.formulas, dxf: { bg: cs.bg, color: cs.color, bold: cs.bold, italic: cs.italic } });
    }
    if (rules.length) out.push({ ranges: cells.map((p) => ({ r1: p.r, c1: p.c, r2: p.r, c2: p.c })), rules });
  }
  if (out.length) sheet.condFormats = out;
}

// --- conditional formatting (LibreOffice calcext extension) ----------------
const attrByLocal = (el: Element, local: string): string | null => {
  for (const a of Array.from(el.attributes)) if (a.localName === local) return a.value;
  return null;
};
const kidsByLocal = (el: Element, local: string): Element[] => Array.from(el.children).filter((c) => c.localName === local);

// Parse <calcext:conditional-formats> into per-sheet CondFormat lists, resolving a condition's
// apply-style-name to a dxf via the cell-style map, and reading colour scales / data bars / icons.
function parseOdsCondFormats(doc: Document): Map<string, CondFormat[]> {
  const bySheet = new Map<string, CondFormat[]>();
  // Map calcext threshold types onto the ones the renderer's stopValue understands.
  const cfvoType: Record<string, string> = { minimum: "min", maximum: "max", number: "num", value: "num", "auto-minimum": "min", "auto-maximum": "max" };
  const cfvo = (e: Element): { type: string; val?: number } => {
    const type = attrByLocal(e, "type") ?? "value";
    const raw = attrByLocal(e, "value");
    const n = raw != null ? Number(raw) : undefined;
    return { type: cfvoType[type] ?? type, val: Number.isFinite(n) ? n : undefined };
  };
  const colorOf = (e: Element): string => attrByLocal(e, "color") ?? "#ffffff";
  for (const cf of Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "conditional-format")) {
    const target = attrByLocal(cf, "target-range-address");
    if (!target) continue;
    // "Sheet1.A1:Sheet1.A5 Sheet1.C1" -> per-sheet rects.
    const perSheet = new Map<string, { r1: number; c1: number; r2: number; c2: number }[]>();
    for (const addr of target.split(/\s+/)) {
      const a1 = odfAddrToA1(addr.replace(/^\[/, "").replace(/\]$/, ""));
      const bang = a1.indexOf("!");
      const sheetName = bang >= 0 ? a1.slice(0, bang) : "";
      const body = (bang >= 0 ? a1.slice(bang + 1) : a1).replace(/\$/g, "");
      const [a, b] = body.split(":");
      const p1 = parseA1Ref(a ?? ""); const p2 = b ? parseA1Ref(b) : p1;
      if (!p1 || !p2) continue;
      (perSheet.get(sheetName) ?? perSheet.set(sheetName, []).get(sheetName)!).push({ r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) });
    }
    const rules: CondFormat["rules"] = [];
    let priority = 1;
    for (const child of Array.from(cf.children)) {
      const l = child.localName;
      // calcext:condition (a cellIs) is only LibreOffice's mirror of the standard <style:map>,
      // which buildOdsStyleMapCf already reads; skip it here to avoid duplicate rules. Colour
      // scales / data bars / icon sets have no style:map form, so they are read only from calcext.
      if (l === "color-scale") {
        const entries = kidsByLocal(child, "color-scale-entry");
        rules.push({ type: "colorScale", priority: priority++, colorScale: { cfvo: entries.map(cfvo), colors: entries.map(colorOf) } });
      } else if (l === "data-bar") {
        const entries = kidsByLocal(child, "formatting-entry");
        rules.push({ type: "dataBar", priority: priority++, dataBar: { color: attrByLocal(child, "positive-color") ?? "#638ec6", min: entries[0] ? cfvo(entries[0]) : { type: "min" }, max: entries[entries.length - 1] ? cfvo(entries[entries.length - 1]!) : { type: "max" } } });
      } else if (l === "icon-set") {
        const entries = kidsByLocal(child, "formatting-entry");
        rules.push({ type: "iconSet", priority: priority++, iconSet: { set: attrByLocal(child, "icon-set-type") ?? "3TrafficLights1", cfvo: entries.map((e) => ({ ...cfvo(e), gte: true })) } });
      }
    }
    if (!rules.length) continue;
    for (const [sheetName, ranges] of perSheet)
      (bySheet.get(sheetName) ?? bySheet.set(sheetName, []).get(sheetName)!).push({ ranges, rules });
  }
  return bySheet;
}

// AutoFilter: an ODF <table:database-range table:display-filter-buttons="true"> whose
// target-range-address names the sheet + range -> sheet.autoFilter (shows the filter carets).
function readOdsAutoFilter(wb: Workbook): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  for (const dr of Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "database-range")) {
    if (dr.getAttribute("table:display-filter-buttons") !== "true" && attrByLocal(dr, "display-filter-buttons") !== "true") continue;
    const target = attrByLocal(dr, "target-range-address");
    if (!target) continue;
    const a1 = odfAddrToA1(target.replace(/\$/g, ""));
    const bang = a1.indexOf("!");
    const sheetName = bang >= 0 ? a1.slice(0, bang) : "";
    const body = (bang >= 0 ? a1.slice(bang + 1) : a1).split(":");
    const p1 = parseA1Ref(body[0] ?? ""); const p2 = body[1] ? parseA1Ref(body[1]) : p1;
    const sheet = wb.sheets.find((s) => s.name === sheetName) ?? (sheetName === "" ? wb.sheets[0] : undefined);
    if (!sheet || !p1 || !p2) continue;
    sheet.autoFilter = { r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) };
  }
}

// LibreOffice sparklines: <calcext:sparkline-groups> is a child of <table:table> (after the rows).
// Each group carries the type + colours; each <calcext:sparkline> binds a host cell-address to a
// data-range. Populates sheet.sparklines (rendered by the same mini-chart drawer as xlsx).
function readOdsSparklines(wb: Workbook): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  for (const group of Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "sparkline-group")) {
    let t: Element | null = group.parentElement;
    while (t && t.localName !== "table") t = t.parentElement;
    const sheet = t ? wb.sheets.find((s) => s.tableEl === t) : undefined;
    if (!sheet) continue;
    const type = (attrByLocal(group, "type") as "line" | "column" | "stacked") || "line";
    const colorRaw = attrByLocal(group, "color-series") || "#376092";
    const color = colorRaw.startsWith("#") ? colorRaw : `#${colorRaw}`;
    const negRaw = attrByLocal(group, "color-negative");
    const negColor = negRaw ? (negRaw.startsWith("#") ? negRaw : `#${negRaw}`) : undefined;
    for (const sp of Array.from(group.getElementsByTagName("*")).filter((e) => e.localName === "sparkline")) {
      const addr = attrByLocal(sp, "cell-address");
      const data = attrByLocal(sp, "data-range");
      if (!addr || !data) continue;
      const hostA1 = odfRefToA1(addr.replace(/\$/g, ""));
      const p = parseA1Ref((hostA1.includes("!") ? hostA1.split("!")[1]! : hostA1).split(":")[0]!);
      if (!p) continue;
      (sheet.sparklines ??= []).push({ type, color, negColor, host: { r: p.row, c: p.col }, dataRef: odfAddrToA1(data.replace(/\$/g, "")) });
    }
  }
}

// One ODF range/address -> { sheet, r1..c2 } (1-based inclusive), or null.
function odfRange(addr: string | null): { sheet: string; r1: number; c1: number; r2: number; c2: number } | null {
  if (!addr) return null;
  const a1 = odfAddrToA1(addr.replace(/\$/g, ""));
  const bang = a1.indexOf("!");
  const sheet = bang >= 0 ? a1.slice(0, bang) : "";
  const body = (bang >= 0 ? a1.slice(bang + 1) : a1).split(":");
  const p1 = parseA1Ref(body[0] ?? ""); const p2 = body[1] ? parseA1Ref(body[1]) : p1;
  if (!p1 || !p2) return null;
  return { sheet, r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) };
}

// Data-pilot (pivot) tables: <table:data-pilot-table> carries the output target-range-address and,
// per source field, an orientation (row/column/data/page). Read-only: the output already renders as
// the cells LibreOffice materialised; this models the definition so the UI can surface it.
function readOdsPivots(wb: Workbook): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  for (const pt of Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "data-pilot-table")) {
    const target = odfRange(attrByLocal(pt, "target-range-address"));
    const host = target ? (wb.sheets.find((s) => s.name === target.sheet) ?? (target.sheet === "" ? wb.sheets[0] : undefined)) : wb.sheets[0];
    if (!host) continue;
    const info: import("../../core/model").PivotTableInfo = { name: attrByLocal(pt, "name") ?? "DataPilot", rowFields: [], colFields: [], pageFields: [], dataFields: [] };
    if (target) info.targetRange = { r1: target.r1, c1: target.c1, r2: target.r2, c2: target.c2 };
    const src = Array.from(pt.children).find((c) => c.localName === "source-cell-range");
    const srcR = src ? odfRange(attrByLocal(src, "cell-range-address")) : null;
    if (srcR) { info.sourceSheet = srcR.sheet; info.sourceRange = { r1: srcR.r1, c1: srcR.c1, r2: srcR.r2, c2: srcR.c2 }; }
    info.hostSheet = host.name;
    // Map source-field-name -> column offset within the source range, to reconstruct the spec.
    const srcSheet = srcR ? wb.sheets.find((s) => s.name === srcR.sheet) : undefined;
    const nameOff = new Map<string, number>();
    if (srcR && srcSheet) for (let c = 0; c <= srcR.c2 - srcR.c1; c++) { const nm = (getCell(srcSheet, srcR.r1, srcR.c1 + c)?.value ?? "").trim(); if (nm && !nameOff.has(nm)) nameOff.set(nm, c); }
    const rows: number[] = [], cols: number[] = [], values: { field: number; func: import("../../core/pivot").PivotFunc }[] = [], pages: { field: number; item: number | null }[] = [];
    let subtotals = false, resolvable = !!(srcR && srcSheet);
    const hasSubtotal = (f: Element): boolean => Array.from(f.getElementsByTagName("*")).some((e) => e.localName === "data-pilot-subtotal" || e.localName === "data-pilot-subtotals");
    const pageMember = (f: Element): number | null => {
      const mem = Array.from(f.getElementsByTagName("*")).find((e) => e.localName === "data-pilot-member" && attrByLocal(e, "display") === "true");
      const off = nameOff.get(attrByLocal(f, "source-field-name") ?? "");
      if (!mem || off == null || !srcSheet || !srcR) return null;
      const nm = attrByLocal(mem, "name");
      const items = pivotColumnItems(srcSheet, { r1: srcR.r1, c1: srcR.c1, r2: srcR.r2, c2: srcR.c2 }, off);
      const i = items.findIndex((it) => String(it.value) === nm);
      return i >= 0 ? i : null;
    };
    for (const f of Array.from(pt.children)) {
      if (f.localName !== "data-pilot-field") continue;
      const name = attrByLocal(f, "source-field-name");
      if (!name) continue; // the values-layout placeholder has no source field
      const orient = attrByLocal(f, "orientation");
      const off = nameOff.get(name);
      if (off == null) resolvable = false;
      if (orient === "row") { info.rowFields.push(name); if (off != null) rows.push(off); if (hasSubtotal(f)) subtotals = true; }
      else if (orient === "column") { info.colFields.push(name); if (off != null) cols.push(off); if (hasSubtotal(f)) subtotals = true; }
      else if (orient === "page") { info.pageFields.push(name); if (off != null) pages.push({ field: off, item: pageMember(f) }); }
      else if (orient === "data") {
        const fn = (attrByLocal(f, "function") ?? "").toLowerCase();
        info.dataFields.push({ name, func: fn && fn !== "auto" ? fn : "sum" });
        if (off != null) values.push({ field: off, func: ODF_TO_FUNC[fn] ?? "sum" });
      }
    }
    if (resolvable && srcR && values.length) info.authorSpec = { source: { r1: srcR.r1, c1: srcR.c1, r2: srcR.r2, c2: srcR.c2 }, rows, cols, values, pages: pages.length ? pages : undefined, subtotals: subtotals || undefined };
    (host.pivotTables ??= []).push(info);
  }
}

export function readOdsTable(sheet: Sheet, table: Element, styles: OdsStyles): void {
  // Column widths + hidden state: walk <table:table-column> (each may repeat) and map to px.
  const cols = new Map<number, number>();
  const hiddenCols = new Set<number>();
  const isHidden = (el: Element): boolean => {
    const v = el.getAttribute("table:visibility");
    return v === "collapse" || v === "filter";
  };
  let colIdx = 0;
  const collectCols = (parent: Element) => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-column") {
        const rep = Math.max(1, Number(ch.getAttribute("table:number-columns-repeated") || "1"));
        const w = styles.colW.get(ch.getAttribute("table:style-name") ?? "");
        const hidden = isHidden(ch);
        for (let i = 0; i < Math.min(rep, REPEAT_CAP); i++) {
          colIdx++;
          if (w) cols.set(colIdx, w);
          if (hidden) hiddenCols.add(colIdx);
        }
      } else if (ch.localName === "table-header-columns" || ch.localName === "table-columns") {
        collectCols(ch);
      }
    }
  };
  collectCols(table);
  if (cols.size) sheet.colWidths = cols;
  if (hiddenCols.size) sheet.hiddenCols = hiddenCols;

  let rowNum = 0;
  const rows: { el: Element; header: boolean }[] = [];
  let headerGroupEl: Element | undefined;
  const collect = (parent: Element, header = false) => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-row") rows.push({ el: ch, header });
      else if (ch.localName === "table-header-rows" || ch.localName === "table-rows") {
        if (ch.localName === "table-header-rows") headerGroupEl = ch;
        collect(ch, header || ch.localName === "table-header-rows");
      }
    }
  };
  collect(table);
  const rowHeights = new Map<number, number>();
  const hiddenRows = new Set<number>();
  const rowStyles = new Map<number, string>();
  const rowEls = new Map<number, Element>();
  const rowRuns: { from: number; to: number; el: Element }[] = [];
  const coveredEls = new Map<string, Element>();
  const merges: { r1: number; c1: number; r2: number; c2: number }[] = [];
  let headerFrom = 0;
  let headerTo = 0;
  for (const { el: rowEl, header } of rows) {
    const rrep = Math.max(1, Number(rowEl.getAttribute("table:number-rows-repeated") || "1"));
    const rowStyle = rowEl.getAttribute("table:style-name") ?? undefined;
    const rh = styles.rowH.get(rowStyle ?? "");
    const parsedCells = parseOdsRow(rowEl, styles);
    const rowHasContent = parsedCells.some((c) => c.has);
    const copies = rowHasContent ? Math.min(rrep, REPEAT_CAP) : 0;
    if (header) {
      if (!headerFrom) headerFrom = rowNum + 1;
      headerTo = Math.max(headerTo, rowNum + rrep);
    }
    // Keep the original row element per expanded index so its attributes
    // (visibility, default-cell-style-name, ...) survive a rebuild.
    const worthKeeping = Array.from(rowEl.attributes).some((a) => a.name !== "table:number-rows-repeated");
    if (worthKeeping) for (let k = 0; k < Math.min(rrep, REPEAT_CAP); k++) rowEls.set(rowNum + 1 + k, rowEl);
    // A content run repeated beyond the cap: the un-expanded tail is preserved verbatim.
    if (rowHasContent && rrep > REPEAT_CAP) rowRuns.push({ from: rowNum + 1 + REPEAT_CAP, to: rowNum + rrep, el: rowEl });
    const rowHidden = isHidden(rowEl);
    for (let k = 0; k < copies; k++) {
      const r = rowNum + 1 + k;
      if (rh) rowHeights.set(r, rh);
      if (rowHidden) hiddenRows.add(r);
      if (rowStyle) rowStyles.set(r, rowStyle);
      for (const pc of parsedCells) {
        if (!pc.has) {
          // A covered (merged-away) cell may still carry content/attributes; keep it.
          if (pc.coveredEl) for (let j = 0; j < pc.span; j++) coveredEls.set(key(r, pc.startCol + j), pc.coveredEl);
          continue;
        }
        const c = pc.cell!;
        if ((pc.colSpan ?? 1) > 1 || (pc.rowSpan ?? 1) > 1) {
          merges.push({ r1: r, c1: c.col, r2: r + (pc.rowSpan ?? 1) - 1, c2: c.col + (pc.colSpan ?? 1) - 1 });
        }
        for (let j = 0; j < pc.span; j++) {
          sheet.cells.set(key(r, c.col + j), { ...c, row: r, col: c.col + j });
          noteExtent(sheet, r, c.col + j);
        }
      }
    }
    rowNum += rrep;
  }
  if (rowHeights.size) sheet.rowHeights = rowHeights;
  if (hiddenRows.size) sheet.hiddenRows = hiddenRows;
  if (rowStyles.size) sheet.odsRowStyles = rowStyles;
  if (rowEls.size) sheet.odsRowEls = rowEls;
  if (rowRuns.length) sheet.odsRowRuns = rowRuns;
  if (coveredEls.size) sheet.odsCoveredEls = coveredEls;
  if (headerGroupEl && headerFrom) sheet.odsHeaderRows = { el: headerGroupEl, from: headerFrom, to: headerTo };
  if (merges.length) sheet.merges = merges;
}

export interface ParsedOdsCell {
  has: boolean;
  span: number;
  startCol: number;
  colSpan?: number;
  rowSpan?: number;
  cell?: Cell;
  /** A covered-table-cell that carries content or attributes, preserved on save. */
  coveredEl?: Element;
}

// Per-run rich text: the first <text:p>'s text nodes + <text:span> runs, each span's style resolved
// to bold/italic/underline/strike/colour/size/font. Returned only when there is real per-run
// variation (>1 run and at least one styled span), mirroring the xlsx rich-text path.
function odsCellRuns(cellEl: Element, styles: OdsStyles): TextRun[] | undefined {
  const p = Array.from(cellEl.children).find((c) => c.localName === "p");
  if (!p) return undefined;
  const runs: TextRun[] = [];
  let styled = 0;
  for (const node of Array.from(p.childNodes)) {
    if (node.nodeType === 3) { const t = node.textContent ?? ""; if (t) runs.push({ text: t }); continue; }
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    if (el.localName === "span") {
      const cs = styles.text.get(el.getAttribute("text:style-name") ?? "") ?? {};
      const run: TextRun = { text: el.textContent ?? "" };
      if (cs.bold) run.bold = true;
      if (cs.italic) run.italic = true;
      if (cs.underline) run.underline = true;
      if (cs.strike) run.strike = true;
      if (cs.color) run.color = cs.color;
      if (cs.fontSize) run.size = cs.fontSize;
      if (cs.fontFamily) run.font = cs.fontFamily;
      if (run.bold || run.italic || run.underline || run.strike || run.color || run.size || run.font) styled++;
      if (run.text) runs.push(run);
    } else if (el.localName === "a") {
      runs.push({ text: el.textContent ?? "" });
    } else if (el.localName === "s") {
      runs.push({ text: " ".repeat(Math.max(1, Number(el.getAttribute("text:c") || "1"))) });
    }
  }
  return styled > 0 && runs.length > 1 ? runs : undefined;
}

export function parseOdsRow(rowEl: Element, styles: OdsStyles): ParsedOdsCell[] {
  const out: ParsedOdsCell[] = [];
  let col = 0;
  for (const cellEl of Array.from(rowEl.children)) {
    const local = cellEl.localName;
    if (local !== "table-cell" && local !== "covered-table-cell") continue;
    const crep = Math.max(1, Number(cellEl.getAttribute("table:number-columns-repeated") || "1"));
    const startCol = col + 1;
    col += crep;
    if (local === "covered-table-cell") {
      // Merged-away cell: keep it when it carries anything beyond the repeat count.
      const worth = cellEl.children.length > 0 || Array.from(cellEl.attributes).some((a) => a.name !== "table:number-columns-repeated");
      out.push({ has: false, span: Math.min(crep, REPEAT_CAP), startCol, coveredEl: worth ? cellEl : undefined });
      continue;
    }
    const colSpan = Math.max(1, Number(cellEl.getAttribute("table:number-columns-spanned") || "1"));
    const rowSpan = Math.max(1, Number(cellEl.getAttribute("table:number-rows-spanned") || "1"));
    const valueType = cellEl.getAttribute("office:value-type");
    const formulaRaw = cellEl.getAttribute("table:formula") ?? undefined;
    const style = cellEl.getAttribute("table:style-name") ?? undefined;
    const text = odsCellText(cellEl);
    let value = "";
    let kind: CellKind = "blank";
    let display: string | undefined;
    let numFmt: string | undefined;
    let odsValueType: string | undefined;
    let odsCurrency: string | undefined;
    let phonetic: Phonetic[] | undefined;
    if (valueType === "float" || valueType === "percentage" || valueType === "currency") {
      value = cellEl.getAttribute("office:value") ?? text;
      // ODF stores the producer's formatted text in <text:p>; use it as the display.
      if (text !== "" && text !== value) display = text;
      kind = "n";
      if (valueType !== "float") {
        odsValueType = valueType;
        odsCurrency = cellEl.getAttribute("office:currency") ?? undefined;
        if (valueType === "percentage") numFmt = "0.00%";
      }
    } else if (valueType === "boolean") {
      value = cellEl.getAttribute("office:boolean-value") === "true" ? "TRUE" : "FALSE";
      kind = "b";
    } else if (valueType === "string") {
      const rich = odsCellRich(cellEl); // base text + furigana (text:ruby), not the reading folded in
      value = cellEl.getAttribute("office:string-value") ?? rich.text;
      phonetic = rich.phonetic;
      kind = "s";
    } else if (valueType === "date") {
      // Model dates as serials so arithmetic and re-formatting work; the
      // producer's formatted <text:p> stays as the display.
      const iso = cellEl.getAttribute("office:date-value") ?? "";
      const serial = iso === "" ? null : isoToSerial(iso);
      if (serial != null) {
        value = numToStr(serial);
        kind = "n";
        odsValueType = "date";
        numFmt = iso.includes("T") ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd";
      } else {
        value = iso || text;
        kind = value === "" ? "blank" : "s";
      }
      if (text !== "" && text !== value) display = text;
    } else if (valueType === "time") {
      const dur = cellEl.getAttribute("office:time-value") ?? "";
      const serial = dur === "" ? null : durationToSerial(dur);
      if (serial != null) {
        value = numToStr(serial);
        kind = "n";
        odsValueType = "time";
        numFmt = "hh:mm:ss";
      } else {
        value = dur || text;
        kind = value === "" ? "blank" : "s";
      }
      if (text !== "" && text !== value) display = text;
    } else {
      const rich = odsCellRich(cellEl);
      value = rich.text;
      phonetic = rich.phonetic;
      kind = value === "" ? "blank" : "s";
    }
    const link = odsCellLink(cellEl);
    const comments = odsCellComments(cellEl);
    const richRuns = kind === "s" && !phonetic ? odsCellRuns(cellEl, styles) : undefined;
    const odsValidationName = cellEl.getAttribute("table:content-validation-name") ?? undefined;
    const has = value !== "" || formulaRaw != null || style != null || link != null || comments != null || odsValidationName != null;
    if (!has) {
      out.push({ has: false, span: crep, startCol });
      continue;
    }
    const cell: Cell = {
      row: 0,
      col: startCol,
      value,
      kind,
      display: display ?? (kind === "n" && numFmt != null ? formatNumber(numFmt, value) ?? undefined : undefined),
      numFmt,
      odsValueType,
      odsCurrency,
      formula: formulaRaw ? odfToA1(formulaRaw) : undefined,
      odfFormula: formulaRaw,
      style,
      cellStyle: style ? styles.cell.get(style) : undefined,
      el: cellEl,
      phonetic,
      link,
      comments,
      richRuns,
      odsValidationName,
    };
    out.push({ has: true, span: Math.min(crep, REPEAT_CAP), startCol, colSpan, rowSpan, cell });
  }
  return out;
}

