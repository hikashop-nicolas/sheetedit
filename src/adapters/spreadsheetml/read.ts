// Excel 2003 XML Spreadsheet (SpreadsheetML): Workbook / Worksheet / Table / Column / Row / Cell with
// styles, R1C1 formulas, names, panes, comments, validations, conditional formats and filters.
import { dateToSerial } from "../../core/dates";
import { DEFAULT_PALETTE } from "../xls/palette";
import { cssColor, excelNumFmt, parseDecls, declsToStyle } from "../tabular/css";
import { cellKey, newSheet, uniqueSheetNames, type TBook, type TBorder, type TCell, type TFont, type TRange, type TRun, type TSheet, type TStyle, type TValidation, type TValue } from "../tabular/model";
import { rangeA1 } from "../tabular/to-xlsx";
import { childEls, firstEl, isNode, parseXmlTree, textOf, type HNode } from "../tabular/tree";
import { r1c1Range, r1c1ToA1 } from "./r1c1";

export interface SpreadsheetMlReadOptions {
  /** Parse with DOMParser when available (default true); false forces the built-in tokenizer. */
  useDom?: boolean;
  sheetName?: (n: number) => string;
}

const at = (n: HNode | undefined, name: string): string | undefined => n?.attrs[name.toLowerCase()];
const bool = (v: string | undefined): boolean => v === "1" || v?.toLowerCase() === "true";
const numAt = (n: HNode | undefined, name: string): number | undefined => {
  const v = at(n, name);
  if (v == null || v === "") return undefined;
  const x = Number(v);
  return Number.isFinite(x) ? x : undefined;
};

// ---- styles -------------------------------------------------------------------------------------

const PATTERNS: Record<string, string> = {
  solid: "solid", gray75: "darkGray", gray50: "mediumGray", gray25: "lightGray", gray125: "gray125", gray0625: "gray0625",
  horzstripe: "darkHorizontal", vertstripe: "darkVertical", reversediagstripe: "darkDown", diagstripe: "darkUp", diagcross: "darkGrid",
  thickdiagcross: "darkTrellis", thinhorzstripe: "lightHorizontal", thinvertstripe: "lightVertical", thinreversediagstripe: "lightDown",
  thindiagstripe: "lightUp", thinhorzcross: "lightGrid", thindiagcross: "lightTrellis",
};
const H_ALIGN: Record<string, string> = { left: "left", center: "center", right: "right", fill: "fill", justify: "justify", centeracrossselection: "centerContinuous", distributed: "distributed", justifydistributed: "distributed" };
const V_ALIGN: Record<string, string> = { top: "top", center: "center", justify: "justify", distributed: "distributed", justifydistributed: "distributed" };

function borderOf(b: HNode): TBorder | undefined {
  const line = (at(b, "LineStyle") ?? "").toLowerCase();
  if (!line || line === "none") return undefined;
  const w = numAt(b, "Weight") ?? 1;
  let style: string;
  switch (line) {
    case "double": style = "double"; break;
    case "dot": style = w >= 2 ? "mediumDashDot" : w === 0 ? "hair" : "dotted"; break;
    case "dash": style = w >= 2 ? "mediumDashed" : "dashed"; break;
    case "dashdot": style = w >= 2 ? "mediumDashDot" : "dashDot"; break;
    case "dashdotdot": style = w >= 2 ? "mediumDashDotDot" : "dashDotDot"; break;
    case "slantdashdot": style = "slantDashDot"; break;
    default: style = w >= 3 ? "thick" : w >= 2 ? "medium" : w === 0 ? "hair" : "thin";
  }
  const color = cssColor(at(b, "Color"));
  return { style, color: color && color !== "#000000" ? color : undefined };
}

function fontOf(f: HNode): TFont {
  const font: TFont = {};
  const name = at(f, "FontName");
  if (name) font.name = name;
  const size = numAt(f, "Size");
  if (size) font.size = size;
  if (bool(at(f, "Bold"))) font.bold = true;
  if (bool(at(f, "Italic"))) font.italic = true;
  if (bool(at(f, "StrikeThrough"))) font.strike = true;
  const u = (at(f, "Underline") ?? "").toLowerCase();
  if (u === "single") font.underline = "single";
  else if (u === "double") font.underline = "double";
  else if (u === "singleaccounting") font.underline = "singleAccounting";
  else if (u === "doubleaccounting") font.underline = "doubleAccounting";
  const va = (at(f, "VerticalAlign") ?? "").toLowerCase();
  if (va === "superscript") font.vertAlign = "superscript";
  else if (va === "subscript") font.vertAlign = "subscript";
  const color = cssColor(at(f, "Color"));
  if (color && color !== "#000000") font.color = color;
  return font;
}

/** One <Style>'s own settings, before its parent and the Default style are merged in. */
function ownStyle(s: HNode): TStyle {
  const st: TStyle = {};
  for (const k of childEls(s)) {
    switch (k.tag) {
      case "font": st.font = fontOf(k); break;
      case "interior": {
        const color = cssColor(at(k, "Color"));
        const pat = (at(k, "Pattern") ?? (color ? "solid" : "none")).toLowerCase();
        const kind = PATTERNS[pat];
        if (kind === "solid" && color) st.fill = { pattern: "solid", fg: color };
        // Solid with no colour is the automatic (white) background, i.e. no fill.
        else if (kind && kind !== "solid") st.fill = { pattern: kind, fg: cssColor(at(k, "PatternColor")) ?? "#000000", bg: color };
        else if (kind === "solid") st.fill = { pattern: "none" };
        else if (pat === "none") st.fill = { pattern: "none" };
        break;
      }
      case "borders": {
        const border: NonNullable<TStyle["border"]> = {};
        for (const b of childEls(k, "border")) {
          const side = (at(b, "Position") ?? "").toLowerCase();
          const v = borderOf(b);
          if (!v) continue;
          if (side === "left" || side === "right" || side === "top" || side === "bottom") border[side] = v;
          else if (side === "diagonalleft") { border.diagonal = v; border.diagDown = true; }
          else if (side === "diagonalright") { border.diagonal = v; border.diagUp = true; }
        }
        st.border = border;
        break;
      }
      case "alignment": {
        const al: NonNullable<TStyle["align"]> = {};
        const h = H_ALIGN[(at(k, "Horizontal") ?? "").toLowerCase()];
        if (h) al.h = h;
        const v = V_ALIGN[(at(k, "Vertical") ?? "").toLowerCase()];
        if (v) al.v = v;
        if (bool(at(k, "WrapText"))) al.wrap = true;
        if (bool(at(k, "ShrinkToFit"))) al.shrink = true;
        const ind = numAt(k, "Indent");
        if (ind) al.indent = ind;
        const rot = numAt(k, "Rotate");
        if (rot) al.rotate = rot > 0 ? Math.min(90, rot) : Math.min(180, 90 - rot);
        if (bool(at(k, "VerticalText"))) al.rotate = 255;
        st.align = al;
        break;
      }
      case "numberformat": {
        const f = excelNumFmt(at(k, "Format"));
        st.numFmt = f ?? "General";
        break;
      }
      case "protection":
        if (at(k, "Protected") != null) st.locked = bool(at(k, "Protected"));
        if (bool(at(k, "HideFormula"))) st.hidden = true;
        break;
    }
  }
  return st;
}

function mergeStyle(base: TStyle, over: TStyle): TStyle {
  const out: TStyle = { ...base };
  if (over.font) out.font = { ...base.font, ...over.font };
  if (over.fill) out.fill = over.fill;
  if (over.border) out.border = { ...base.border, ...over.border };
  if (over.align) out.align = { ...base.align, ...over.align };
  if (over.numFmt != null) out.numFmt = over.numFmt;
  if (over.locked != null) out.locked = over.locked;
  if (over.hidden != null) out.hidden = over.hidden;
  return out;
}

/** Drop the parts that only restate the defaults, so unstyled cells share the normal style. */
function tidy(st: TStyle, dflt: { name: string; size: number }): TStyle | undefined {
  const out: TStyle = { ...st };
  if (out.font) {
    const f = { ...out.font };
    if (f.name === dflt.name) delete f.name;
    if (f.size === dflt.size) delete f.size;
    out.font = Object.keys(f).length ? f : undefined;
    if (!out.font) delete out.font;
  }
  if (out.fill?.pattern === "none") delete out.fill;
  if (out.border && !Object.keys(out.border).length) delete out.border;
  if (out.align && !Object.keys(out.align).length) delete out.align;
  if (out.numFmt === "General") delete out.numFmt;
  if (out.locked === true) delete out.locked;
  return Object.keys(out).length ? out : undefined;
}

// ---- cell data ------------------------------------------------------------------------------------

/** Rich text inside <Data> (html:B, html:Font...) as runs; plain text gives none. */
function dataRuns(data: HNode): TRun[] | undefined {
  if (!data.kids.some(isNode)) return undefined;
  const runs: TRun[] = [];
  const walk = (n: HNode, font: TFont): void => {
    for (const k of n.kids) {
      if (typeof k === "string") { if (k) runs.push({ text: k, font: Object.keys(font).length ? font : undefined }); continue; }
      const f: TFont = { ...font };
      switch (k.tag) {
        case "b": f.bold = true; break;
        case "i": f.italic = true; break;
        case "u": f.underline = "single"; break;
        case "s": f.strike = true; break;
        case "sup": f.vertAlign = "superscript"; break;
        case "sub": f.vertAlign = "subscript"; break;
        case "font": {
          const face = at(k, "Face"), size = numAt(k, "Size"), color = cssColor(at(k, "Color"));
          if (face) f.name = face;
          if (size) f.size = size;
          if (color && color !== "#000000") f.color = color;
          break;
        }
        case "span": {
          const st = declsToStyle(parseDecls(at(k, "style") ?? ""), { baseSize: 10, excel: true });
          Object.assign(f, st.font);
          break;
        }
        case "br": runs.push({ text: "\n", font: Object.keys(font).length ? font : undefined }); continue;
      }
      walk(k, f);
    }
  };
  walk(data, {});
  return runs.length > 1 ? runs : undefined;
}

function dateSerial(v: string, date1904: boolean): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(v.trim());
  if (!m) return null;
  let y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  // Excel writes pure times on its day zero, 1899-12-31.
  if (y === 1899 && mo === 12 && d === 31) return (Number(m[4] ?? 0) * 3600 + Number(m[5] ?? 0) * 60 + Number(m[6] ?? 0)) / 86400;
  if (y < 1900) { y = 1900; mo = 1; d = 1; }
  let s = dateToSerial(y, mo, d, Number(m[4] ?? 0), Number(m[5] ?? 0), 0) + (Number(m[6] ?? 0)) / 86400;
  if (date1904) s -= 1462;
  return s;
}

// ---- worksheet options, validation, conditional formats --------------------------------------

const VAL_TYPE: Record<string, TValidation["type"]> = { whole: "whole", decimal: "decimal", list: "list", date: "date", time: "time", textlength: "textLength", custom: "custom" };
const QUALIFIER: Record<string, string> = { between: "between", notbetween: "notBetween", equal: "equal", notequal: "notEqual", greater: "greaterThan", less: "lessThan", greaterorequal: "greaterThanOrEqual", lessorequal: "lessThanOrEqual" };

const ranges = (spec: string | undefined, row = 0, col = 0): TRange[] =>
  (spec ?? "").split(",").map((p) => r1c1Range(p, row, col)).filter((r): r is TRange => !!r);

/** A formula from a validation or condition: R1C1 relative to the range's first cell. */
function relFormula(text: string | undefined, rs: TRange[]): string | undefined {
  if (text == null) return undefined;
  const t = text.trim().replace(/^=/, "");
  return r1c1ToA1(t, rs[0]?.r1 ?? 0, rs[0]?.c1 ?? 0);
}

function readValidation(n: HNode, sheet: TSheet): void {
  const t = (tag: string): string | undefined => { const k = firstEl(n, tag.toLowerCase()); return k ? textOf(k) : undefined; };
  const rs = ranges(t("Range"));
  if (!rs.length) return;
  const type = VAL_TYPE[(t("Type") ?? "").toLowerCase()] ?? (t("CellRangeList") != null ? "list" : undefined);
  if (!type) return;
  const qual = QUALIFIER[(t("Qualifier") ?? "").toLowerCase()];
  const v: TValidation = { ranges: rs, type };
  if (v.type === "list") {
    const crl = t("CellRangeList"), val = t("Value");
    if (crl != null) v.f1 = relFormula(crl, rs);
    else if (val != null) {
      const raw = val.trim();
      v.f1 = raw.startsWith('"') ? raw : /^=|!|^R(\[|\d|C)|^C(\[|\d)/.test(raw) ? relFormula(raw, rs) : `"${raw}"`;
    } else return;
  } else {
    const min = t("Min"), max = t("Max"), val = t("Value");
    v.op = qual ?? (min != null && max != null ? "between" : "equal");
    if (v.op === "between" || v.op === "notBetween") { v.f1 = relFormula(min ?? val, rs); v.f2 = relFormula(max, rs); }
    else v.f1 = relFormula(val ?? min ?? max, rs);
    if (v.f1 == null) return;
  }
  if (firstEl(n, "useblank")) v.allowBlank = true;
  if (firstEl(n, "combohide")) v.noDropdown = true;
  v.promptTitle = t("InputTitle");
  v.prompt = t("InputMessage");
  v.errorTitle = t("ErrorTitle");
  v.error = t("ErrorMessage");
  const es = (t("ErrorStyle") ?? "").toLowerCase();
  if (es === "warn") v.errorStyle = "warning";
  else if (es === "info") v.errorStyle = "information";
  sheet.validations.push(v);
}

function readCondFormat(n: HNode, sheet: TSheet): void {
  const rangeEl = firstEl(n, "range");
  const rs = ranges(rangeEl ? textOf(rangeEl) : undefined);
  if (!rs.length) return;
  const rules = childEls(n, "condition").map((c) => {
    const t = (tag: string): string | undefined => { const k = firstEl(c, tag); return k ? textOf(k) : undefined; };
    const qual = QUALIFIER[(t("qualifier") ?? "").toLowerCase()];
    const fmt = firstEl(c, "format");
    const style = declsToStyle(parseDecls(at(fmt, "Style") ?? ""), { baseSize: 10, excel: true });
    // Conditions colour the cell background, which a dxf states as a solid fill.
    if (style.fill && style.fill.pattern === "solid") style.fill = { pattern: "solid", fg: style.fill.fg };
    const f1 = relFormula(t("value1"), rs);
    if (f1 == null) return null;
    return qual ? { kind: "cellIs" as const, op: qual, f1, f2: relFormula(t("value2"), rs), style } : { kind: "expression" as const, f1, style };
  }).filter((r): r is NonNullable<typeof r> => !!r);
  if (rules.length) sheet.condFormats.push({ ranges: rs, rules });
}

function readOptions(o: HNode, sheet: TSheet): { selected: boolean } {
  const t = (tag: string): string | undefined => { const k = firstEl(o, tag); return k ? textOf(k).trim() : undefined; };
  if (firstEl(o, "freezepanes") || firstEl(o, "frozennosplit")) {
    const rows = Number(t("splithorizontal") ?? 0), cols = Number(t("splitvertical") ?? 0);
    if (rows || cols) sheet.freeze = { rows, cols, top: Number(t("toprowbottompane") ?? rows) || rows, left: Number(t("leftcolumnrightpane") ?? cols) || cols };
  }
  const vis = (t("visible") ?? "").toLowerCase();
  if (vis === "sheethidden") sheet.state = "hidden";
  else if (vis === "sheetveryhidden") sheet.state = "veryHidden";
  const tab = Number(t("tabcolorindex") ?? NaN);
  if (tab >= 0 && tab < 64) sheet.tabColor = DEFAULT_PALETTE[tab < 8 ? tab : tab - 8];
  if (firstEl(o, "donotdisplaygridlines")) sheet.hideGrid = true;
  if (firstEl(o, "donotdisplayzeros")) sheet.hideZeros = true;
  if (firstEl(o, "displayrighttoleft")) sheet.rtl = true;
  const print: NonNullable<TSheet["print"]> = {};
  const ps = firstEl(o, "pagesetup");
  if (ps) {
    const layout = firstEl(ps, "layout");
    if ((at(layout, "Orientation") ?? "").toLowerCase() === "landscape") print.landscape = true;
    if (bool(at(layout, "CenterHorizontal"))) print.hCenter = true;
    if (bool(at(layout, "CenterVertical"))) print.vCenter = true;
    const header = firstEl(ps, "header"), footer = firstEl(ps, "footer"), pm = firstEl(ps, "pagemargins");
    if (at(header, "Data")) print.header = at(header, "Data");
    if (at(footer, "Data")) print.footer = at(footer, "Data");
    if (pm || header || footer) print.margins = { left: numAt(pm, "Left"), right: numAt(pm, "Right"), top: numAt(pm, "Top"), bottom: numAt(pm, "Bottom"), header: numAt(header, "Margin"), footer: numAt(footer, "Margin") };
  }
  const pr = firstEl(o, "print");
  if (pr) {
    const pt = (tag: string): number | undefined => { const k = firstEl(pr, tag); const v = k ? Number(textOf(k)) : NaN; return Number.isFinite(v) ? v : undefined; };
    print.fitWidth = pt("fitwidth");
    print.fitHeight = pt("fitheight");
    print.scale = pt("scale");
    print.paperSize = pt("papersizeindex");
    if (firstEl(pr, "gridlines")) print.gridLines = true;
  }
  if (firstEl(o, "fittopage")) print.fitToPage = true;
  if (Object.values(print).some((v) => v != null)) sheet.print = print;
  return { selected: !!firstEl(o, "selected") };
}

// ---- the reader ------------------------------------------------------------------------------------

/** Read an Excel 2003 XML Spreadsheet into the intermediate book. */
export function readSpreadsheetMlBook(text: string, opts: SpreadsheetMlReadOptions = {}): TBook {
  const base = opts.sheetName ?? ((n: number) => `Sheet${n}`);
  const tree = parseXmlTree(text, opts.useDom !== false);
  const wbEl = childEls(tree).find((k) => k.tag === "workbook") ?? tree;
  const book: TBook = { sheets: [], names: [] };
  const xlWb = firstEl(wbEl, "excelworkbook");
  const date1904 = !!(xlWb && firstEl(xlWb, "date1904"));
  if (date1904) book.date1904 = true;
  // Styles, with Default as everyone's implicit parent.
  const styleEls = new Map<string, HNode>();
  for (const s of childEls(firstEl(wbEl, "styles") ?? { tag: "", attrs: {}, kids: [] }, "style")) {
    const id = at(s, "ID");
    if (id) styleEls.set(id, s);
  }
  const resolved = new Map<string, TStyle>();
  const resolve = (id: string, depth = 0): TStyle => {
    const hit = resolved.get(id);
    if (hit) return hit;
    const el = styleEls.get(id);
    if (!el || depth > 20) return {};
    const parentId = at(el, "Parent") ?? (id === "Default" ? undefined : "Default");
    const parent = parentId && parentId !== id ? resolve(parentId, depth + 1) : {};
    const st = mergeStyle(parent, ownStyle(el));
    resolved.set(id, st);
    return st;
  };
  const dfltStyle = resolve("Default");
  const dflt = { name: dfltStyle.font?.name ?? "Calibri", size: dfltStyle.font?.size ?? 11 };
  book.defaultFont = dflt;
  const tidied = new Map<string, TStyle | undefined>();
  const styleFor = (id: string | undefined): TStyle | undefined => {
    const key = id && styleEls.has(id) ? id : "Default";
    if (!tidied.has(key)) tidied.set(key, tidy(resolve(key), dflt));
    return tidied.get(key);
  };
  const isDateStyle = (st: TStyle | undefined): boolean => /[dmyhs]/i.test((st?.numFmt ?? "").replace(/"[^"]*"|\[[^\]]*\]/g, ""));

  const worksheets = childEls(wbEl, "worksheet");
  let selected = -1;
  const localNames: { el: HNode; sheet: number }[] = [];
  worksheets.forEach((ws, si) => {
    const sheet = newSheet(at(ws, "Name") ?? "");
    book.sheets.push(sheet);
    const table = firstEl(ws, "table");
    if (table) {
      const dcw = numAt(table, "DefaultColumnWidth");
      if (dcw != null) sheet.defaultColWidth = Math.round((dcw * 4) / 3);
      const drh = numAt(table, "DefaultRowHeight");
      if (drh != null) sheet.defaultRowHeight = drh;
      const colStyles = new Map<number, string>();
      let c = 0;
      for (const col of childEls(table, "column")) {
        const idx = numAt(col, "Index");
        if (idx) c = idx - 1;
        const span = Math.min(16384, (numAt(col, "Span") ?? 0) + 1);
        const w = numAt(col, "Width");
        const hidden = bool(at(col, "Hidden"));
        const sid = at(col, "StyleID");
        for (let i = 0; i < span; i++, c++) {
          if (w != null || hidden || sid) sheet.cols.set(c, { width: w != null ? Math.round((w * 4) / 3) : undefined, hidden: hidden || undefined, style: sid ? styleFor(sid) : undefined });
          if (sid) colStyles.set(c, sid);
        }
      }
      const arrays: TRange[] = [];
      let r = 0;
      for (const row of childEls(table, "row")) {
        const idx = numAt(row, "Index");
        if (idx) r = idx - 1;
        const span = Math.min(1048576, (numAt(row, "Span") ?? 0) + 1);
        const h = numAt(row, "Height");
        const hidden = bool(at(row, "Hidden"));
        const rowSid = at(row, "StyleID");
        const customH = h != null && at(row, "AutoFitHeight") !== "1";
        for (let i = 0; i < span; i++) if (customH || hidden || rowSid) sheet.rows.set(r + i, { height: customH ? h : undefined, hidden: hidden || undefined, style: rowSid ? styleFor(rowSid) : undefined });
        let cc = 0;
        for (const cellEl of childEls(row, "cell")) {
          const ci = numAt(cellEl, "Index");
          if (ci) cc = ci - 1;
          const sid = at(cellEl, "StyleID") ?? rowSid ?? colStyles.get(cc);
          const style = styleFor(sid);
          const cell: TCell = { r, c: cc, style };
          const data = firstEl(cellEl, "data") ?? childEls(cellEl).find((k) => k.tag === "data");
          if (data) {
            const type = (at(data, "Type") ?? "String").toLowerCase();
            const raw = textOf(data);
            let v: TValue | undefined;
            if (type === "number") { const n = Number(raw.trim()); v = Number.isFinite(n) && raw.trim() !== "" ? { t: "n", v: n } : { t: "s", v: raw }; }
            else if (type === "boolean") v = { t: "b", v: raw.trim() === "1" || raw.trim().toLowerCase() === "true" };
            else if (type === "error") v = { t: "e", v: raw.trim() };
            else if (type === "datetime") {
              const n = dateSerial(raw, date1904);
              v = n != null ? { t: "n", v: n } : { t: "s", v: raw };
              if (n != null && !isDateStyle(style)) cell.style = { ...style, numFmt: n < 1 ? "h:mm:ss" : n % 1 ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd" };
            } else {
              const runs = dataRuns(data);
              v = { t: "s", v: raw, runs };
            }
            cell.v = v;
          }
          const formula = at(cellEl, "Formula");
          if (formula && formula.startsWith("=")) cell.f = r1c1ToA1(formula.slice(1), r, cc);
          const arr = at(cellEl, "ArrayRange");
          if (arr && cell.f) {
            const rg = r1c1Range(arr, r, cc);
            if (rg) { cell.array = rangeA1(rg); arrays.push(rg); }
          }
          const href = at(cellEl, "HRef");
          if (href) cell.link = href.startsWith("#") ? { location: href.slice(1), tip: at(cellEl, "HRefScreenTip") } : { href, tip: at(cellEl, "HRefScreenTip") };
          const comment = firstEl(cellEl, "comment");
          if (comment) {
            const ctext = textOf(firstEl(comment, "data") ?? comment).replace(/\r\n?/g, "\n").trim();
            if (ctext) cell.comment = { author: at(comment, "Author") || undefined, text: ctext };
          }
          const ma = numAt(cellEl, "MergeAcross") ?? 0, md = numAt(cellEl, "MergeDown") ?? 0;
          if (ma > 0 || md > 0) sheet.merges.push({ r1: r, c1: cc, r2: r + Math.max(0, md), c2: cc + Math.max(0, ma) });
          if (cell.v || cell.f || cell.style || cell.link || cell.comment) sheet.cells.set(cellKey(r, cc), cell);
          cc += 1 + Math.max(0, ma);
        }
        r += span;
      }
      // Array formulas: only the top-left cell keeps the formula, the rest hold cached values.
      for (const rg of arrays) for (const cell of sheet.cells.values())
        if (cell.r >= rg.r1 && cell.r <= rg.r2 && cell.c >= rg.c1 && cell.c <= rg.c2 && (cell.r !== rg.r1 || cell.c !== rg.c1)) { delete cell.f; delete cell.array; }
    }
    for (const k of childEls(ws)) {
      if (k.tag === "worksheetoptions") { if (readOptions(k, sheet).selected && selected < 0) selected = si; }
      else if (k.tag === "datavalidation") readValidation(k, sheet);
      else if (k.tag === "conditionalformatting") readCondFormat(k, sheet);
      else if (k.tag === "autofilter") { const rg = ranges(at(k, "Range"))[0]; if (rg) sheet.autoFilter = rg; }
      else if (k.tag === "names") for (const nr of childEls(k, "namedrange")) localNames.push({ el: nr, sheet: si });
    }
  });
  if (!book.sheets.length) book.sheets.push(newSheet(""));
  const names = uniqueSheetNames(book.sheets.map((s) => s.name), base);
  const renamed = new Map<string, string>();
  book.sheets.forEach((s, i) => { if (s.name && s.name !== names[i]) renamed.set(s.name, names[i]!); s.name = names[i]!; });
  const active = xlWb ? Number(textOf(firstEl(xlWb, "activesheet") ?? { tag: "", attrs: {}, kids: ["NaN"] })) : NaN;
  if (Number.isFinite(active) && active >= 0 && active < book.sheets.length) book.activeSheet = active;
  else if (selected > 0) book.activeSheet = selected;
  const addName = (nr: HNode, sheet?: number): void => {
    const name = at(nr, "Name"), ref = at(nr, "RefersTo");
    if (!name || !ref) return;
    const builtin = /^(Print_Area|Print_Titles|_FilterDatabase|Criteria|Extract)$/i.test(name);
    book.names.push({ name: builtin ? "_xlnm." + name : name, ref: r1c1ToA1(ref.replace(/^=/, "")), sheet, hidden: bool(at(nr, "Hidden")) || undefined });
  };
  for (const nr of childEls(firstEl(wbEl, "names") ?? { tag: "", attrs: {}, kids: [] }, "namedrange")) addName(nr);
  for (const { el, sheet } of localNames) addName(el, sheet);
  if (renamed.size) for (const n of book.names) for (const [from, to] of renamed) n.ref = n.ref.split(`${from}!`).join(`'${to.replace(/'/g, "''")}'!`);
  return book;
}
