// BIFF8 writer: a workbook model (with its xlsx style pool) as an Excel 97-2003 .xls compound file.
// Styles are rebuilt from the live styles.xml so every edit made through the xlsx paths lands in
// the XF / FONT / FORMAT tables; formulas are compiled to tokens with their cached results.
import type { Cell, CfRule, Sheet, SheetControl, SheetShape, TextRun, Workbook } from "../../core/model";
import { activeXObjBody, CAPTION_GRBIT, ctlObjBody, type CtlWriteEnv, newControlShape, ungroupedShape, withPib } from "./control";
import { geomOf } from "../xlsx/shape-read";
import { type ChartWriteEnv, writeChartStream } from "./chart-write";
import { sptOf } from "./drawing";
import { colToLetters, firstByLocal } from "../../core/model";
import { SHEET_LOCK_DEFAULTS, type SheetLock } from "../../core/protection";
import { argbToCss, colWidthToPx, maxDigitWidth, readTheme, resolveColor } from "../xlsx/read";
import { colWidthChars, timePeriodFormula } from "../xlsx/write";
import { body, Bytes, MAX_REC, RecStream } from "./bytes";
import { readCfbAll, writeCfbStreams } from "./cfb";
import { vbaPartOf } from "../../core/vba";
import { md5 } from "./crypto";
import { CompileError, compileFormula, type CompileEnv, type Ctx } from "./formula-write";
import { DEFAULT_PALETTE, icvCss, PaletteBuilder } from "./palette";
import { BORDER_STYLES } from "./to-xlsx";

const PATTERNS = ["none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"];
const H_ALIGN = ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"];
const V_ALIGN = ["top", "center", "bottom", "justify", "distributed"];
const ERR_CODE: Record<string, number> = { "#NULL!": 0x00, "#DIV/0!": 0x07, "#VALUE!": 0x0f, "#REF!": 0x17, "#NAME?": 0x1d, "#NUM!": 0x24, "#N/A": 0x2a, "#GETTING_DATA": 0x2b };
const LOCK_ORDER: SheetLock[] = ["objects", "scenarios", "formatCells", "formatColumns", "formatRows", "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows", "selectLockedCells", "sort", "autoFilter", "pivotTables", "selectUnlockedCells"];

/** What the BIFF8 writer could not keep, so a host can tell the user. */
export interface XlsWriteReport { droppedFormulas: number; droppedImages: number; droppedRules: number; truncatedCells: number }

interface FontKey { height: number; italic: boolean; strike: boolean; weight: number; underline: number; escapement: number; color: number; name: string }

const MAX_ROWS = 65536, MAX_COLS = 256;

/** Write the workbook as BIFF8. `report` (optional) collects what had to be dropped. */
export function writeXls(wb: Workbook, report: XlsWriteReport = { droppedFormulas: 0, droppedImages: 0, droppedRules: 0, truncatedCells: 0 }): Uint8Array {
  const doc = wb.stylesDoc;
  const vbaBin = vbaPartOf(wb.files);
  const theme = readTheme(wb.files["xl/theme/theme1.xml"]);
  const srcPalette = wb.xls?.palette ?? DEFAULT_PALETTE;
  const palette = new PaletteBuilder(srcPalette);
  // Colours stated by palette index stay where they are.
  if (doc) for (const el of Array.from(doc.getElementsByTagName("*"))) { const ix = el.getAttribute("indexed"); if (ix != null) palette.reserve(Number(ix)); }
  const icvOf = (el: Element | undefined, auto: number): number => {
    if (!el) return auto;
    if (el.getAttribute("auto") === "1" || el.getAttribute("auto") === "true") return auto;
    const css = argbToCss(el.getAttribute("rgb")) ?? resolveColor(el, theme);
    const ix = el.getAttribute("indexed");
    if (ix != null) {
      const i = Number(ix);
      if (i >= 64) return auto;
      if (!css || icvCss(srcPalette, i) === css) return i;
    }
    return css ? palette.index(css, auto) : auto;
  };

  // ---- fonts / formats / XFs -------------------------------------------------------------
  const pool = (name: string): Element[] => {
    const parent = doc ? firstByLocal(doc.documentElement, name) : undefined;
    return parent ? Array.from(parent.children).filter((e) => e.localName === (name === "cellXfs" ? "xf" : name.replace(/s$/, ""))) : [];
  };
  const xFonts = pool("fonts"), xFills = pool("fills"), xBorders = pool("borders"), xXfs = pool("cellXfs");
  const customFmt = new Map<number, string>();
  if (doc) for (const nf of Array.from(doc.getElementsByTagName("numFmt"))) customFmt.set(Number(nf.getAttribute("numFmtId")), nf.getAttribute("formatCode") ?? "General");
  const fontKeyOf = (f: Element | undefined): FontKey => {
    const v = (tag: string): string | null | undefined => (f ? firstByLocal(f, tag)?.getAttribute("val") : undefined);
    const has = (tag: string): boolean => { const e = f && firstByLocal(f, tag); return !!e && e.getAttribute("val") !== "0" && e.getAttribute("val") !== "false"; };
    const u = f && firstByLocal(f, "u");
    const uv = u?.getAttribute("val") ?? "single";
    const va = v("vertAlign");
    return {
      height: Math.round((Number(v("sz")) || 10) * 20), italic: has("i"), strike: has("strike"), weight: has("b") ? 700 : 400,
      underline: !u || uv === "none" ? 0 : uv === "double" ? 2 : uv === "singleAccounting" ? 0x21 : uv === "doubleAccounting" ? 0x22 : 1,
      escapement: va === "superscript" ? 1 : va === "subscript" ? 2 : 0, color: icvOf(f ? firstByLocal(f, "color") : undefined, 0x7fff),
      name: v("name") || v("rFont") || wb.defaultFontName || "Arial",
    };
  };
  const fonts: FontKey[] = [];
  const fontIdx = new Map<string, number>();
  const baseFont = fontKeyOf(xFonts[0]);
  for (let i = 0; i < 4; i++) fonts.push(baseFont);
  fontIdx.set(JSON.stringify(baseFont), 0);
  const ifnt = (k: FontKey): number => {
    const s = JSON.stringify(k);
    let i = fontIdx.get(s);
    if (i == null) {
      fonts.push(k);
      i = fonts.length; // index 4 does not exist: the fifth font is 5
      fontIdx.set(s, i);
    }
    return i;
  };
  const formats = new Map<string, number>();
  const formatRecs: [number, string][] = [];
  let nextFmt = 164;
  for (const [id, code] of customFmt) if (id < 164) { formatRecs.push([id, code]); formats.set(code, id); }
  const ifmtOf = (fmt: string | number | undefined): number => {
    if (fmt == null || fmt === "" || fmt === "General") return 0;
    if (typeof fmt === "number") {
      const code = customFmt.get(fmt);
      return code != null && fmt >= 164 ? ifmtOf(code) : fmt < 164 ? fmt : 0;
    }
    let id = formats.get(fmt);
    if (id == null) { id = nextFmt++; formats.set(fmt, id); formatRecs.push([id, fmt]); }
    return id;
  };
  interface XfKey { font: number; fmt: number; locked: boolean; hidden: boolean; h: number; v: number; wrap: boolean; rot: number; indent: number; shrink: boolean; b: [number, number][]; diag: [number, number, number]; fls: number; fore: number; back: number }
  const xfKeyOf = (xf: Element | undefined, fmtOverride?: string | number): XfKey => {
    const fill = xFills[Number(xf?.getAttribute("fillId") || 0)];
    const pat = fill && firstByLocal(fill, "patternFill");
    const fls = Math.max(0, PATTERNS.indexOf(pat?.getAttribute("patternType") ?? "none"));
    const border = xBorders[Number(xf?.getAttribute("borderId") || 0)];
    const side = (name: string): [number, number] => {
      const s = border && firstByLocal(border, name);
      const st = Math.max(0, BORDER_STYLES.indexOf(s?.getAttribute("style") ?? ""));
      return st ? [st, icvOf(s ? firstByLocal(s, "color") : undefined, 64)] : [0, 0];
    };
    const d = side("diagonal");
    const diagDir = (border?.getAttribute("diagonalDown") === "1" ? 1 : 0) | (border?.getAttribute("diagonalUp") === "1" ? 2 : 0);
    const al = xf && firstByLocal(xf, "alignment");
    const prot = xf && firstByLocal(xf, "protection");
    const numFmtId = Number(xf?.getAttribute("numFmtId") || 0);
    return {
      font: ifnt(fontKeyOf(xFonts[Number(xf?.getAttribute("fontId") || 0)])),
      fmt: ifmtOf(fmtOverride !== undefined ? fmtOverride : numFmtId),
      locked: !(prot?.getAttribute("locked") === "0" || prot?.getAttribute("locked") === "false"),
      hidden: prot?.getAttribute("hidden") === "1" || prot?.getAttribute("hidden") === "true",
      h: Math.max(0, H_ALIGN.indexOf(al?.getAttribute("horizontal") ?? "general")),
      v: (() => { const i = V_ALIGN.indexOf(al?.getAttribute("vertical") ?? "bottom"); return i < 0 ? 2 : i; })(),
      wrap: al?.getAttribute("wrapText") === "1" || al?.getAttribute("wrapText") === "true",
      rot: Math.min(255, Math.max(0, Number(al?.getAttribute("textRotation") || 0))),
      indent: Math.min(15, Number(al?.getAttribute("indent") || 0)),
      shrink: al?.getAttribute("shrinkToFit") === "1",
      b: [side("left"), side("right"), side("top"), side("bottom")],
      diag: diagDir && d[0] ? [diagDir, d[0], d[1]] : [0, 0, 0],
      fls,
      fore: fls ? icvOf(pat ? firstByLocal(pat, "fgColor") : undefined, 64) : 64,
      back: fls ? icvOf(pat ? firstByLocal(pat, "bgColor") : undefined, 65) : 65,
    };
  };
  const xfKeys: XfKey[] = [];
  const xfIdx = new Map<string, number>();
  const ixfe = (k: XfKey): number => {
    const s = JSON.stringify(k);
    let i = xfIdx.get(s);
    if (i == null) {
      if (xfKeys.length >= 4000) return 15; // BIFF8's practical XF limit
      i = 16 + xfKeys.length;
      xfKeys.push(k);
      xfIdx.set(s, i);
    }
    return i;
  };
  // The default cell format (XF 15): the file's own when it came from .xls, else the first xf.
  const defaultKey = xfKeyOf(xXfs[wb.xls && xXfs.length > 15 ? 15 : 0]);
  xfIdx.set(JSON.stringify(defaultKey), 15);
  const xfCache = new Map<string, number>();
  const cellXf = (cell: Cell | undefined, styleIdx?: string): number => {
    const s = styleIdx ?? cell?.style;
    const override = cell?.numFmtDirty ? cell.numFmt ?? "" : undefined;
    if (s == null && override === undefined) return 15;
    const ck = `${s ?? ""}|${override ?? ""}`;
    let i = xfCache.get(ck);
    if (i == null) { i = ixfe(xfKeyOf(xXfs[Number(s ?? 0)], override)); xfCache.set(ck, i); }
    return i;
  };
  const runFont = (run: TextRun, base: FontKey): number => {
    if (!run.own && !run.bold && !run.italic && !run.underline && !run.strike && run.size == null && !run.color && !run.font) return ifnt(base);
    const k: FontKey = run.own
      ? { ...base, weight: run.bold ? 700 : 400, italic: !!run.italic, underline: run.underline ? 1 : 0, strike: !!run.strike }
      : { ...base, weight: run.bold || base.weight >= 700 ? 700 : 400, italic: !!run.italic || base.italic, underline: run.underline ? 1 : base.underline, strike: !!run.strike || base.strike };
    if (run.size != null) k.height = Math.round(run.size * 20);
    if (run.color) k.color = palette.index(run.color, 0x7fff);
    if (run.font) k.name = run.font;
    return ifnt(k);
  };

  // ---- strings ----------------------------------------------------------------------------
  const sst: { text: string; runs?: [number, number][] }[] = [];
  const sstIdx = new Map<string, number>();
  let sstTotal = 0;
  const isst = (text: string, runs?: [number, number][]): number => {
    sstTotal++;
    const k = runs ? text + "\u0000" + JSON.stringify(runs) : text;
    let i = sstIdx.get(k);
    if (i == null) { i = sst.length; sst.push({ text, runs }); sstIdx.set(k, i); }
    return i;
  };

  // ---- names, external sheets, add-in functions ---------------------------------------------
  const sheets = wb.sheets.slice(0, 4000);
  const sheetIdx = new Map(sheets.map((s, i) => [s.name.toLowerCase(), i]));
  const externals = (wb.xls?.externals ?? []).map((e) => ({ ...e, names: [...e.names] }));
  interface NameRec { name: string; itab: number; hidden: boolean; builtin?: number; func?: boolean; grbit?: number; text?: string; rgce?: Uint8Array; rgcb?: Uint8Array }
  const names: NameRec[] = [];
  const quote = (n: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
  const abs = (g: { r1: number; c1: number; r2: number; c2: number }): string => `$${colToLetters(g.c1)}$${g.r1}:$${colToLetters(g.c2)}$${g.r2}`;
  // Names Excel keeps as built-ins (stored as one character code), other than the print ones.
  const BUILTIN: Record<string, number> = { consolidate_area: 0, auto_open: 1, auto_close: 2, extract: 3, database: 4, criteria: 5, recorder: 8, data_form: 9, auto_activate: 10, auto_deactivate: 11, sheet_title: 12 };
  const named = (name: string, itab: number, text: string): NameRec => {
    const b = BUILTIN[name.toLowerCase()];
    return b != null ? { name, builtin: b, itab, hidden: false, text } : { name, itab, hidden: false, text };
  };
  for (const [name, ref] of wb.definedNames ?? []) names.push(named(name, 0, ref));
  for (const ln of wb.localNames ?? []) { const i = sheetIdx.get(ln.sheet.toLowerCase()); if (i != null) names.push(named(ln.name, i + 1, ln.ref)); }
  sheets.forEach((s, i) => {
    const q = quote(s.name);
    const p = s.printSetup;
    if (p?.printArea?.length) names.push({ name: "Print_Area", builtin: 6, itab: i + 1, hidden: false, text: p.printArea.map((g) => `${q}!${abs({ r1: Math.min(g.r1, MAX_ROWS), c1: Math.min(g.c1, MAX_COLS), r2: Math.min(g.r2, MAX_ROWS), c2: Math.min(g.c2, MAX_COLS) })}`).join(",") });
    const titles: string[] = [];
    if (p?.titleCols) titles.push(`${q}!$${colToLetters(p.titleCols.from)}:$${colToLetters(p.titleCols.to)}`);
    if (p?.titleRows) titles.push(`${q}!$${p.titleRows.from}:$${p.titleRows.to}`);
    if (titles.length) names.push({ name: "Print_Titles", builtin: 7, itab: i + 1, hidden: false, text: titles.join(",") });
    if (s.autoFilter) names.push({ name: "_FilterDatabase", builtin: 13, itab: i + 1, hidden: true, text: `${q}!${abs(s.autoFilter)}` });
  });
  const udfs = new Map((wb.xls?.udfs ?? []).map((n) => [n.toUpperCase(), n]));
  const xtis: [number, number, number][] = [];
  const xtiIdx = new Map<string, number>();
  const addinNames: string[] = [];
  const addinBook = 1 + externals.length;
  const xtiOf = (book: number, first: number, last: number): number => {
    const k = `${book}:${first}:${last}`;
    let i = xtiIdx.get(k);
    if (i == null) { i = xtis.length; xtis.push([book, first, last]); xtiIdx.set(k, i); }
    return i;
  };
  const envFor = (sheet: number, force3d = false): CompileEnv => ({
    sheet,
    force3d,
    sheetIndex: (n) => sheetIdx.get(n.toLowerCase()) ?? -1,
    xti: (first, last, ext) => {
      if (!ext) return xtiOf(0, first, last);
      const book = externals[ext.book - 1];
      if (!book) throw new CompileError("unknown external book");
      const f = ext.first ? book.sheets.findIndex((s) => s.toLowerCase() === ext.first.toLowerCase()) : -2;
      const l = ext.last ? book.sheets.findIndex((s) => s.toLowerCase() === ext.last.toLowerCase()) : f;
      return xtiOf(ext.book, f, l);
    },
    nameIndex: (n) => {
      const lower = n.toLowerCase();
      let i = sheet >= 0 ? names.findIndex((x) => x.itab === sheet + 1 && x.builtin == null && x.name.toLowerCase() === lower) : -1;
      if (i < 0) i = names.findIndex((x) => x.itab === 0 && x.name.toLowerCase() === lower);
      if (i < 0) { names.push({ name: n, itab: 0, hidden: false }); i = names.length - 1; }
      return i + 1;
    },
    addinFunc: (n) => {
      let i = addinNames.indexOf(n);
      if (i < 0) { addinNames.push(n); i = addinNames.length - 1; }
      return { ixti: xtiOf(addinBook, -2, -2), index: i + 1 };
    },
    externName: (book, n) => {
      const ext = externals[book - 1];
      if (!ext) throw new CompileError("unknown external book");
      let i = ext.names.findIndex((x) => x.toLowerCase() === n.toLowerCase());
      if (i < 0) { ext.names.push(n); i = ext.names.length - 1; }
      return { ixti: xtiOf(book, -2, -2), index: i + 1 };
    },
    sheetName: (sheetName, n) => {
      const si = sheetIdx.get(sheetName.toLowerCase());
      if (si == null) throw new CompileError("unknown sheet");
      const lower = n.toLowerCase();
      let i = names.findIndex((x) => x.itab === si + 1 && x.builtin == null && x.name.toLowerCase() === lower);
      if (i < 0) i = names.findIndex((x) => x.itab === 0 && x.name.toLowerCase() === lower);
      if (i < 0) throw new CompileError("unknown name");
      return { ixti: xtiOf(0, si, si), index: i + 1 };
    },
    funcName: (n) => {
      const real = udfs.get(n.toUpperCase()) ?? "_xlfn." + n.toUpperCase();
      let i = names.findIndex((x) => x.func && x.name === real);
      // Future functions are hidden function names (0x0B, as Excel writes them); macros are VBA procedures.
      if (i < 0) { names.push({ name: real, itab: 0, hidden: real.startsWith("_xlfn."), func: true, grbit: real.startsWith("_xlfn.") ? 0x000b : 0x000e }); i = names.length - 1; }
      return i + 1;
    },
  });
  const compile = (text: string, env: CompileEnv, top: Ctx = "V"): { rgce: Uint8Array; rgcb: Uint8Array } | null => {
    try {
      const r = compileFormula(text.replace(/^=/, ""), env, top);
      return r.rgce.length > 1800 || r.rgcb.length > 6000 ? null : r;
    } catch {
      return null;
    }
  };
  for (let i = 0; i < names.length; i++) {
    const n = names[i]!;
    if (n.text == null) continue;
    const r = compile(n.text, envFor(n.itab ? n.itab - 1 : -1, true), "R");
    if (r) { n.rgce = r.rgce; n.rgcb = r.rgcb; }
  }

  // ---- sheets -----------------------------------------------------------------------------
  const blips: { type: number; data: Uint8Array }[] = [];
  const drawings: { dg: number; shapes: number }[] = [];
  // ActiveX controls' persisted data, one slice each.
  const ctlsStream = new Bytes();
  // A macro a control runs: a hidden-less VBA procedure name (fOB | fProc), as Excel writes it.
  const macroName = (n: string): number => {
    let i = names.findIndex((x) => x.itab === 0 && x.builtin == null && x.name.toLowerCase() === n.toLowerCase());
    if (i < 0) { names.push({ name: n, itab: 0, hidden: false, grbit: 0x000c }); i = names.length - 1; }
    return i + 1;
  };
  const sheetBodies = sheets.map((s, si) => writeSheet(s, si));

  function writeSheet(s: Sheet, si: number): { rs: RecStream; indexAt: number; dbcells: number[]; defColAt: number } {
    const rs = new RecStream();
    const sheetChart = s.chartSheet ? s.charts?.[0] : undefined;
    if (sheetChart) {
      const a = sheetChart.anchor;
      const w = Math.max(100, a.toColOff - a.fromColOff), h = Math.max(100, a.toRowOff - a.fromRowOff);
      rs.bytes(writeChartStream(sheetChart, chartEnv(si), { w: w * 0.75, h: h * 0.75 }, { sheet: true, codeName: s.codeName }));
      return { rs, indexAt: -1, dbcells: [], defColAt: 0 };
    }
    const env = envFor(si);
    rs.rec(0x0809, body().u16(0x0600).u16(0x0010).u16(0x0dbb).u16(0x07cc).u32(0x000100c1).u32(0x00000006));
    // Cells, grouped by row.
    const rows = new Map<number, Cell[]>();
    let maxR = -1, maxC = -1, minR = MAX_ROWS, minC = MAX_COLS;
    const arrayMembers = new Map<string, { r: number; c: number }>();
    for (const cell of s.cells.values()) {
      if (cell.arrayRef && cell.formula != null) {
        const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(cell.arrayRef);
        if (m) {
          const r2 = Number(m[4] ?? m[2]), c2 = m[3] ? lettersCol(m[3]) : lettersCol(m[1]!);
          for (let r = cell.row; r <= Math.min(r2, MAX_ROWS); r++) for (let c = cell.col; c <= Math.min(c2, MAX_COLS); c++) arrayMembers.set(`${r}:${c}`, { r: cell.row, c: cell.col });
        }
      }
    }
    for (const cell of s.cells.values()) {
      if (cell.row > MAX_ROWS || cell.col > MAX_COLS) { report.truncatedCells++; continue; }
      const hasContent = cell.kind !== "blank" && cell.value !== "" || cell.formula != null || cell.style != null || arrayMembers.has(`${cell.row}:${cell.col}`);
      if (!hasContent) continue;
      let list = rows.get(cell.row - 1);
      if (!list) rows.set(cell.row - 1, (list = []));
      list.push(cell);
      maxR = Math.max(maxR, cell.row - 1); maxC = Math.max(maxC, cell.col - 1);
      minR = Math.min(minR, cell.row - 1); minC = Math.min(minC, cell.col - 1);
    }
    const defRowTw = s.defaultRowHeight ? Math.round(s.defaultRowHeight * 15) : 255;
    const rowNums = new Set<number>(rows.keys());
    for (const r of s.rowHeights?.keys() ?? []) if (r <= MAX_ROWS) rowNums.add(r - 1);
    for (const r of s.hiddenRows ?? []) if (r <= MAX_ROWS) rowNums.add(r - 1);
    for (const r of s.rowOutline?.keys() ?? []) if (r <= MAX_ROWS) rowNums.add(r - 1);
    const sortedRows = [...rowNums].sort((a, b) => a - b);
    const blocks = new Map<number, number[]>();
    for (const r of sortedRows) { const b = r >> 5; let l = blocks.get(b); if (!l) blocks.set(b, (l = [])); l.push(r); }

    // INDEX, sized now and filled in once the sheet's position in the stream is known.
    const indexAt = rs.len + 4;
    rs.rec(0x020b, body().u32(0).u32(maxR < 0 ? 0 : minR).u32(maxR + 1).u32(0).zeros(4 * blocks.size));
    rs.rec(0x000d, body().u16(1)); // CALCMODE
    rs.rec(0x000c, body().u16(100));
    rs.rec(0x000f, body().u16(1));
    rs.rec(0x0011, body().u16(0));
    rs.rec(0x0010, body().f64(0.001));
    rs.rec(0x005f, body().u16(1));
    const p = s.printSetup ?? {};
    rs.rec(0x002a, body().u16(p.headings ? 1 : 0));
    rs.rec(0x002b, body().u16(p.gridLines ? 1 : 0));
    rs.rec(0x0082, body().u16(1));
    const rowLevels = Math.max(0, ...(s.rowOutline?.values() ?? []));
    const colLevels = Math.max(0, ...(s.colOutline?.values() ?? []));
    rs.rec(0x0080, body().u16(rowLevels ? 12 * rowLevels + 17 : 0).u16(colLevels ? 12 * colLevels + 17 : 0).u16(rowLevels ? rowLevels + 1 : 0).u16(colLevels ? colLevels + 1 : 0));
    rs.rec(0x0225, body().u16(defRowTw !== 255 ? 1 : 0).u16(defRowTw));
    rs.rec(0x0081, body().u16(0x0401 | (s.summaryBelow === false ? 0 : 0x40) | (s.summaryRight === false ? 0 : 0x80) | (p.fitToPage ? 0x100 : 0) | (rowLevels || colLevels ? 0x0c00 : 0)));
    const breaks = (type: number, list: number[] | undefined, max: number): void => {
      const l = (list ?? []).filter((b) => b > 1 && b <= max).slice(0, 1026);
      if (!l.length) return;
      const b = body().u16(l.length);
      for (const x of l) b.u16(x - 1).u16(0).u16(type === 0x001b ? 255 : 65535);
      rs.rec(type, b);
    };
    breaks(0x001b, p.rowBreaks, MAX_ROWS);
    breaks(0x001a, p.colBreaks, MAX_COLS);
    const hf = (h?: { left?: string; center?: string; right?: string }): string => (h ? (h.left ? "&L" + h.left : "") + (h.center ? "&C" + h.center : "") + (h.right ? "&R" + h.right : "") : "");
    const head = hf(p.header).slice(0, 255), foot = hf(p.footer).slice(0, 255);
    rs.rec(0x0014, head ? body().str16(head) : undefined);
    rs.rec(0x0015, foot ? body().str16(foot) : undefined);
    rs.rec(0x0083, body().u16(p.horizontalCentered ? 1 : 0));
    rs.rec(0x0084, body().u16(p.verticalCentered ? 1 : 0));
    const m = p.margins;
    if (m) { rs.rec(0x0026, body().f64(m.left)); rs.rec(0x0027, body().f64(m.right)); rs.rec(0x0028, body().f64(m.top)); rs.rec(0x0029, body().f64(m.bottom)); }
    const known = p.paperSize != null || p.orientation != null || p.scale != null;
    const setup = (p.pageOrder === "overThenDown" ? 1 : 0) | (p.orientation === "landscape" ? 0 : 2) | (known ? 0 : 4) | (p.orientation ? 0 : 0x40) | (p.firstPageNumber != null ? 0x80 : 0);
    rs.rec(0x00a1, body().u16(p.paperSize ?? 9).u16(p.scale ?? 100).u16(p.firstPageNumber ?? 1).u16(p.fitToWidth ?? 1).u16(p.fitToHeight ?? 1).u16(setup).u16(600).u16(600).f64(m?.header ?? 0.5).f64(m?.footer ?? 0.5).u16(1));
    const prot = s.protection;
    if (prot?.sheet) {
      const locked = (f: SheetLock): boolean => prot.locks?.[f] ?? SHEET_LOCK_DEFAULTS[f];
      rs.rec(0x0012, body().u16(1));
      if (locked("scenarios")) rs.rec(0x00dd, body().u16(1));
      if (locked("objects")) rs.rec(0x0063, body().u16(1));
      const pw = prot.password?.legacy ? parseInt(prot.password.legacy, 16) : 0;
      if (pw) rs.rec(0x0013, body().u16(pw));
    }
    // Columns.
    const mdw = s.maxDigitWidth ?? maxDigitWidth(undefined);
    const defW = s.defaultColWidth ? Math.max(1, Math.round(colWidthChars(s.defaultColWidth, mdw) - 0.71)) : 8;
    const defColAt = rs.len;
    rs.rec(0x0055, body().u16(defW));
    const colInfo = new Map<number, { w?: number; hidden: boolean; level: number; collapsed: boolean }>();
    const touch = (c: number) => { let v = colInfo.get(c); if (!v) colInfo.set(c, (v = { hidden: false, level: 0, collapsed: false })); return v; };
    for (const [c, px] of s.colWidths ?? []) if (c <= MAX_COLS) touch(c).w = Math.round(colWidthChars(px, mdw) * 256);
    for (const c of s.hiddenCols ?? []) if (c <= MAX_COLS) touch(c).hidden = true;
    for (const [c, l] of s.colOutline ?? []) if (c <= MAX_COLS) touch(c).level = Math.min(7, l);
    for (const c of s.colCollapsed ?? []) if (c <= MAX_COLS) touch(c).collapsed = true;
    const cols = [...colInfo.keys()].sort((a, b) => a - b);
    for (let i = 0; i < cols.length;) {
      const c = cols[i]!, info = colInfo.get(c)!;
      const sig = JSON.stringify(info);
      let j = i;
      while (j + 1 < cols.length && cols[j + 1] === cols[j]! + 1 && JSON.stringify(colInfo.get(cols[j + 1]!)) === sig) j++;
      const w = info.w ?? Math.round(colWidthChars(s.defaultColWidth ?? colWidthToPx(defW + 0.71, mdw), mdw) * 256);
      rs.rec(0x007d, body().u16(c - 1).u16(cols[j]! - 1).u16(Math.min(0xffff, w)).u16(15).u16((info.hidden ? 1 : 0) | (info.w != null ? 2 : 0) | (info.level << 8) | (info.collapsed ? 0x1000 : 0)).u16(0));
      i = j + 1;
    }
    rs.rec(0x0200, body().u32(maxR < 0 ? 0 : minR).u32(maxR + 1).u16(maxR < 0 ? 0 : minC).u16(maxC + 1).u16(0));

    // Row blocks: up to 32 ROW records, their cells, then a DBCELL.
    const dbcells: number[] = [];
    for (const blockRows of blocks.values()) {
      const blockStart = rs.len;
      for (const r of blockRows) {
        const list = rows.get(r) ?? [];
        const first = list.length ? Math.min(...list.map((c) => c.col - 1)) : 0;
        const last = list.length ? Math.max(...list.map((c) => c.col)) : 0;
        const ht = s.rowHeights?.get(r + 1);
        const tw = ht != null ? Math.min(0x7fff, Math.round(ht * 15)) : defRowTw;
        const level = Math.min(7, s.rowOutline?.get(r + 1) ?? 0);
        const flags = level | (s.rowCollapsed?.has(r + 1) ? 0x10 : 0) | (s.hiddenRows?.has(r + 1) ? 0x20 : 0) | (ht != null ? 0x40 : 0) | 0x100;
        rs.rec(0x0208, body().u16(r).u16(first).u16(last).u16(tw).u16(0).u16(0).u16(flags).u16(15));
      }
      const rowEnd = rs.len;
      const offsets: number[] = [];
      let prev = rowEnd - 20;
      for (const r of blockRows) {
        const at = rs.len;
        offsets.push(at - prev);
        prev = at;
        for (const cell of (rows.get(r) ?? []).sort((a, b) => a.col - b.col)) writeCell(rs, cell, env, arrayMembers);
      }
      const dbAt = rs.len;
      const db = body().u32(dbAt - blockStart);
      offsets.forEach((o) => db.u16(Math.min(0xffff, o)));
      rs.rec(0x00d7, db);
      dbcells.push(dbAt);
    }
    writeObjects(rs, s, si);
    // Window.
    const fz = s.freeze && (s.freeze.rows || s.freeze.cols) ? s.freeze : undefined;
    const fr = fz ? Math.min(fz.rows, MAX_ROWS - 1) : 0, fc = fz ? Math.min(fz.cols, MAX_COLS - 1) : 0;
    const active = si === Math.min(wb.activeSheet ?? 0, sheets.length - 1);
    const split = !!fz && !!s.paneSplit;
    const w2 = (s.hideGridLines ? 0 : 0x2) | 0x4 | 0x10 | 0x20 | 0x80 | (fz && !split ? 0x108 : 0) | (active ? 0x600 : 0);
    rs.rec(0x023e, body().u16(w2).u16(0).u16(0).u16(64).u16(0).u16(0).u16(0).u32(0));
    let pnn = 3;
    if (fz) {
      pnn = fr && fc ? 0 : fr ? 2 : 1;
      let x = fc, y = fr;
      if (split) {
        x = 0; y = 0;
        for (let c = 1; c <= fc; c++) x += (s.colWidths?.get(c) ?? s.defaultColWidth ?? 64) * 15;
        for (let r = 1; r <= fr; r++) y += (s.rowHeights?.get(r) ?? s.defaultRowHeight ?? 20) * 15;
        x = Math.min(0xffff, Math.round(x)); y = Math.min(0xffff, Math.round(y));
      }
      rs.rec(0x0041, body().u16(x).u16(y).u16(fr).u16(fc).u8(pnn).u8(0));
    }
    rs.rec(0x001d, body().u8(pnn).u16(fr).u16(fc).u16(0).u16(1).u16(fr).u16(fr).u8(fc).u8(fc));
    // Merged cells, 1026 per record.
    const merges = (s.merges ?? []).filter((g) => g.r1 <= MAX_ROWS && g.c1 <= MAX_COLS);
    for (let i = 0; i < merges.length; i += 1026) {
      const chunk = merges.slice(i, i + 1026);
      const b = body().u16(chunk.length);
      for (const g of chunk) b.u16(g.r1 - 1).u16(Math.min(g.r2, MAX_ROWS) - 1).u16(g.c1 - 1).u16(Math.min(g.c2, MAX_COLS) - 1);
      rs.rec(0x00e5, b);
    }
    writeCondFormats(rs, s, env);
    writeLinks(rs, s);
    writeValidations(rs, s, env);
    if (s.codeName) rs.rec(0x01ba, body().str16(s.codeName.slice(0, 31)));
    if (s.tabColor) {
      const icv = palette.index(s.tabColor, 0x7f);
      rs.rec(0x0862, body().u16(0x0862).u16(0).zeros(8).u32(0x14).u32(icv & 0x7f));
    }
    if (prot?.sheet) {
      let allow = 0;
      LOCK_ORDER.forEach((f, i) => { if (!(prot.locks?.[f] ?? SHEET_LOCK_DEFAULTS[f])) allow |= 1 << i; });
      rs.rec(0x0867, body().u16(0x0867).u16(0).zeros(8).u16(2).u8(1).u32(4).u32(allow));
    }
    rs.rec(0x000a);
    return { rs, indexAt, dbcells, defColAt };
  }

  function writeCell(rs: RecStream, cell: Cell, env: CompileEnv, arrayMembers: Map<string, { r: number; c: number }>): void {
    const r = cell.row - 1, c = cell.col - 1;
    const xf = cellXf(cell);
    const head = (b: Bytes): Bytes => b.u16(r).u16(c).u16(xf);
    const member = arrayMembers.get(`${cell.row}:${cell.col}`);
    let formula: { rgce: Uint8Array; rgcb: Uint8Array } | null = null;
    let isAnchor = false;
    if (member) {
      formula = { rgce: new Uint8Array([0x01, (member.r - 1) & 255, (member.r - 1) >> 8, (member.c - 1) & 255, (member.c - 1) >> 8]), rgcb: new Uint8Array(0) };
      isAnchor = member.r === cell.row && member.c === cell.col;
    } else if (cell.formula != null) {
      formula = compile(cell.formula, env);
      if (!formula) report.droppedFormulas++;
    }
    const text = cell.kind === "s" ? truncate(cell.value) : "";
    if (formula) {
      const b = head(body());
      let str: string | null = null;
      if (cell.kind === "n" && Number.isFinite(Number(cell.value)) && cell.value !== "") b.f64(Number(cell.value));
      else if (cell.kind === "b") b.u8(1).u8(0).u8(cell.value === "TRUE" || cell.value === "1" ? 1 : 0).u8(0).u8(0).u8(0).u16(0xffff);
      else if (cell.kind === "e") b.u8(2).u8(0).u8(ERR_CODE[cell.value] ?? 0x2a).u8(0).u8(0).u8(0).u16(0xffff);
      else if (cell.kind === "s" && text !== "") { b.u8(0).zeros(5).u16(0xffff); str = text; }
      else if (cell.kind === "s" || cell.kind === "blank") b.u8(3).zeros(5).u16(0xffff);
      else b.f64(0);
      b.u16(0x0002).u32(0).u16(formula.rgce.length).bytes(formula.rgce).bytes(formula.rgcb);
      rs.rec(0x0006, b);
      if (isAnchor && cell.arrayRef && cell.formula != null) {
        const arr = compile(cell.formula, env, "A");
        const mm = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(cell.arrayRef)!;
        const r2 = Math.min(Number(mm[4] ?? mm[2]), MAX_ROWS) - 1, c2 = Math.min(mm[3] ? lettersCol(mm[3]) : lettersCol(mm[1]!), MAX_COLS) - 1;
        const a = body().u16(r).u16(r2).u8(c).u8(c2).u16(0).u32(0);
        if (arr) a.u16(arr.rgce.length).bytes(arr.rgce).bytes(arr.rgcb);
        else a.u16(2).u8(0x1c).u8(0x2a);
        rs.rec(0x0221, a);
      }
      if (str != null) {
        // A long result continues in CONTINUE records, each restarting with its own flag byte.
        const wide = /[^\u0000-ÿ]/.test(str);
        const per = wide ? Math.floor((MAX_REC - 4) / 2) : MAX_REC - 4;
        rs.rec(0x0207, body().u16(str.length).chars(str.slice(0, per)));
        const more = wide ? Math.floor((MAX_REC - 1) / 2) : MAX_REC - 1;
        for (let o = per; o < str.length; o += more) {
          const part = str.slice(o, o + more);
          const b = body().u8(wide ? 1 : 0);
          for (let i = 0; i < part.length; i++) { const code = part.charCodeAt(i); if (wide) b.u16(code); else b.u8(code); }
          rs.rec(0x003c, b);
        }
      }
      return;
    }
    switch (cell.kind) {
      case "n": {
        const v = Number(cell.value);
        if (!Number.isFinite(v)) { rs.rec(0x0201, head(body())); return; }
        const rkv = toRk(v);
        if (rkv != null) rs.rec(0x027e, head(body()).u32(rkv));
        else rs.rec(0x0203, head(body()).f64(v));
        return;
      }
      case "s": {
        if (text === "") { rs.rec(0x0201, head(body())); return; }
        let runs: [number, number][] | undefined;
        if (cell.richRuns?.length) {
          const base = fonts[fontPos((xf === 15 ? defaultKey : xfKeys[xf - 16])?.font ?? 0)] ?? baseFont;
          runs = [];
          let at = 0;
          for (const run of cell.richRuns) {
            if (at < text.length && run.text) runs.push([at, runFont(run, base)]);
            at += run.text.length;
          }
          if (!runs.length) runs = undefined;
        }
        rs.rec(0x00fd, head(body()).u32(isst(text, runs)));
        return;
      }
      case "b": rs.rec(0x0205, head(body()).u8(cell.value === "TRUE" || cell.value === "1" ? 1 : 0).u8(0)); return;
      case "e": rs.rec(0x0205, head(body()).u8(ERR_CODE[cell.value] ?? 0x2a).u8(1)); return;
      default: rs.rec(0x0201, head(body()));
    }
  }

  function truncate(s: string): string {
    if (s.length <= 32767) return s;
    report.truncatedCells++;
    return s.slice(0, 32767);
  }

  // ---- drawing objects: pictures, charts, shapes and comments --------------------------------
  function writeObjects(rs: RecStream, s: Sheet, si: number): void {
    const notes: { r: number; c: number; author: string; text: string }[] = [];
    for (const cell of s.cells.values()) {
      if (!cell.comments?.length || cell.row > MAX_ROWS || cell.col > MAX_COLS) continue;
      notes.push({ r: cell.row - 1, c: cell.col - 1, author: cell.comments[0]!.author ?? "", text: cell.comments.map((x) => x.text).join("\n") });
    }
    const pics: { anchor: NonNullable<Sheet["images"]>[number]["anchor"]; blip: number }[] = [];
    for (const im of s.images ?? []) {
      const b = blipOf(im.replaceBytes ? bytesUri(im.replaceBytes, im.replaceExt ?? "png") : im.dataUri);
      if (b == null) { report.droppedImages++; continue; }
      pics.push({ anchor: im.anchor, blip: b });
    }
    const charts = s.chartSheet ? [] : s.charts ?? [];
    const shapes = (s.shapes ?? []).slice(0, 1000);
    // ActiveX controls go back only with the persisted data they were read with.
    const controls = (s.controls ?? []).filter((c) => !c.activeX || (c.xls?.progId && c.activeXBinPath && wb.files[c.activeXBinPath])).slice(0, 1000);
    if (!notes.length && !pics.length && !charts.length && !shapes.length && !controls.length) return;
    const dg = drawings.length + 1;
    const spid = (k: number): number => dg * 1024 + k;
    // A shape container with its OBJ, and its TXO or chart substream; groups hold their members.
    type Txo = { text: string; runs: [number, number][]; grbit: number; rot: number };
    type Item = { sp: Uint8Array; obj: Bytes; txo?: Txo; chart?: Uint8Array; kids?: Item[]; ctl?: boolean };
    const items: Item[] = [];
    const colPx = (c: number): number => (s.hiddenCols?.has(c) ? 0 : s.colWidths?.get(c) ?? s.defaultColWidth ?? 64);
    const rowPx = (r: number): number => (s.hiddenRows?.has(r) ? 0 : s.rowHeights?.get(r) ?? s.defaultRowHeight ?? 17);
    const anchorBytes = (a: { fromCol: number; fromRow: number; fromColOff: number; fromRowOff: number; toCol: number; toRow: number; toColOff: number; toRowOff: number }): Bytes => {
      // Offsets in px may run past their cell; walk them into the right column / row first.
      const walk = (line: number, off: number, size: (n: number) => number, max: number): [number, number] => {
        let n = line, o = off;
        while (n < max && size(n) > 0 && o >= size(n)) { o -= size(n); n++; }
        const sz = size(n) || 1;
        return [Math.min(n, max) - 1, Math.max(0, Math.min(1, o / sz))];
      };
      const [c1, x1] = walk(a.fromCol, a.fromColOff, colPx, MAX_COLS), [r1, y1] = walk(a.fromRow, a.fromRowOff, rowPx, MAX_ROWS);
      const [c2, x2] = walk(a.toCol, a.toColOff, colPx, MAX_COLS), [r2, y2] = walk(a.toRow, a.toRowOff, rowPx, MAX_ROWS);
      return body().u16(0).u16(c1).u16(Math.round(x1 * 1023)).u16(r1).u16(Math.round(y1 * 255)).u16(c2).u16(Math.round(x2 * 1023)).u16(r2).u16(Math.round(y2 * 255));
    };
    const cmo = (ot: number, id: number, flags = 0x6011): Bytes => body().u16(0x15).u16(0x12).u16(ot).u16(id).u16(flags).zeros(12);
    let k = 1;
    // Controls keep their ids where they can: Excel names an unnamed control after its id.
    const reserved = new Set<number>();
    const ctlIds = controls.map((c) => { const id = c.xls?.id; if (id && id < 0xffff && !reserved.has(id)) { reserved.add(id); return id; } return 0; });
    let objId = 1;
    const nextId = (): number => { while (reserved.has(objId)) objId++; return objId++; };
    for (const pic of pics) {
      const sp = new Bytes();
      escher(sp, 0xf00a, 0x2, 75, body().u32(spid(k++)).u32(0xa00));
      escher(sp, 0xf00b, 0x3, 2, body().u16(0x4104).u32(pic.blip).u16(0x03bf).u32(0x00080000));
      escher(sp, 0xf010, 0, 0, anchorBytes(pic.anchor));
      escher(sp, 0xf011, 0, 0, body());
      const obj = cmo(8, nextId()).u16(0x07).u16(2).u16(0xffff).u16(0x08).u16(2).u16(1);
      items.push({ sp: container(0xf004, sp).out().slice(), obj });
    }
    for (const ch of charts) {
      const sp = new Bytes();
      escher(sp, 0xf00a, 0x2, 201, body().u32(spid(k++)).u32(0xa00));
      const props: [number, number][] = [[0x007f, 0x01040104], [0x00bf, 0x00080008], [0x0181, 0x0800004e], [0x0183, 0x0800004d], [0x01bf, 0x00110010], [0x01c0, 0x0800004d], [0x01ff, 0x00080008], [0x023f, 0x00020000], [0x03bf, 0x00080000]];
      const o = body();
      for (const [pid, v] of props) o.u16(pid).u32(v);
      escher(sp, 0xf00b, 0x3, props.length, o);
      escher(sp, 0xf010, 0, 0, anchorBytes(ch.anchor));
      escher(sp, 0xf011, 0, 0, body());
      const a = ch.anchor;
      let wPx = a.toColOff - a.fromColOff, hPx = a.toRowOff - a.fromRowOff;
      for (let c = a.fromCol; c < a.toCol; c++) wPx += colPx(c);
      for (let r = a.fromRow; r < a.toRow; r++) hPx += rowPx(r);
      const stream = writeChartStream(ch, chartEnv(si), { w: Math.max(10, wPx * 0.75), h: Math.max(10, hPx * 0.75) });
      items.push({ sp: container(0xf004, sp).out().slice(), obj: cmo(5, nextId()), chart: stream });
    }
    // Shapes: untouched ones go back as read, groups with them; the others are built from the model.
    const groupItems = new Map<number, Item>();
    for (const sh of shapes) {
      const x = sh.xls;
      const txo = shapeTxo(sh);
      const id = nextId();
      if (x?.groups?.length) {
        let parent: Item[] = items;
        for (const g of x.groups) {
          let gi = groupItems.get(g.id);
          if (!gi) {
            let raw = withSpid(g.raw, spid(k++));
            if (g === x.groups[0] && sh.dirty && moved(sh)) raw = withAnchor(raw, anchorBytes(sh.anchor).out());
            gi = { sp: raw, obj: cmo(0, nextId()).u16(0x06).u16(2).u16(0), kids: [] };
            groupItems.set(g.id, gi);
            parent.push(gi);
          }
          parent = gi.kids!;
        }
        parent.push({ sp: stripTextbox(withSpid(x.raw, spid(k++))), obj: cmo(x.ot, id), txo: hasTextbox(x.raw) ? txo ?? EMPTY_TXO : undefined });
        continue;
      }
      const usable = x && !sh.created && !sh.styleDirty && !sh.xfrmDirty && geomOf(sh.preset ?? null) === sh.geom && (!sh.text || hasTextbox(x.raw));
      if (usable) {
        let raw = withSpid(x.raw, spid(k++));
        if (sh.dirty && moved(sh)) raw = withAnchor(raw, anchorBytes(sh.anchor).out());
        items.push({ sp: stripTextbox(raw), obj: cmo(x.ot, id), txo: hasTextbox(x.raw) ? txo ?? EMPTY_TXO : undefined });
        continue;
      }
      const spt = sptOf(sh.preset, sh.geom) ?? 1;
      const isLine = sh.geom === "line";
      const sp = new Bytes();
      escher(sp, 0xf00a, 0x2, spt, body().u32(spid(k++)).u32(0xa00 | (sh.flipH ? 0x40 : 0) | (sh.flipV ? 0x80 : 0)));
      const props: [number, number][] = [];
      if (sh.rotation) props.push([0x0004, Math.round(sh.rotation * 65536) | 0]);
      if (sh.text) props.push([0x0080, id * 0x10000], [0x00bf, 0x00080008]);
      if (!isLine && sh.fill) {
        const grad = sh.fillGradient;
        if (grad && grad.stops.length >= 2) props.push([0x0180, 7]);
        props.push([0x0181, colorRgb(sh.fill)]);
        if (sh.fillOpacity != null && sh.fillOpacity < 1) props.push([0x0182, Math.round(sh.fillOpacity * 65536)]);
        if (grad && grad.stops.length >= 2) props.push([0x0183, colorRgb(grad.stops[grad.stops.length - 1]!.color)], [0x018b, (Math.round(270 - grad.angle) % 360) * 65536]);
      }
      props.push([0x01bf, !isLine && sh.fill ? 0x00100010 : 0x00100000]);
      if (sh.stroke) {
        props.push([0x01c0, colorRgb(sh.stroke)], [0x01cb, Math.round((sh.strokeWidth ?? 1) * 9525)]);
        const dash = DASH_CODE[sh.dash ?? ""];
        if (dash) props.push([0x01ce, dash]);
        const head = END_CODE[sh.headEnd ?? ""], tail = END_CODE[sh.tailEnd ?? ""];
        if (head) props.push([0x01d0, head]);
        if (tail) props.push([0x01d1, tail]);
      }
      props.push([0x01ff, sh.stroke ? 0x00080008 : 0x00080000], [0x03bf, 0x00020002]);
      const o = body();
      for (const [pid, v] of props.sort((a, b) => a[0] - b[0])) o.u16(pid).u32(v >>> 0);
      escher(sp, 0xf00b, 0x3, props.length, o);
      escher(sp, 0xf010, 0, 0, anchorBytes(sh.anchor));
      escher(sp, 0xf011, 0, 0, body());
      if (sh.text) escher(sp, 0xf00d, 0, 0, body());
      const ot = isLine ? 0x1e : spt === 202 ? 6 : 0x1e;
      items.push({ sp: container(0xf004, sp).out().slice(0, sh.text ? -8 : undefined), obj: cmo(ot, id), txo: sh.text ? txo : undefined });
    }
    // Controls: ids first, since option buttons chain to each other by id.
    ctlIds.forEach((id, i) => { if (!id) ctlIds[i] = nextId(); });
    const byOldId = new Map<number, number>();
    controls.forEach((c, i) => { if (c.xls) byOldId.set(c.xls.id, ctlIds[i]!); });
    const cenv: CtlWriteEnv = {
      ref: (text) => compile(text, envFor(si), "R")?.rgce ?? null,
      macro: (text) => macroRgce(text),
      newId: (old) => byOldId.get(old),
    };
    controls.forEach((ctl, i) => {
      const id = ctlIds[i]!, x = ctl.xls;
      const anchor = ctl.anchor ? anchorBytes(ctl.anchor).out() : undefined;
      if (ctl.activeX) {
        let raw = withSpid(x!.raw, spid(k++));
        if (x!.grouped) raw = ungroupedShape(raw, anchor && ctlMoved(ctl) ? anchor : x!.grouped);
        else if (anchor && ctlMoved(ctl)) raw = withAnchor(raw, anchor);
        const pib = x!.blip ? blipOf(bytesUri(x!.blip.bytes, x!.blip.ext)) : null;
        if (pib != null) raw = withPib(raw, pib);
        const data = wb.files[ctl.activeXBinPath!]!;
        const obj = activeXObjBody(ctl, id, ctlsStream.len, data.length, cenv);
        if (!obj) return;
        ctlsStream.bytes(data);
        items.push({ sp: stripTextbox(raw), obj, ctl: true });
        return;
      }
      if (!anchor && !x) return;
      const txo = controlTxo(ctl);
      let raw: Uint8Array;
      if (x) {
        raw = withSpid(x.raw, spid(k++));
        if (x.grouped) raw = ungroupedShape(raw, anchor && ctlMoved(ctl) ? anchor : x.grouped);
        else if (anchor && ctlMoved(ctl)) raw = withAnchor(raw, anchor);
      } else raw = newControlShape(spid(k++), anchor!, !!txo);
      const boxed = hasTextbox(raw);
      items.push({ sp: stripTextbox(raw), obj: ctlObjBody(ctl, id, cenv), txo: boxed ? txo ?? EMPTY_TXO : undefined, ctl: true });
    });
    const noteIds: number[] = [];
    for (const n of notes) {
      const id = nextId();
      noteIds.push(id);
      const sp = new Bytes();
      escher(sp, 0xf00a, 0x2, 202, body().u32(spid(k++)).u32(0xa00));
      const props: [number, number][] = [[0x0080, 0], [0x00bf, 0x00080008], [0x0158, 0], [0x0181, 0x08000050], [0x0183, 0x08000050], [0x01bf, 0x00100010], [0x0201, 0], [0x023f, 0x00030003], [0x03bf, 0x00020002]];
      const o = body();
      for (const [pid, v] of props) o.u16(pid).u32(v);
      escher(sp, 0xf00b, 0x3, props.length, o);
      const a = { fromCol: n.c + 2, fromRow: Math.max(1, n.r), fromColOff: 15, fromRowOff: 10, toCol: n.c + 4, toRow: n.r + 5, toColOff: 15, toRowOff: 4 };
      escher(sp, 0xf010, 0, 0, anchorBytes(a));
      escher(sp, 0xf011, 0, 0, body());
      escher(sp, 0xf00d, 0, 0, body());
      const obj = cmo(0x19, id, 0x4011).u16(0x0d).u16(0x16).bytes(md5(new TextEncoder().encode(`${si}:${n.r}:${n.c}`))).u16(0).u32(0);
      items.push({ sp: container(0xf004, sp).out().slice(0, -8), obj, txo: { text: n.text, runs: [[0, 0]], grbit: 0x0212, rot: 0 } });
    }
    drawings.push({ dg, shapes: k });
    // Patriarch group, then the shapes; the containers' lengths span every MSODRAWING record.
    const patriarch = new Bytes();
    escher(patriarch, 0xf009, 1, 0, body().zeros(16));
    escher(patriarch, 0xf00a, 2, 0, body().u32(spid(0)).u32(5));
    const patriarchC = container(0xf004, patriarch);
    const lenOf = (it: Item): number => (it.kids ? 8 + it.sp.length + it.kids.reduce((n, c) => n + lenOf(c), 0) : it.sp.length + (it.txo ? 8 : 0));
    const spgrLen = patriarchC.len + items.reduce((n, it) => n + lenOf(it), 0);
    const fdg = new Bytes();
    escher(fdg, 0xf008, 0, dg, body().u32(k).u32(spid(k - 1)));
    const dgLen = fdg.len + 8 + spgrLen;
    let head: Bytes | null = new Bytes().u16(0xf).u16(0xf002).u32(dgLen).bytes(fdg.out()).u16(0xf).u16(0xf003).u32(spgrLen).bytes(patriarchC.out());
    const emit = (it: Item): void => {
      const chunk = new Bytes();
      if (head) { chunk.bytes(head.out()); head = null; }
      if (it.kids) chunk.u16(0xf).u16(0xf003).u32(lenOf(it) - 8);
      chunk.bytes(it.sp);
      rs.rec(0x00ec, chunk.out());
      rs.rec(0x005d, it.ctl ? it.obj : it.obj.u16(0).u16(0));
      if (it.chart) rs.bytes(it.chart);
      if (it.txo) writeTxo(rs, it.txo);
      for (const c of it.kids ?? []) emit(c);
    };
    items.forEach(emit);
    notes.forEach((n, i) => rs.rec(0x001c, body().u16(n.r).u16(n.c).u16(0).u16(noteIds[i]!).str16(n.author.slice(0, 255) || " ").u8(0)));
  }

  /** A shape's text as a TXO: its own runs while the text is as read, else one run from the model. */
  function shapeTxo(sh: SheetShape): { text: string; runs: [number, number][]; grbit: number; rot: number } | undefined {
    const text = sh.text ?? "";
    if (!text) return undefined;
    const h = sh.textAlign === "center" ? 2 : sh.textAlign === "right" ? 3 : 1;
    const v = sh.textValign === "middle" ? 2 : sh.textValign === "bottom" ? 3 : 1;
    const grbit = 0x0200 | (h << 1) | (v << 4);
    const kept = sh.xls?.text;
    if (kept && kept.text === text && !sh.styleDirty) {
      const runs = kept.runs.map((r): [number, number] => [r.ich, ifnt({ height: r.font.height, italic: r.font.italic, strike: r.font.strike, weight: r.font.weight, underline: r.font.underline, escapement: r.font.escapement, color: r.font.color ? palette.index(r.font.color, 0x7fff) : 0x7fff, name: r.font.name })]);
      if (runs.length) return { text, runs, grbit, rot: 0 };
    }
    const f: FontKey = { ...baseFont };
    if (sh.textSize) f.height = Math.round(sh.textSize * 20);
    if (sh.textColor) f.color = palette.index(sh.textColor, 0x7fff);
    return { text, runs: [[0, ifnt(f)]], grbit, rot: 0 };
  }

  function writeTxo(rs: RecStream, txo: { text: string; runs: [number, number][]; grbit: number; rot: number }): void {
    rs.rec(0x00ec, body().u16(0).u16(0xf00d).u32(0));
    const t = txo.text.replace(/\r?\n/g, "\n").slice(0, 32767);
    const runs = txo.runs.filter(([ich]) => ich < t.length);
    if (t.length && (!runs.length || runs[0]![0] !== 0)) runs.unshift([0, runs[0]?.[1] ?? 0]);
    rs.rec(0x01b6, body().u16(txo.grbit).u16(txo.rot).zeros(6).u16(t.length).u16(t.length ? (runs.length + 1) * 8 : 0).u16(0).u32(0));
    if (!t.length) return;
    // Text CONTINUEs: a long text splits on character boundaries, each with its own flag byte.
    const wide = /[^\u0000-ÿ]/.test(t);
    const per = wide ? Math.floor((MAX_REC - 1) / 2) : MAX_REC - 1;
    for (let o = 0; o < t.length; o += per) rs.rec(0x003c, body().chars(t.slice(o, o + per)));
    const rb = body();
    for (const [ich, f] of runs.slice(0, 1000)) rb.u16(ich).u16(f).u32(0);
    rs.rec(0x003c, rb.u16(t.length).u16(0).u32(0));
  }

  /** Whether a control moved since it was read. */
  function ctlMoved(c: SheetControl): boolean {
    const a = c.anchor, b = c.xls?.anchor;
    return !!a && (!b || a.fromCol !== b.fromCol || a.fromRow !== b.fromRow || a.toCol !== b.toCol || a.toRow !== b.toRow || a.fromColOff !== b.fromColOff || a.fromRowOff !== b.fromRowOff || a.toColOff !== b.toColOff || a.toRowOff !== b.toRowOff);
  }

  /** A control's caption as a TXO: its own runs while unchanged, else one run in the default font. */
  function controlTxo(c: SheetControl): { text: string; runs: [number, number][]; grbit: number; rot: number } | undefined {
    const text = c.label ?? "";
    const grbit = c.xls?.txoGrbit ?? CAPTION_GRBIT[c.kind];
    if (grbit == null) return undefined;
    const kept = c.xls?.text;
    if (kept && kept.text === text) {
      const runs = kept.runs.map((r): [number, number] => [r.ich, ifnt({ height: r.font.height, italic: r.font.italic, strike: r.font.strike, weight: r.font.weight, underline: r.font.underline, escapement: r.font.escapement, color: r.font.color ? palette.index(r.font.color, 0x7fff) : 0x7fff, name: r.font.name })]);
      if (runs.length) return { text, runs, grbit, rot: 0 };
    }
    return { text, runs: text ? [[0, ifnt(baseFont)]] : [], grbit, rot: 0 };
  }

  /** The tokens of a control's macro: another workbook's by NameX, this one's by its NAME. */
  function macroRgce(text: string): Uint8Array | null {
    const ext = /^\[(\d+)\]!(.+)$/.exec(text);
    const env = envFor(-1);
    try {
      const t = ext ? env.externName(Number(ext[1]), ext[2]!) : { ixti: env.xti(-2, -2), index: macroName(text.replace(/^'(.*)'$/, "$1")) };
      return body().u8(0x39).u16(t.ixti).u16(t.index).u16(0).out().slice();
    } catch {
      return null;
    }
  }

  /** Whether a shape moved since it was read. */
  function moved(sh: SheetShape): boolean {
    const a = sh.anchor, b = sh.xls?.anchor;
    return !b || a.fromCol !== b.fromCol || a.fromRow !== b.fromRow || a.toCol !== b.toCol || a.toRow !== b.toRow || a.fromColOff !== b.fromColOff || a.fromRowOff !== b.fromRowOff || a.toColOff !== b.toColOff || a.toRowOff !== b.toRowOff;
  }

  /** The chart writer's view of this workbook: references, fonts, colours, formats. */
  function chartEnv(si: number): ChartWriteEnv {
    return {
      ref: (text) => compile(text, envFor(si, true), "R")?.rgce ?? null,
      font: (st, bold, size) => ifnt({ ...baseFont, height: Math.round((size ?? st?.size ?? 10) * 20), weight: (bold ?? st?.bold) ? 700 : 400, italic: !!st?.italic, color: st?.color ? palette.index(st.color, 0x7fff) : 0x7fff, name: st?.font ?? "Arial" }),
      icv: (css) => palette.index(css, 8),
      fmt: (code) => ifmtOf(code),
      warn: () => undefined,
    };
  }

  function blipOf(uri: string): number | null {
    const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(uri);
    if (!m) return null;
    const mime = m[1]!.toLowerCase();
    const raw = m[2] ? Uint8Array.from(atob(m[3]!), (ch) => ch.charCodeAt(0)) : new TextEncoder().encode(decodeURIComponent(m[3]!));
    let type: number, data = raw;
    if (mime === "image/png") type = 6;
    else if (mime === "image/jpeg" || mime === "image/jpg") type = 5;
    else if (mime === "image/bmp") { type = 7; data = raw.subarray(14); }
    else if (mime === "image/emf" || mime === "image/x-emf") type = 2;
    else if (mime === "image/wmf" || mime === "image/x-wmf") { type = 3; if (raw[0] === 0xd7 && raw[1] === 0xcd) data = raw.subarray(22); }
    else return null;
    blips.push({ type, data });
    return blips.length;
  }

  function writeCondFormats(rs: RecStream, s: Sheet, env: CompileEnv): void {
    let id = 1;
    for (const cf of s.condFormats ?? []) {
      const ranges = cf.ranges.filter((g) => g.r1 <= MAX_ROWS && g.c1 <= MAX_COLS).map((g) => ({ r1: g.r1, c1: g.c1, r2: Math.min(g.r2, MAX_ROWS), c2: Math.min(g.c2, MAX_COLS) }));
      if (!ranges.length) continue;
      const top = ranges[0]!;
      const rules: Bytes[] = [];
      for (const rule of [...cf.rules].sort((a, b) => a.priority - b.priority)) {
        const rec = cfRecord(rule, env, top);
        if (rec) rules.push(rec); else report.droppedRules++;
      }
      for (let i = 0; i < rules.length; i += 3) {
        const chunk = rules.slice(i, i + 3);
        const bound = { r1: Math.min(...ranges.map((g) => g.r1)), c1: Math.min(...ranges.map((g) => g.c1)), r2: Math.max(...ranges.map((g) => g.r2)), c2: Math.max(...ranges.map((g) => g.c2)) };
        const b = body().u16(chunk.length).u16(id++ << 1).u16(bound.r1 - 1).u16(bound.r2 - 1).u16(bound.c1 - 1).u16(bound.c2 - 1).u16(ranges.length);
        for (const g of ranges) b.u16(g.r1 - 1).u16(g.r2 - 1).u16(g.c1 - 1).u16(g.c2 - 1);
        rs.rec(0x01b0, b);
        for (const r of chunk) rs.rec(0x01b1, r);
      }
    }
  }

  function cfRecord(rule: CfRule, env: CompileEnv, top: { r1: number; c1: number }): Bytes | null {
    const OPS: Record<string, number> = { between: 1, notBetween: 2, equal: 3, notEqual: 4, greaterThan: 5, lessThan: 6, greaterThanOrEqual: 7, lessThanOrEqual: 8 };
    const a1 = colToLetters(top.c1) + top.r1;
    const q = (t: string): string => '"' + t.replace(/"/g, '""') + '"';
    let ct: number, op = 0, f1: string | undefined, f2: string | undefined;
    switch (rule.type) {
      case "cellIs": ct = 1; op = OPS[rule.operator ?? ""] ?? 0; f1 = rule.formulas?.[0]; f2 = op <= 2 ? rule.formulas?.[1] : undefined; if (!op) return null; break;
      case "expression": ct = 2; f1 = rule.formulas?.[0]; break;
      case "containsText": ct = 2; f1 = rule.formulas?.[0] ?? `NOT(ISERROR(SEARCH(${q(rule.text ?? "")},${a1})))`; break;
      case "notContainsText": ct = 2; f1 = rule.formulas?.[0] ?? `ISERROR(SEARCH(${q(rule.text ?? "")},${a1}))`; break;
      case "beginsWith": ct = 2; f1 = rule.formulas?.[0] ?? `LEFT(${a1},${(rule.text ?? "").length})=${q(rule.text ?? "")}`; break;
      case "endsWith": ct = 2; f1 = rule.formulas?.[0] ?? `RIGHT(${a1},${(rule.text ?? "").length})=${q(rule.text ?? "")}`; break;
      case "timePeriod": ct = 2; f1 = rule.formulas?.[0] ?? (rule.timePeriod ? timePeriodFormula(rule.timePeriod, a1) : undefined); break;
      default: return null;
    }
    if (!f1) return null;
    const c1 = compile(f1, env), c2 = f2 ? compile(f2, env) : null;
    if (!c1 || (f2 && !c2)) return null;
    const d = rule.dxf ?? {};
    const hasFont = d.bold != null || d.italic != null || !!d.color;
    const hasPat = !!d.bg;
    let flags = 0x003fffff;
    if (hasPat) flags &= ~(0x10000 | 0x20000 | 0x40000);
    if (hasFont) flags |= 0x04000000;
    if (hasPat) flags |= 0x20000000;
    const b = body().u8(ct).u8(op).u16(c1.rgce.length).u16(c2?.rgce.length ?? 0).u32(flags).u16(0x8002);
    if (hasFont) {
      b.zeros(64).u32(0xffffffff).u32(d.italic ? 2 : 0).u16(d.bold ? 700 : d.bold === false ? 400 : 0xffff).u16(0).u8(0).u8(0).u8(0).u8(0);
      b.u32(d.color ? palette.index(d.color, 0x7fff) : 0xffffffff).u32(0).u32((d.italic == null ? 2 : 0) | 0x80).u32(1).u32(1).u32(d.bold == null ? 1 : 0).u32(0).u32(0xffffffff).u32(0xffffffff).u16(1);
    }
    if (hasPat) { const icv = palette.index(d.bg, 65); b.u16(1 << 10).u16(icv | (icv << 7)); }
    b.bytes(c1.rgce);
    if (c2) b.bytes(c2.rgce);
    return b;
  }

  function writeLinks(rs: RecStream, s: Sheet): void {
    const STD = [0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
    const URL = [0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
    for (const cell of s.cells.values()) {
      const l = cell.link;
      if (!l || cell.row > MAX_ROWS || cell.col > MAX_COLS) continue;
      const r = cell.row - 1, c = cell.col - 1;
      const b = body().u16(r).u16(r).u16(c).u16(c).bytes(STD).u32(2);
      if (l.internal) {
        const loc = l.href.replace(/^#/, "");
        b.u32(0x08).u32(loc.length + 1).utf16z(loc);
      } else {
        const hash = l.href.indexOf("#");
        const url = hash >= 0 && !/^[a-z]+:\/\//i.test(l.href.slice(0, hash)) ? l.href : hash >= 0 ? l.href.slice(0, hash) : l.href;
        const loc = hash >= 0 && url !== l.href ? l.href.slice(hash + 1) : "";
        b.u32(0x03 | (loc ? 0x08 : 0)).bytes(URL).u32((url.length + 1) * 2).utf16z(url);
        if (loc) b.u32(loc.length + 1).utf16z(loc);
      }
      rs.rec(0x01b8, b);
      if (l.tip) rs.rec(0x0800, body().u16(0x0800).u16(r).u16(r).u16(c).u16(c).utf16z(l.tip.slice(0, 255)));
    }
  }

  function writeValidations(rs: RecStream, s: Sheet, env: CompileEnv): void {
    const TYPES: Record<string, number> = { whole: 1, decimal: 2, list: 3, date: 4, time: 5, textLength: 6, custom: 7 };
    const OPS: Record<string, number> = { between: 0, notBetween: 1, equal: 2, notEqual: 3, greaterThan: 4, lessThan: 5, greaterThanOrEqual: 6, lessThanOrEqual: 7 };
    const ERRS: Record<string, number> = { stop: 0, warning: 1, information: 2 };
    const recs: Bytes[] = [];
    for (const dv of s.validations ?? []) {
      const ranges = dv.ranges.filter((g) => g.r1 <= MAX_ROWS && g.c1 <= MAX_COLS);
      if (!ranges.length) continue;
      const type = TYPES[dv.type ?? "list"] ?? 3;
      let f1: Uint8Array | null = null, f2: Uint8Array | null = null, strList = false;
      if (type === 3) {
        if (dv.values?.length) {
          const joined = dv.values.join("\u0000");
          if (joined.length > 255) { report.droppedRules++; continue; }
          const tb = body().u8(0x17).u8(joined.length).chars(joined);
          f1 = tb.out().slice();
          strList = true;
        } else if (dv.rangeRef) f1 = compile(dv.rangeRef, env, "R")?.rgce ?? null;
      } else {
        if (dv.formula1) f1 = compile(dv.formula1, env)?.rgce ?? null;
        if (dv.formula2) f2 = compile(dv.formula2, env)?.rgce ?? null;
      }
      if (!f1 && type !== 0) { report.droppedRules++; continue; }
      const flags = type | ((ERRS[dv.errorStyle ?? "stop"] ?? 0) << 4) | (strList ? 0x80 : 0) | (dv.allowBlank ? 0x100 : 0) | 0x40000 | 0x80000 | ((OPS[dv.operator ?? "between"] ?? 0) << 20);
      const str = (t?: string): string => (t ? t.slice(0, 255) : "\u0000");
      const b = body().u32(flags).str16(str(dv.promptTitle)).str16(str(dv.errorTitle)).str16(str(dv.promptMessage)).str16(str(dv.errorMessage));
      b.u16(f1?.length ?? 0).u16(0).bytes(f1 ?? []).u16(f2?.length ?? 0).u16(0).bytes(f2 ?? []);
      b.u16(ranges.length);
      for (const g of ranges) b.u16(g.r1 - 1).u16(Math.min(g.r2, MAX_ROWS) - 1).u16(g.c1 - 1).u16(Math.min(g.c2, MAX_COLS) - 1);
      recs.push(b);
    }
    if (!recs.length) return;
    rs.rec(0x01b2, body().u16(0).u32(0).u32(0).u32(0xffffffff).u32(recs.length));
    for (const r of recs) rs.rec(0x01be, r);
  }

  // ---- globals ----------------------------------------------------------------------------
  const g = new RecStream();
  g.rec(0x0809, body().u16(0x0600).u16(0x0005).u16(0x0dbb).u16(0x07cc).u32(0x000100c1).u32(0x00000006));
  g.rec(0x00e1, body().u16(1200));
  g.rec(0x00c1, body().u16(0));
  g.rec(0x00e2);
  const wa = body().str16("sheetedit");
  while (wa.len < 112) wa.u8(0x20);
  g.rec(0x005c, wa);
  g.rec(0x0042, body().u16(1200));
  g.rec(0x0161, body().u16(0));
  const tabid = body();
  sheets.forEach((_, i) => tabid.u16(i + 1));
  g.rec(0x013d, tabid);
  if (vbaBin) g.rec(0x00d3); // OBPROJ: the file carries a VBA project
  if (vbaBin || wb.xls?.codeName) g.rec(0x01ba, body().str16(wb.xls?.codeName || "ThisWorkbook"));
  g.rec(0x009c, body().u16(14));
  const wp = wb.protection;
  g.rec(0x0019, body().u16(wp?.windows ? 1 : 0));
  g.rec(0x0012, body().u16(wp?.structure ? 1 : 0));
  g.rec(0x0013, body().u16(wp?.password?.legacy ? parseInt(wp.password.legacy, 16) || 0 : 0));
  g.rec(0x01af, body().u16(0));
  g.rec(0x01bc, body().u16(0));
  const active = Math.max(0, Math.min(wb.activeSheet ?? 0, sheets.length - 1));
  g.rec(0x003d, body().u16(0x0168).u16(0x001e).u16(0x3a5c).u16(0x2310).u16(0x0038).u16(active).u16(0).u16(1).u16(0x0258));
  g.rec(0x0040, body().u16(0));
  g.rec(0x008d, body().u16(0));
  g.rec(0x0022, body().u16(0));
  g.rec(0x000e, body().u16(1));
  g.rec(0x01b7, body().u16(0));
  g.rec(0x00da, body().u16(0));
  for (const f of fonts) {
    g.rec(0x0031, body().u16(f.height).u16((f.italic ? 2 : 0) | (f.strike ? 8 : 0)).u16(f.color).u16(f.weight).u16(f.escapement).u8(f.underline).u8(0).u8(0).u8(0).str8(f.name));
  }
  for (const [id, code] of formatRecs) g.rec(0x041e, body().u16(id).str16(code.slice(0, 255)));
  const xfRec = (k: XfKey, style: boolean): Bytes => {
    const [l, r, t, bm] = k.b;
    const b1 = (l![0] | (r![0] << 4) | (t![0] << 8) | (bm![0] << 12) | (l![1] << 16) | (r![1] << 23) | (k.diag[0] << 30)) >>> 0;
    const b2 = (t![1] | (bm![1] << 7) | (k.diag[2] << 14) | (k.diag[1] << 21) | (k.fls << 26)) >>> 0;
    return body().u16(k.font).u16(k.fmt).u16(style ? 0xfff5 : (k.locked ? 1 : 0) | (k.hidden ? 2 : 0))
      .u8(k.h | (k.wrap ? 8 : 0) | (k.v << 4)).u8(k.rot).u8(k.indent | (k.shrink ? 16 : 0)).u8(style ? 0 : 0xfc)
      .u32(b1).u32(b2).u16((k.fore & 0x7f) | ((k.back & 0x7f) << 7));
  };
  const plain: XfKey = { font: 0, fmt: 0, locked: true, hidden: false, h: 0, v: 2, wrap: false, rot: 0, indent: 0, shrink: false, b: [[0, 0], [0, 0], [0, 0], [0, 0]], diag: [0, 0, 0], fls: 0, fore: 64, back: 65 };
  for (let i = 0; i < 15; i++) g.rec(0x00e0, xfRec({ ...plain, font: i === 1 || i === 2 ? 1 : i === 3 || i === 4 ? 2 : 0 }, true));
  g.rec(0x00e0, xfRec(defaultKey, false));
  for (const k of xfKeys) g.rec(0x00e0, xfRec(k, false));
  g.rec(0x0293, body().u16(0x8000).u8(0).u8(0xff));
  if (palette.changed || srcPalette.some((c, i) => c !== DEFAULT_PALETTE[i])) {
    const b = body().u16(56);
    for (const c of palette.colors) b.u8(parseInt(c.slice(1, 3), 16)).u8(parseInt(c.slice(3, 5), 16)).u8(parseInt(c.slice(5, 7), 16)).u8(0);
    g.rec(0x0092, b);
  }
  g.rec(0x0160, body().u16(1));
  const boundAt: number[] = [];
  sheets.forEach((s) => {
    boundAt.push(g.len + 4);
    const name = s.name.replace(/[\\/?*[\]:]/g, "_").slice(0, 255) || "Sheet";
    g.rec(0x0085, body().u32(0).u8(s.visibility === "hidden" ? 1 : s.visibility === "veryHidden" ? 2 : 0).u8(s.chartSheet && s.charts?.length ? 2 : 0).str8(name));
  });
  g.rec(0x008c, body().u16(1).u16(1));
  if (names.length || xtis.length) {
    g.rec(0x01ae, body().u16(sheets.length).u16(0x0401));
    for (const ext of externals) {
      const b = body().u16(ext.sheets.length).u16(ext.path.length).chars(ext.path);
      for (const sh of ext.sheets) b.str16(sh);
      g.rec(0x01ae, b);
      ext.names.forEach((nm, i) => g.rec(0x0023, ext.rawNames?.[i] ?? body().u16(0).u32(0).str8(nm).u16(2).u8(0x1c).u8(0x17)));
      for (const c of ext.cache ?? []) g.rec(c.type, c.data);
    }
    if (addinNames.length) {
      g.rec(0x01ae, body().u16(1).u16(0x3a01));
      for (const nm of addinNames) g.rec(0x0023, body().u16(0).u32(0).str8(nm).u16(2).u8(0x1c).u8(0x17));
    }
    if (!xtis.length) xtiOf(0, 0, 0);
    const b = body().u16(xtis.length);
    for (const [bk, f, l] of xtis) b.u16(bk).u16(f).u16(l);
    g.rec(0x0017, b);
    for (const n of names) {
      const rgce = n.rgce ?? (n.func ? new Uint8Array([0x1c, 0x1d]) : new Uint8Array(0));
      const rgcb = n.rgcb ?? new Uint8Array(0);
      const text = n.builtin != null ? String.fromCharCode(n.builtin) : n.name.slice(0, 255);
      const grbit = n.grbit ?? (n.hidden ? 1 : 0) | (n.func ? 2 : 0) | (n.builtin != null ? 0x20 : 0);
      g.rec(0x0018, body().u16(grbit).u8(0).u8(text.length).u16(rgce.length).u16(0).u16(n.itab).u32(0).chars(text).bytes(rgce).bytes(rgcb));
    }
  }
  if (blips.length || drawings.length) writeDrawingGroup(g);
  writeSst(g);
  g.rec(0x000a);

  // ---- assemble: patch sheet offsets, INDEX and DBCELL positions ----------------------------
  let at = g.len;
  const gOut = g.out();
  const parts: Uint8Array[] = [gOut];
  const gv = new DataView(gOut.buffer, gOut.byteOffset);
  sheetBodies.forEach((sb, i) => {
    gv.setUint32(boundAt[i]!, at, true);
    // INDEX: the DEFCOLWIDTH position, then each DBCELL position, all absolute.
    if (sb.indexAt >= 0) {
      sb.rs.set32(sb.indexAt + 12, at + sb.defColAt);
      sb.dbcells.forEach((d, k) => sb.rs.set32(sb.indexAt + 16 + 4 * k, at + d));
    }
    parts.push(sb.rs.out());
    at += sb.rs.len;
  });
  const stream = new Bytes();
  for (const p of parts) stream.bytes(p);
  const entries = [{ name: "Workbook", data: stream.out() }];
  const summary = wb.xls?.summary;
  if (summary?.size) for (const [k, v] of summary) entries.push({ name: k, data: v });
  else entries.push({ name: "\u0005SummaryInformation", data: propertySet(SUMMARY_FMTID) }, { name: "\u0005DocumentSummaryInformation", data: propertySet(DOC_SUMMARY_FMTID) });
  // The macro project goes back verbatim (or as edited) under _VBA_PROJECT_CUR, its signature at the root.
  if (vbaBin) {
    try {
      for (const e of readCfbAll(vbaBin)) entries.push({ name: e.name === "_VBA_SIGNATURE" ? "\u0005DigitalSignature" : "_VBA_PROJECT_CUR/" + e.name, data: e.data });
    } catch {
      // A project that is not a compound file is left out rather than written damaged.
    }
  }
  if (ctlsStream.len) entries.push({ name: "Ctls", data: ctlsStream.out() });
  return writeCfbStreams(entries);

  function writeDrawingGroup(gs: RecStream): void {
    const bstore = new Bytes();
    const BT_INST: Record<number, number> = { 2: 0x3d4, 3: 0x216, 5: 0x46a, 6: 0x6e0, 7: 0x7a8 };
    const BT_REC: Record<number, number> = { 2: 0xf01a, 3: 0xf01b, 5: 0xf01d, 6: 0xf01e, 7: 0xf01f };
    for (const bl of blips) {
      const uid = md5(bl.data);
      const blip = body().bytes(uid);
      if (bl.type === 2 || bl.type === 3) blip.u32(bl.data.length).u32(0).u32(0).u32(1000).u32(1000).u32(914400).u32(914400).u32(bl.data.length).u8(0xfe).u8(0xfe).bytes(bl.data);
      else blip.u8(0xff).bytes(bl.data);
      const blipRec = new Bytes();
      escher(blipRec, BT_REC[bl.type]!, 0, BT_INST[bl.type]!, blip);
      const bse = body().u8(bl.type).u8(bl.type).bytes(uid).u16(0xff).u32(blipRec.len).u32(1).u32(0).u8(0).u8(0).u8(0).u8(0).bytes(blipRec.out());
      escher(bstore, 0xf007, 2, bl.type, bse);
    }
    const dgg = new Bytes();
    const spidMax = Math.max(...drawings.map((d) => d.dg * 1024 + d.shapes), 1024);
    const fdgg = body().u32(spidMax).u32(drawings.length + 1).u32(drawings.reduce((n, d) => n + d.shapes, 0)).u32(drawings.length);
    for (const d of drawings) fdgg.u32(d.dg).u32(d.shapes);
    escher(dgg, 0xf006, 0, 0, fdgg);
    if (blips.length) escher(dgg, 0xf001, 0xf, blips.length, bstore);
    escher(dgg, 0xf00b, 3, 3, body().u16(0x00bf).u32(0x00080008).u16(0x0181).u32(0x08000041).u16(0x01c0).u32(0x08000040));
    const all = new Bytes();
    escher(all, 0xf000, 0xf, 0, dgg);
    gs.rec(0x00eb, all.out());
  }

  function writeSst(gs: RecStream): void {
    // Strings may split across CONTINUE records only inside their characters (with a new flag
    // byte) or between formatting runs; EXTSST indexes every eighth string's position.
    const dsst = Math.max(8, Math.ceil(sst.length / 128));
    const ext: [number, number][] = [];
    let cur = body().u32(sstTotal).u32(sst.length);
    let recStart = gs.len;
    let first = true;
    const flush = (): void => {
      gs.u16(first ? 0x00fc : 0x003c).u16(cur.len).bytes(cur.out());
      first = false;
      recStart = gs.len;
      cur = body();
    };
    sst.forEach((s, i) => {
      const wide = /[^\u0000-ÿ]/.test(s.text);
      const headLen = 3 + (s.runs ? 2 : 0);
      if (cur.len + headLen + (wide ? 2 : 1) > MAX_REC) flush();
      if (i % dsst === 0) ext.push([recStart + 4 + cur.len, 4 + cur.len]);
      cur.u16(s.text.length).u8((wide ? 1 : 0) | (s.runs ? 8 : 0));
      if (s.runs) cur.u16(s.runs.length);
      let k = 0;
      while (k < s.text.length) {
        const room = Math.floor((MAX_REC - cur.len) / (wide ? 2 : 1));
        if (room <= 0) { flush(); cur.u8(wide ? 1 : 0); continue; }
        const end = Math.min(s.text.length, k + room);
        for (let j = k; j < end; j++) { const code = s.text.charCodeAt(j); if (wide) cur.u16(code); else cur.u8(code); }
        k = end;
      }
      for (const [ich, f] of s.runs ?? []) {
        if (cur.len + 4 > MAX_REC) flush();
        cur.u16(ich).u16(f);
      }
    });
    flush();
    const e = body().u16(dsst);
    for (const [ib, cb] of ext) e.u32(ib).u16(cb).u16(0);
    gs.rec(0x00ff, e);
  }
}

const fontPos = (i: number): number => (i < 4 ? i : i - 1);
const lettersCol = (l: string): number => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

function toRk(v: number): number | null {
  const tryInt = (n: number, div: number): number | null => (Number.isInteger(n) && n >= -0x20000000 && n < 0x20000000 ? ((n << 2) | 2 | div) >>> 0 : null);
  const i = tryInt(v, 0);
  if (i != null) return i;
  const d = new DataView(new ArrayBuffer(8));
  const asDouble = (n: number, div: number): number | null => {
    d.setFloat64(0, n, true);
    if (d.getUint32(0, true) !== 0 || (d.getUint32(4, true) & 3) !== 0) return null;
    return (d.getUint32(4, true) | div) >>> 0;
  };
  const dbl = asDouble(v, 0);
  if (dbl != null) return dbl;
  const h = Math.round(v * 100);
  if (Math.abs(h / 100 - v) < 1e-12 && h / 100 === v) {
    const ih = tryInt(h, 1);
    if (ih != null) return ih;
  }
  return null;
}

const EMPTY_TXO = { text: "", runs: [] as [number, number][], grbit: 0x0212, rot: 0 };
const DASH_CODE: Record<string, number> = { sysDash: 1, sysDot: 2, sysDashDot: 3, sysDashDotDot: 4, dot: 5, dash: 6, lgDash: 7, dashDot: 8, lgDashDot: 9, lgDashDotDot: 10 };
const END_CODE: Record<string, number> = { triangle: 1, stealth: 2, diamond: 3, oval: 4, arrow: 5 };
const colorRgb = (css: string): number => /^#[0-9a-f]{6}$/i.test(css) ? (parseInt(css.slice(1, 3), 16) | (parseInt(css.slice(3, 5), 16) << 8) | (parseInt(css.slice(5, 7), 16) << 16)) : 0;

/** The direct children of a shape container: [type, data start, data length]. */
function spChildren(raw: Uint8Array): [number, number, number][] {
  const out: [number, number, number][] = [];
  for (let p = 8; p + 8 <= raw.length;) {
    const type = raw[p + 2]! | (raw[p + 3]! << 8);
    const len = (raw[p + 4]! | (raw[p + 5]! << 8) | (raw[p + 6]! << 16) | (raw[p + 7]! << 24)) >>> 0;
    out.push([type, p + 8, len]);
    p += 8 + len;
  }
  return out;
}
const hasTextbox = (raw: Uint8Array): boolean => spChildren(raw).some(([t]) => t === 0xf00d);
/** A copy of a shape container with a new shape id. */
function withSpid(raw: Uint8Array, id: number): Uint8Array {
  const out = raw.slice();
  const sp = spChildren(out).find(([t]) => t === 0xf00a);
  if (sp) new DataView(out.buffer).setUint32(sp[1], id, true);
  return out;
}
/** A shape container with its cell anchor replaced. */
function withAnchor(raw: Uint8Array, anchor: Uint8Array): Uint8Array {
  const a = spChildren(raw).find(([t, , l]) => t === 0xf010 && l === anchor.length);
  if (a) raw.set(anchor, a[1]);
  return raw;
}
/** The container without its trailing ClientTextbox, which goes in a record of its own. */
function stripTextbox(raw: Uint8Array): Uint8Array {
  const kids = spChildren(raw);
  const last = kids[kids.length - 1];
  return last && last[0] === 0xf00d && last[1] + last[2] === raw.length ? raw.subarray(0, last[1] - 8) : raw;
}

function escher(out: Bytes, type: number, ver: number, inst: number, data: Bytes): void {
  out.u16((ver & 0xf) | (inst << 4)).u16(type).u32(data.len).bytes(data.out());
}
function container(type: number, kids: Bytes): Bytes {
  const b = new Bytes();
  escher(b, type, 0xf, 0, kids);
  return b;
}

function bytesUri(bytes: Uint8Array, ext: string): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "bmp" ? "image/bmp" : ext === "emf" ? "image/emf" : ext === "wmf" ? "image/wmf" : "image/" + ext;
  return `data:${mime};base64,${btoa(bin)}`;
}

const SUMMARY_FMTID = [0xe0, 0x85, 0x9f, 0xf2, 0xf9, 0x4f, 0x68, 0x10, 0xab, 0x91, 0x08, 0x00, 0x2b, 0x27, 0xb3, 0xd9];
const DOC_SUMMARY_FMTID = [0x02, 0xd5, 0xcd, 0xd5, 0x9c, 0x2e, 0x1b, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2b, 0x2c, 0xf9, 0xae];

/** A minimal property set stream holding only its code page. */
function propertySet(fmtid: number[]): Uint8Array {
  const b = new Bytes();
  b.u16(0xfffe).u16(0).u32(0x00020006).zeros(16).u32(1).bytes(fmtid).u32(48);
  b.u32(8 + 8 + 8).u32(1).u32(1).u32(16).u32(2).u16(1252).u16(0);
  return b.out().slice();
}
