// BIFF8 writer: a workbook model (with its xlsx style pool) as an Excel 97-2003 .xls compound file.
// Styles are rebuilt from the live styles.xml so every edit made through the xlsx paths lands in
// the XF / FONT / FORMAT tables; formulas are compiled to tokens with their cached results.
import type { Cell, CfRule, Sheet, TextRun, Workbook } from "../../core/model";
import { colToLetters, firstByLocal } from "../../core/model";
import { SHEET_LOCK_DEFAULTS, type SheetLock } from "../../core/protection";
import { argbToCss, colWidthToPx, maxDigitWidth, readTheme, resolveColor } from "../xlsx/read";
import { colWidthChars, timePeriodFormula } from "../xlsx/write";
import { body, Bytes, MAX_REC, RecStream } from "./bytes";
import { writeCfbStreams } from "./cfb";
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
  const sheetBodies = sheets.map((s, si) => writeSheet(s, si));

  function writeSheet(s: Sheet, si: number): { rs: RecStream; indexAt: number; dbcells: number[]; defColAt: number } {
    const rs = new RecStream();
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
      if (str != null) rs.rec(0x0207, body().str16(str));
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

  // ---- drawing objects: comments and pictures ------------------------------------------------
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
    if (!notes.length && !pics.length) return;
    const dg = drawings.length + 1;
    const total = notes.length + pics.length;
    drawings.push({ dg, shapes: total + 1 });
    const spid = (k: number): number => dg * 1024 + k;
    // Each shape container, and where its client data / textbox ends (record split points).
    type Part = { bytes: Bytes; obj: Bytes; txo?: { text: string } };
    const parts: Part[] = [];
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
    let k = 1;
    let objId = 1;
    for (const pic of pics) {
      const id = objId++;
      const sp = new Bytes();
      escher(sp, 0xf00a, 0x2, 75, body().u32(spid(k++)).u32(0xa00));
      escher(sp, 0xf00b, 0x3, 2, body().u16(0x4104).u32(pic.blip).u16(0x03bf).u32(0x00080000));
      escher(sp, 0xf010, 0, 0, anchorBytes(pic.anchor));
      escher(sp, 0xf011, 0, 0, body());
      const obj = body().u16(0x15).u16(0x12).u16(8).u16(id).u16(0x6011).zeros(12).u16(0x07).u16(2).u16(0xffff).u16(0x08).u16(2).u16(1).u16(0).u16(0);
      parts.push({ bytes: container(0xf004, sp), obj });
    }
    const noteIds: number[] = [];
    for (const n of notes) {
      const id = objId++;
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
      const obj = body().u16(0x15).u16(0x12).u16(0x19).u16(id).u16(0x4011).zeros(12).u16(0x0d).u16(0x16).bytes(md5(new TextEncoder().encode(`${si}:${n.r}:${n.c}`))).u16(0).u32(0).u16(0).u16(0);
      parts.push({ bytes: container(0xf004, sp), obj, txo: { text: n.text } });
    }
    // Patriarch group, then the shapes; the containers' lengths span every MSODRAWING record.
    const patriarch = new Bytes();
    escher(patriarch, 0xf009, 1, 0, body().zeros(16));
    escher(patriarch, 0xf00a, 2, 0, body().u32(spid(0)).u32(5));
    const patriarchC = container(0xf004, patriarch);
    const textboxLen = 8;
    let spgrLen = patriarchC.len;
    for (const p of parts) spgrLen += p.bytes.len + (p.txo ? textboxLen : 0);
    const fdg = new Bytes();
    escher(fdg, 0xf008, 0, dg, body().u32(total + 1).u32(spid(k - 1)));
    const dgLen = fdg.len + 8 + spgrLen;
    const headBytes = new Bytes();
    headBytes.u16(0xf).u16(0xf002).u32(dgLen).bytes(fdg.out()).u16(0xf).u16(0xf003).u32(spgrLen).bytes(patriarchC.out());
    parts.forEach((p, i) => {
      const chunk = i === 0 ? new Bytes().bytes(headBytes.out()).bytes(p.bytes.out()) : p.bytes;
      rs.rec(0x00ec, chunk.out());
      rs.rec(0x005d, p.obj);
      if (p.txo) {
        rs.rec(0x00ec, body().u16(0).u16(0xf00d).u32(0));
        const t = p.txo.text.replace(/\r?\n/g, "\n").slice(0, 32767);
        rs.rec(0x01b6, body().u16(0x0212).u16(0).zeros(6).u16(t.length).u16(t.length ? 16 : 0).u16(0).u32(0));
        if (t.length) {
          const tb = body().chars(t);
          // Text CONTINUEs: a long note splits on character boundaries, each with its own flag byte.
          const wide = tb.out()[0] === 1;
          const per = wide ? Math.floor((MAX_REC - 1) / 2) : MAX_REC - 1;
          for (let o = 0; o < t.length; o += per) rs.rec(0x003c, body().chars(wide ? t.slice(o, o + per) : t.slice(o, o + per)));
          rs.rec(0x003c, body().u16(0).u16(0).u32(0).u16(t.length).u16(0).u32(0));
        }
      }
    });
    notes.forEach((n, i) => rs.rec(0x001c, body().u16(n.r).u16(n.c).u16(0).u16(noteIds[i]!).str16(n.author.slice(0, 255) || " ").u8(0)));
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
    g.rec(0x0085, body().u32(0).u8(s.visibility === "hidden" ? 1 : s.visibility === "veryHidden" ? 2 : 0).u8(0).str8(name));
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
  if (blips.length) writeDrawingGroup(g);
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
    sb.rs.set32(sb.indexAt + 12, at + sb.defColAt);
    sb.dbcells.forEach((d, k) => sb.rs.set32(sb.indexAt + 16 + 4 * k, at + d));
    parts.push(sb.rs.out());
    at += sb.rs.len;
  });
  const stream = new Bytes();
  for (const p of parts) stream.bytes(p);
  const entries = [{ name: "Workbook", data: stream.out() }];
  const summary = wb.xls?.summary;
  if (summary?.size) for (const [k, v] of summary) entries.push({ name: k, data: v });
  else entries.push({ name: "\u0005SummaryInformation", data: propertySet(SUMMARY_FMTID) }, { name: "\u0005DocumentSummaryInformation", data: propertySet(DOC_SUMMARY_FMTID) });
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
    escher(dgg, 0xf001, 0xf, blips.length, bstore);
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
