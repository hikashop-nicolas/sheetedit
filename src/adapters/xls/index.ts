// Legacy Excel .xls (BIFF8, plus BIFF2-7 on read): the workbook is converted to xlsx parts on open,
// so every xlsx editing path applies, and written back as a BIFF8 compound file on save.
import type { ChartModel } from "../../core/chart-model";
import type { Workbook } from "../../core/model";
import { readXlsShapes, sptOf, xlsChartModel } from "./drawing";
import { readXlsx } from "../xlsx/read";
import { isCfb, writeCfbStreams } from "./cfb";
import { isRawBiff, readXlsBook, type XlsReadOptions } from "./read";
import type { XlsBook } from "./types";
import { xlsBookToXlsx } from "./to-xlsx";

export { XlsPasswordError, type XlsReadOptions } from "./read";
export { writeXls } from "./write";

/** What a workbook read from .xls carries beyond its xlsx parts, for the BIFF8 writer and the UI. */
export interface XlsSource {
  /** BIFF version the file was read from (2-8); it is always saved as BIFF8. */
  biff: number;
  /** The file's 56-colour palette ("#rrggbb"), reused on save so colours map back exactly. */
  palette: string[];
  /** The summary information streams, preserved verbatim. */
  summary?: Map<string, Uint8Array>;
  /** External workbooks referenced by formulas, by their 1-based [n] in formula text. */
  externals: { path: string; sheets: string[]; names: string[]; rawNames?: Uint8Array[]; cache?: { type: number; data: Uint8Array }[] }[];
  /** Chart sheets / VB modules that were not loaded, and embedded charts and drawing shapes. */
  skippedSheets: string[];
  charts: number;
  shapes: number;
  /** User-defined (macro) function names, which formulas call by name rather than as _xlfn. */
  udfs: string[];
  encrypted?: "xor" | "rc4" | "cryptoapi";
  /** The workbook's VBA code name ("ThisWorkbook"), written back as CODENAME. */
  codeName?: string;
}

/** Whether these bytes look like a legacy .xls (compound file or raw BIFF stream). */
export const looksLikeXls = (bytes: Uint8Array): boolean => isCfb(bytes) || isRawBiff(bytes);

/** Read a .xls into a workbook the editor treats as xlsx, flagged with its BIFF origin. */
export function readXls(bytes: Uint8Array, opts: XlsReadOptions & { lazySheets?: boolean } = {}): Workbook {
  const book = readXlsBook(bytes, opts);
  const files = xlsBookToXlsx(book);
  // The VBA storage holds what an .xlsm's vbaProject.bin does, so it is read and edited as one.
  const vbaBin = book.vba ? writeCfbStreams(book.vba) : undefined;
  if (vbaBin) files["xl/vbaProject.bin"] = vbaBin;
  const wb = readXlsx(files, { lazySheets: opts.lazySheets });
  // Charts, chart sheets and drawing shapes go straight into the model.
  book.sheets.forEach((s, i) => {
    const sheet = wb.sheets[i];
    if (!sheet) return;
    const charts: ChartModel[] = [];
    if (s.chartSheet) {
      const px = (pt: number, dflt: number): number => (pt > 10 ? Math.round((pt * 96) / 72) : dflt);
      const w = px(s.chartSheet.width, 720), h = px(s.chartSheet.height, 480);
      charts.push({ ...s.chartSheet.model, id: "chart-1", anchor: { fromCol: 1, fromRow: 1, fromColOff: 8, fromRowOff: 8, toCol: 1, toRow: 1, toColOff: 8 + w, toRowOff: 8 + h } } as ChartModel);
      sheet.chartSheet = true;
      sheet.hideGridLines = true;
    }
    for (const c of s.chartObjs) { const m = xlsChartModel(c, s, `chart-${charts.length + 1}`); if (m) charts.push(m); }
    const autos = [s.chartSheet, ...s.chartObjs.filter((c) => c.anchor).map((c) => c.chart)].filter((c) => !!c);
    autos.forEach((c, k) => { const t = c.autoTitle && cellText(book, c.autoTitle); if (t && charts[k]) charts[k]!.title = t; });
    if (charts.length) sheet.charts = charts;
    const shapes = readXlsShapes(s, book);
    if (shapes.length) sheet.shapes = shapes;
  });
  wb.xls = {
    biff: book.biff,
    palette: book.palette,
    summary: book.summary,
    externals: book.supbooks.filter((s) => s.kind === "external").map((s) => ({ path: s.path ?? "", sheets: s.sheets, names: s.names, rawNames: s.rawNames, cache: s.cache })),
    skippedSheets: book.skippedSheets,
    charts: book.sheets.reduce((n, s) => n + s.charts, 0),
    shapes: book.sheets.reduce((n, s) => n + s.shapes, 0),
    udfs: book.names.filter((n) => (n.func || !n.formula) && n.builtin == null && !n.name.startsWith("_xlfn.")).map((n) => n.name),
    encrypted: book.encrypted,
    codeName: book.codeName,
  };
  return wb;
}

/** The text of the cell a reference such as "Sheet1!$B$1" names. */
function cellText(book: XlsBook, ref: string): string | undefined {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref);
  if (!m) return undefined;
  const name = (m[1] ?? m[2] ?? "").replace(/''/g, "'");
  const sh = book.sheets.find((s) => s.name === name);
  const col = [...m[3]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const v = sh?.cells.get((Number(m[4]) - 1) * 256 + col)?.val;
  return v && (v.k === "s" || v.k === "n") ? String(v.v) : undefined;
}

/** What saving this .xls cannot keep, or keeps only as its closest BIFF8 form, as i18n keys. */
export function xlsSaveWarnings(wb: Workbook): { key: string; n?: number; names?: string }[] {
  const x = wb.xls;
  if (!x) return [];
  const out: { key: string; n?: number; names?: string }[] = [];
  if (x.charts) out.push({ key: "xlsDropsCharts", n: x.charts });
  if (x.shapes) out.push({ key: "xlsDropsControls", n: x.shapes });
  if (x.skippedSheets.length) out.push({ key: "xlsDropsSheets", names: x.skippedSheets.join(", ") });
  let standIn = 0, errBars = false, combo = false;
  const xy = (k: string | undefined): boolean => k === "scatter" || k === "bubble";
  const round = (k: string | undefined): boolean => k === "pie" || k === "doughnut";
  for (const s of wb.sheets) {
    for (const sh of s.shapes ?? []) {
      const kept = sh.xls && !sh.created && !sh.styleDirty && !sh.xfrmDirty;
      if (!kept && sptOf(sh.preset, sh.geom) == null) standIn++;
    }
    for (const c of s.charts ?? []) {
      for (const ser of c.series) {
        if (ser.errorBars?.valueType === "cust") errBars = true;
        const k = ser.type ?? c.kind;
        if (k !== c.kind && (round(c.kind) || round(k) || xy(k) !== xy(c.kind))) combo = true;
      }
    }
  }
  if (standIn) out.push({ key: "xlsShapeStandIn", n: standIn });
  if (combo) out.push({ key: "xlsChartCombo" });
  if (errBars) out.push({ key: "xlsChartErrorBars" });
  return out;
}

/** Whether a shape geometry is written to a .xls as itself. */
export { xlsCanWrite } from "./drawing";

export { xlsPasswordState } from "./read";
