// Legacy Excel .xls (BIFF8, plus BIFF2-7 on read): the workbook is converted to xlsx parts on open,
// so every xlsx editing path applies, and written back as a BIFF8 compound file on save.
import type { Workbook } from "../../core/model";
import { readXlsx } from "../xlsx/read";
import { isCfb } from "./cfb";
import { isRawBiff, readXlsBook, type XlsReadOptions } from "./read";
import { xlsBookToXlsx } from "./to-xlsx";

export { XlsPasswordError, type XlsReadOptions } from "./read";

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
}

/** Whether these bytes look like a legacy .xls (compound file or raw BIFF stream). */
export const looksLikeXls = (bytes: Uint8Array): boolean => isCfb(bytes) || isRawBiff(bytes);

/** Read a .xls into a workbook the editor treats as xlsx, flagged with its BIFF origin. */
export function readXls(bytes: Uint8Array, opts: XlsReadOptions & { lazySheets?: boolean } = {}): Workbook {
  const book = readXlsBook(bytes, opts);
  const files = xlsBookToXlsx(book);
  const wb = readXlsx(files, { lazySheets: opts.lazySheets });
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
  };
  return wb;
}

/** What saving this .xls cannot keep: embedded charts, drawing shapes and skipped sheets. */
export function xlsLosses(wb: Workbook): { charts: number; shapes: number; sheets: string[] } | null {
  const x = wb.xls;
  if (!x || (!x.charts && !x.shapes && !x.skippedSheets.length)) return null;
  return { charts: x.charts, shapes: x.shapes, sheets: x.skippedSheets };
}

export { xlsPasswordState } from "./read";
