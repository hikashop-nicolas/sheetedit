// Spreadsheet exports that are text: HTML tables (Excel's "Save as Web Page", bank and web exports,
// often named .xls) and Excel 2003 XML Spreadsheets. They open as xlsx-backed workbooks; the source
// format is never written back, the host picks .xls (BIFF8) or .xlsx with setSaveFormat.
import type { Workbook } from "../../core/model";
import { t } from "../../core/i18n";
import { readHtmlBook } from "../html/read";
import { readSpreadsheetMlBook } from "../spreadsheetml/read";
import { readXlsx } from "../xlsx/read";
import { recalc } from "../../core/recalc";
import { DEFAULT_PALETTE } from "../xls/palette";
import type { TBook } from "./model";
import { decodeSheetText, sniffSheetText, type SheetTextKind } from "./text";
import { tBookToXlsx } from "./to-xlsx";

export { sniffSheetText, decodeSheetText, type SheetTextKind } from "./text";

export interface SheetTextReadOptions {
  /** Parse with DOMParser when available (default true); false forces the built-in tokenizer. */
  useDom?: boolean;
}

const sheetName = (n: number): string => t("sheetDefaultName", { n });

function toWorkbook(book: TBook, from: SheetTextKind): Workbook {
  const wb = readXlsx(tBookToXlsx(book));
  wb.importedFrom = from;
  // Exporters other than Excel often leave formula results out or stale, so compute them on open.
  if (wb.sheets.some((s) => [...s.cells.values()].some((c) => c.formula != null))) recalc(wb);
  return wb;
}

const asText = (src: Uint8Array | string): string => (typeof src === "string" ? src : decodeSheetText(src));

/** Read an HTML table export (or Excel's single-file web page) into a workbook. */
export function readHtmlWorkbook(src: Uint8Array | string, opts: SheetTextReadOptions = {}): Workbook {
  return toWorkbook(readHtmlBook(asText(src), { useDom: opts.useDom, sheetName }), "html");
}

/** Read an Excel 2003 XML Spreadsheet into a workbook. */
export function readSpreadsheetMl(src: Uint8Array | string, opts: SheetTextReadOptions = {}): Workbook {
  return toWorkbook(readSpreadsheetMlBook(asText(src), { useDom: opts.useDom, sheetName }), "spreadsheetml");
}

/** Read either kind, sniffing the content when `kind` is not given; null when it is neither. */
export function readSheetText(bytes: Uint8Array, kind?: SheetTextKind | null, opts: SheetTextReadOptions = {}): Workbook | null {
  const k = kind ?? sniffSheetText(bytes, "xls");
  if (k === "html") return readHtmlWorkbook(bytes, opts);
  if (k === "spreadsheetml") return readSpreadsheetMl(bytes, opts);
  return null;
}

/** The format writeWorkbook (and the editor's getBytes) produces for this workbook. */
export type SaveFormat = "xlsx" | "xls" | "ods" | "csv";

export function saveFormat(wb: Workbook): SaveFormat {
  if (wb.kind === "xlsx") return wb.xls ? "xls" : "xlsx";
  return wb.kind;
}

/**
 * Choose what an xlsx-backed workbook saves as: "xls" writes Excel 97-2003 BIFF8 (writeXls), "xlsx"
 * writes Office Open XML. Meant for workbooks read from HTML or XML exports and for new ones.
 */
export function setSaveFormat(wb: Workbook, format: "xls" | "xlsx"): void {
  if (wb.kind !== "xlsx") throw new Error(`a ${wb.kind} workbook cannot be saved as ${format}`);
  if (format === "xlsx") { delete wb.xls; return; }
  wb.xls ??= { biff: 8, palette: DEFAULT_PALETTE.slice(), externals: [], skippedSheets: [], charts: 0, chartDrawings: 0, shapes: 0, controls: 0, udfs: [] };
}

/** A one-time notice for a workbook read from an HTML or XML export, naming what it is saved as. */
export function importedSaveNotice(wb: Workbook): string | null {
  if (!wb.importedFrom) return null;
  const format = wb.xls ? "Excel 97-2003 (.xls)" : "Excel (.xlsx)";
  return t(wb.importedFrom === "html" ? "importedHtmlNotice" : "importedXmlNotice", { format });
}
