// New, empty workbooks for a host's "New spreadsheet" command, written with sheetedit's own writers.
import { strToU8 } from "fflate";
import type { Workbook } from "./model";
import { t } from "./i18n";
import { writeWorkbook } from "./workbook";
import { readOds } from "../adapters/ods";
import { readXlsx } from "../adapters/xlsx";
import { newSheet } from "../adapters/tabular/model";
import { tBookToXlsx } from "../adapters/tabular/to-xlsx";
import { setSaveFormat } from "../adapters/tabular";

export type BlankFormat = "xlsx" | "xls" | "ods";

export interface BlankOptions {
  /** The one sheet's name; defaults to the locale's "Sheet1". */
  sheetName?: string;
}

const esc = (s: string): string => s.replace(/[&<>"]/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : "&quot;"));
const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"';
const DECL = '<?xml version="1.0" encoding="UTF-8"?>\n';

function blankOdsFiles(name: string): Record<string, Uint8Array> {
  const content = DECL + `<office:document-content ${NS} office:version="1.3"><office:automatic-styles/><office:body><office:spreadsheet>` +
    `<table:table table:name="${esc(name)}"><table:table-column/><table:table-row><table:table-cell/></table:table-row></table:table>` +
    "</office:spreadsheet></office:body></office:document-content>";
  const styles = DECL + `<office:document-styles ${NS} office:version="1.3"><office:styles><style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="1.25cm"/></style:default-style></office:styles></office:document-styles>`;
  const meta = DECL + '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>sheetedit</meta:generator></office:meta></office:document-meta>';
  const manifest = DECL + '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">' +
    '<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>' +
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>';
  return {
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(content),
    "styles.xml": strToU8(styles),
    "meta.xml": strToU8(meta),
    "META-INF/manifest.xml": strToU8(manifest),
  };
}

/** A new workbook with one empty sheet, ready for the editor; writeWorkbook saves it in `format`. */
export function createWorkbook(format: BlankFormat = "xlsx", opts: BlankOptions = {}): Workbook {
  const name = (opts.sheetName ?? t("sheetDefaultName", { n: 1 })).replace(/[[\]:*?/\\]/g, " ").trim().slice(0, 31) || "Sheet1";
  if (format === "ods") return readOds(blankOdsFiles(name));
  const wb = readXlsx(tBookToXlsx({ sheets: [newSheet(name)], names: [] }));
  if (format === "xls") setSaveFormat(wb, "xls");
  return wb;
}

/** The bytes of a new, empty workbook as .xlsx, .xls (BIFF8) or .ods. */
export function blankWorkbookBytes(format: BlankFormat = "xlsx", opts: BlankOptions = {}): Uint8Array {
  return writeWorkbook(createWorkbook(format, opts));
}
