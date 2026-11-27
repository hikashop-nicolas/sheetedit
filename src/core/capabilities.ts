import { MAX_COL, MAX_ROW, type Workbook } from "./model";

// What a given workbook format supports, so the UI advertises exactly the features that work for
// the open file instead of scattering `wb.kind === "xlsx"` checks. Each authoring control is gated
// on its capability; a format enabling a feature (e.g. ODS gaining hyperlinks) flips one flag here
// and the button appears, with the write path dispatched per-format at the call site.
export interface WorkbookCapabilities {
  /** Insert/edit charts over the grid. */
  charts: boolean;
  /** In-cell sparklines (author + render). */
  sparklines: boolean;
  /** AutoFilter toggle + per-column filter menu. */
  autofilter: boolean;
  /** Cell hyperlinks (author + follow). */
  hyperlinks: boolean;
  /** List data-validation rules. */
  dataValidation: boolean;
  /** Cell notes / comments. */
  comments: boolean;
  /** Conditional formatting (author + render). */
  conditionalFormat: boolean;
  /** Create a pivot table / data-pilot from a source range. */
  pivots: boolean;
  /** Draw shapes (rect / ellipse / line / ...) over the grid. */
  shapes: boolean;
  /** Row / column outline grouping (gutter + group / ungroup / collapse). */
  outline: boolean;
  /** Freeze panes (author + write back the view settings). */
  freezePanes: boolean;
  /** Sheet / workbook protection (enforce, author, write back). */
  protection: boolean;
  /** Page setup for printing (author + write back; the grid shows breaks and the print area). */
  printSetup: boolean;
  /** The workbook's own theme: its colour palette and heading / body fonts. */
  themes: boolean;
  /** Form controls (checkbox, dropdown, spinner) rendered and driving their linked cell. */
  formControls: boolean;
  /** Defined names: the Name Manager and defining or jumping to a name from the name box. */
  definedNames: boolean;
  /** Named tables (xlsx ListObjects, ODF database ranges): create / edit / delete. */
  tables: boolean;
  /** Pictures: insert and delete (moving and resizing existing ones is always allowed). */
  images: boolean;
  /** The conditional-formatting kinds beyond value / formula / text / date rules: colour scales,
      data bars, icon sets, top-N, above average and duplicate / unique. */
  conditionalFormatAdvanced: boolean;
  /** Power Query: the query panel and editor, whose definitions live in an xlsx DataMashup part. */
  powerQuery: boolean;
  /** The grid's size: rows and columns the format can hold. */
  maxRows: number;
  maxCols: number;
}

const NONE: WorkbookCapabilities = {
  charts: false,
  sparklines: false,
  autofilter: false,
  hyperlinks: false,
  dataValidation: false,
  comments: false,
  conditionalFormat: false,
  pivots: false,
  shapes: false,
  outline: false,
  freezePanes: false,
  protection: false,
  printSetup: false,
  themes: false,
  formControls: false,
  definedNames: false,
  tables: false,
  images: false,
  conditionalFormatAdvanced: false,
  powerQuery: false,
  maxRows: MAX_ROW,
  maxCols: MAX_COL,
};

const XLSX: WorkbookCapabilities = {
  charts: true,
  sparklines: true,
  autofilter: true,
  hyperlinks: true,
  dataValidation: true,
  comments: true,
  conditionalFormat: true,
  pivots: true,
  shapes: true,
  outline: true,
  freezePanes: true,
  protection: true,
  printSetup: true,
  themes: true,
  formControls: true,
  definedNames: true,
  tables: true,
  images: true,
  conditionalFormatAdvanced: true,
  powerQuery: true,
  maxRows: MAX_ROW,
  maxCols: MAX_COL,
};

// ODF has native hyperlinks, comments, validation, conditional formatting, charts and (via the
// LibreOffice calcext extension) sparklines. The x14 autofilter model differs but is wired.
const ODS: WorkbookCapabilities = {
  charts: true,
  sparklines: true,
  autofilter: true,
  hyperlinks: true,
  dataValidation: true,
  comments: true,
  conditionalFormat: true,
  pivots: true,
  shapes: true,
  // ODF groups ROWS with table:table-row-group; column groups are read and preserved but a
  // rebuilt column run is not re-nested, so authoring them is xlsx-only.
  outline: true,
  // ODF keeps frozen panes in settings.xml (the view settings), not in the table itself.
  freezePanes: true,
  // ODF has table:protected + the loext permission flags; it lacks the format / sort / filter /
  // pivot granularity, so those keep their (blocked) defaults on an ods round-trip.
  protection: true,
  // ODF states page setup as a page-layout + master-page pair in styles.xml, reached through the
  // table style; print ranges, repeated rows and page breaks live in content.xml.
  printSetup: true,
  // ODF has no theme palette that cells reference by index, so there is nothing to switch.
  themes: false,
  // ODF keeps controls in office:forms with a draw:control frame, a different model that is
  // preserved rather than rendered.
  formControls: false,
  definedNames: true,
  tables: true,
  images: true,
  conditionalFormatAdvanced: true,
  powerQuery: false,
  maxRows: MAX_ROW,
  maxCols: MAX_COL,
};

// A legacy .xls is edited through the xlsx paths but saved as BIFF8, which has no tables, slicers,
// sparklines, themes, pivot caches the editor can author, drawing shapes or charts the writer
// emits, only value / formula conditional formats, and a 65,536 x 256 grid.
const XLS: WorkbookCapabilities = {
  ...XLSX,
  charts: false,
  sparklines: false,
  pivots: false,
  shapes: false,
  themes: false,
  formControls: false,
  tables: false,
  conditionalFormatAdvanced: false,
  powerQuery: false,
  maxRows: 65536,
  maxCols: 256,
};

/** The capabilities of a workbook format. CSV (and anything else) supports none of the above. */
export function capabilitiesFor(kind: Workbook["kind"] | "xls"): WorkbookCapabilities {
  return kind === "xls" ? XLS : kind === "xlsx" ? XLSX : kind === "ods" ? ODS : NONE;
}

/** The capabilities of an open workbook: its kind, or BIFF8's when it came from a .xls. */
export const capabilitiesOf = (wb: Workbook): WorkbookCapabilities => capabilitiesFor(wb.xls ? "xls" : wb.kind);
