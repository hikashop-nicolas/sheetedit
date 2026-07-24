import type { Workbook } from "./model";

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
};

/** The capabilities of a workbook format. CSV (and anything else) supports none of the above. */
export function capabilitiesFor(kind: Workbook["kind"]): WorkbookCapabilities {
  return kind === "xlsx" ? XLSX : kind === "ods" ? ODS : NONE;
}
