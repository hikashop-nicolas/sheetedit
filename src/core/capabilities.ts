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
};

// ODF has native hyperlinks, comments, validation and conditional formatting, but no Excel-style
// sparklines or the x14 autofilter model. Only charts are wired for ODS so far; the ODF authoring
// paths flip the remaining flags on as they land.
const ODS: WorkbookCapabilities = {
  charts: true,
  sparklines: false,
  autofilter: true,
  hyperlinks: true,
  dataValidation: true,
  comments: true,
  conditionalFormat: true,
  pivots: true,
};

/** The capabilities of a workbook format. CSV (and anything else) supports none of the above. */
export function capabilitiesFor(kind: Workbook["kind"]): WorkbookCapabilities {
  return kind === "xlsx" ? XLSX : kind === "ods" ? ODS : NONE;
}
