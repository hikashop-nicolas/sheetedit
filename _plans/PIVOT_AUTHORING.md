# Pivot authoring plan

Status: v1 shipped. Builds on the read-only pivot support (detect / outline / preserve / refresh).
Adds creating a pivot table from a source range, in both xlsx and ODS, verified through LibreOffice
round-trips. The dialog has a live preview and guards the "no data selected" case.

## Scope (v1)

- **Fields**: any number of row fields (nested), any number of column fields (nested), one or more
  value fields each with an aggregation. Page/report filters deferred.
- **Aggregations**: sum, count (non-empty), countNums, average, min, max.
- **No intermediate subtotals** (`defaultSubtotal="0"`, exactly as LibreOffice emits): only leaf
  group combinations plus grand totals. Subtotals can be a follow-up.
- **Output**: a new sheet by default (Excel's default; avoids relayout overlap), or a chosen anchor.

## Architecture

- `src/core/pivot.ts` (pure, unit-tested): given the source cells + a `PivotSpec`
  (row/col/value field indices + funcs), compute distinct sorted items per field, the occurring
  row-key / col-key tuples, the aggregated values with grand totals, a materialised output matrix,
  and the raw records. Format-agnostic; the single source of truth both writers consume.
- **xlsx** (`pivot-write.ts`): emit `pivotCacheDefinition` (cacheFields + sharedItems for grouping
  fields, type flags for others), `pivotCacheRecords`, `pivotTableDefinition` (pivotFields,
  rowFields/rowItems with the delta encoding `@r` = unchanged-prefix length, colFields/colItems,
  dataFields, location), the rels + `[Content_Types].xml` overrides + workbook `<pivotCaches>`, and
  the materialised output cells. `refreshOnLoad="1"` so the app rebuilds layout on open.
- **ODS** (`write.ts`): emit `<table:data-pilot-table>` (source-cell-range + one
  `<table:data-pilot-field>` per orientation with its function) plus the materialised output cells.
- **UI** (`pivot-ui.ts`): an insert dialog (source range prefilled from selection, assign each
  source column to Rows / Columns / Values(+func) / unused, pick output). Gated on a new
  `pivots` capability (xlsx + ods true).

## Verification

- Engine unit tests (single + multi field, multi value, each function, grand totals).
- Writer round-trip: author, write, reopen through LibreOffice, assert the recomputed values match
  the engine (the reference generator is LibreOffice's own pivots, so the emitted XML mirrors a
  known-good structure). Excel is the one consumer we cannot run; the XML mirrors LibreOffice's
  Excel-compatible output and sets refreshOnLoad, but complex layouts should be sanity-checked there.
- e2e: author a pivot through the dialog and assert the output cells + the read-back model.
