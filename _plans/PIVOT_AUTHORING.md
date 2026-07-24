# Pivot authoring plan

Status: feature-complete for the common + advanced cases. Builds on the read-only pivot support
(detect / outline / preserve / refresh) and adds creating, editing and refreshing a pivot in both
xlsx and ODS, verified through LibreOffice round-trips. The dialog has a live preview and guards the
"no data selected" case.

Advanced value features (2026-07-24): **show values as** (% of grand/column/row total, running
total), **calculated fields** (a formula over the source columns' sums, via a small in-engine
arithmetic parser), and a **pivot chart** (chart the pivot output, grand totals excluded, from the
tag menu; bound to the output cells so it updates on refresh), plus **calculated items** (a
synthetic member of a row/column field from a formula over that field's items, e.g. "North +
South"; appears as an extra row/column, grand totals exclude it). xlsx emits showDataAs, calculated
cacheFields, and `<calculatedItems>` (+ `f="1"` items); all are reconstructed on read so they
survive an edit. Excel and sheetedit's display honour them; LibreOffice's xlsx pivot rebuild does
not. The calculated-item OOXML is emitted per the ECMA-376 spec but could not be verified in Excel
here (only that the file opens cleanly in LibreOffice, which ignores calc items). Pivots are now
feature-complete apart from byte-identical Excel layout (both apps re-flow the body on open anyway).

## Scope

- **Fields**: any number of nested row fields, any number of nested column fields, one or more value
  fields each with an aggregation, and report/page filters. Assigned per-column in the dialog with a
  live preview.
- **Aggregations**: sum, count (non-empty), average, min, max.
- **Subtotals**: optional per-group subtotals for the outer nested fields (toggle in the dialog);
  emitted as `<item t="default"/>` + `<i t="default">` lines (xlsx) / `data-pilot-subtotals` (ODS).
- **Report filters**: a field restricted to one value (or All); filters the aggregation and emits
  `axisPage` + `<pageFields>` (xlsx) / a page field with member display flags (ODS).
- **Output**: a new sheet (Excel's default; avoids relayout overlap).
- **Edit / refresh**: clicking a pivot's overlay tag opens a menu. Refresh recomputes the output
  from the current source; Edit reopens the dialog prefilled and rewrites the pivot in place (old
  definition parts removed, new ones emitted at the same anchor). This works for pivots created
  in-app AND for pivots read from a file: the reader reconstructs the authoring spec from the
  definition (field roles from the pivotFields/rowFields/colFields/dataFields/pageFields for xlsx,
  from the data-pilot-field orientations for ODS; functions, page selections and subtotals too).
  A pivot is only read-only when its spec can't be reconstructed (e.g. an external/non-worksheet
  cache source), in which case the menu says so.

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
