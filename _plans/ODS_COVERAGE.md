# ODS (ODF spreadsheet) coverage

sheetedit reads and writes .ods via surgical edits to `content.xml` (untouched cells/sheets are
cloned verbatim). This tracks the ODF-native authoring features and the known spec gaps, audited
against the OASIS ODF 1.3 spec.

## Supported

- Cells, formulas (ODF `of:` <-> A1), number/date/time/percentage/currency types, styles
  (fonts/fills/borders/alignment/wrap), merges, frozen panes, column/row sizes, furigana (ruby).
- Charts (read + author).
- Hyperlinks (`<text:a>`), notes (`<office:annotation>`), list data validation
  (`<table:content-validation>` cell-content-is-in-list), and conditional formatting.
- Embedded images (`<draw:frame>`/`<draw:image>` from `Pictures/` or inline base64), rendered on
  the overlay layer and preserved verbatim.
- Per-run rich text in a cell (`<text:span>` styles resolved to bold/italic/colour/etc.).
- LibreOffice sparklines (`<calcext:sparkline-groups>`) rendered; render-only.
- AutoFilter (`<table:database-range>` with filter buttons): toggle, per-column filter menu, sort,
  and row hiding persisted as `table:visibility`. Verified through a LibreOffice round-trip.
- Data-pilot (pivot) tables (`<table:data-pilot-table>`): detected and modelled (row/column/data/
  page fields, source range, output range), outlined read-only on the grid, and preserved verbatim.

## Conditional formatting (how ODF represents it)

Standard ODF value-based CF is `<style:map style:condition style:apply-style-name
style:base-cell-address>` on a cell's base style; `calcext:` is a LibreOffice-private extension
that additionally stores colour scales / data bars / icon sets (no standard form exists) and
mirrors the value conditions. LibreOffice imports value CF from `style:map`; a calcext-only file is
ignored by it.

- **Read**: value conditions from `<style:map>` (the interoperable path: `cell-content()` compares
  and `cell-content-is-between`); colour scales / data bars / icon sets from calcext. The calcext
  value-condition mirror is skipped to avoid duplicating the style:map rule.
- **Author**: highlight (cellIs) is written as a standard `<style:map>` + an applied fill style,
  which LibreOffice and other ODF consumers honour (verified: survives a LibreOffice round-trip).
  Colour-scale / data-bar / contains-text authoring are NOT offered for ODS (calcext-only, which
  LibreOffice drops when externally authored); rendering them from LibreOffice files still works.

## Pass-through preservation

Authoring a link / note / validation, or editing a value, patches a clone of the original cell
element (see `patchOdsCell`), so anything not explicitly changed survives: note position/size and
formatting, a validation's condition + error/help messages, unmodelled cell structure. Only the
aspect you change is rewritten. Verified against a LibreOffice round-trip.

## Known gaps (audited vs ODF 1.3)

- **Data validation**: only `cell-content-is-in-list` drives a dropdown / is authorable. Other
  condition types (is-between, text-length, is-true-formula, ...) are preserved (the
  `<table:content-validations>` block is untouched and the cell's content-validation-name is
  re-emitted) but not surfaced in the UI, and authoring a new list does not add error/help messages.
- **Hyperlinks**: read the first `<text:a>`; author a whole-cell anchor (verified to survive
  LibreOffice, which requires the linked string cell to omit `office:string-value`). A link on
  part of a cell's text, or multiple links, is preserved while the cell is untouched but flattened
  to one whole-cell link if you edit that cell's value or link. Optional `text:a` attributes
  (target-frame, show, style-name) are not preserved on an explicit edit.
- **Comments**: note position/formatting and a 2nd+ annotation are preserved while untouched, and
  are kept when you edit the cell's value; editing the note itself keeps the first annotation's
  position/creator/date but drops extra annotations (single-note model). `dc:date` is written.
- **CF condition grammar**: `is-true-formula(...)` and text-based conditions are not rendered.
- **CF colour scale / data bar / icon set for ODS**: read-only (from LibreOffice files); not
  authorable (no interoperable ODF form).

## Pivot tables (both formats)

Pivot tables are detected, modelled, surfaced, preserved and (v1) authorable, in either format:

- **Authoring**: the Insert-pivot dialog (source range from the selection, assign each column to
  Rows / Columns / Values with a function, live preview) builds the pivot on a new sheet. v1
  supports any number of nested row fields, at most one column field, one or more value fields
  (sum / count / average / min / max), no intermediate subtotals. It emits the native definition
  (xlsx `pivotCache`+`pivotTable`, ODS `data-pilot-table`) with `refreshOnLoad` so Excel/LibreOffice
  rebuild from the source on open, plus the materialised output cells. Verified end-to-end through a
  LibreOffice round-trip (the recomputed values match the engine).
- **ODS** reads `<table:data-pilot-table>`; **xlsx** reads `pivotTable*.xml` + its `pivotCache`.
  Each pivot's output already renders as the cells it materialises; on top of that the output range
  is outlined and labelled read-only so the region reads as a live pivot, not editable cells.
- **Preservation**: every pivot part passes through the writer verbatim (xlsx cache/table parts;
  the ODS `data-pilot-tables` block rides along in the untouched content.xml subtree).
- **Source-edit refresh (xlsx)**: editing a cell inside a pivot's worksheet source sets
  `refreshOnLoad="1"` on that cache definition, so **Excel** rebuilds the pivot from the changed
  source on open instead of showing the stale cached output. (LibreOffice recomputes data pilots on
  open regardless, so the flag is a no-op there, harmless.) The pivot's own cached records are not
  recomputed in-app; Excel/LibreOffice do that on load.
- **Source-edit refresh (ODS)**: ODF data pilots carry no refresh-on-load flag and LibreOffice does
  not auto-refresh them on open, so an edited ODS source leaves the pilot output stale until it is
  refreshed in LibreOffice, matching LibreOffice's own behaviour when you edit pivot source.

## Not applicable to ODF

- **Power Query**: ODF has no equivalent to Power Query / the M language / DataMashup. The nearest
  concepts (database ranges, pivot tables, DDE/external-data links) do not store a portable
  transformation pipeline. A Power-Query workbook saved as .ods keeps only the cached result values.
