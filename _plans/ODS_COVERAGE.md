# ODS (ODF spreadsheet) coverage

sheetedit reads and writes .ods via surgical edits to `content.xml` (untouched cells/sheets are
cloned verbatim). This tracks the ODF-native authoring features and the known spec gaps, audited
against the OASIS ODF 1.3 spec.

## Supported

- Cells, formulas (ODF `of:` <-> A1), number/date/time/percentage/currency types, styles
  (fonts/fills/borders/alignment/wrap), merges, frozen panes, column/row sizes, furigana (ruby).
- Charts (read + author).
- Hyperlinks (`<text:a>`), notes (`<office:annotation>`), data validation
  (`<table:content-validation>`: list, whole/decimal/date/time, text-length and formula rules), and
  conditional formatting (every rule kind reads; see below for which half of the file states it).
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

**The catch, and it is the whole story for LibreOffice files.** A LibreOffice-written file puts the
`<style:map>` on an automatic style that it then applies to NO CELL, so the standard half states the
rule but never says which cells it covers: the only such statement is calcext's
`target-range-address`. Reading the interoperable half alone (and skipping calcext as "a mere
mirror") therefore found nothing at all in a real LibreOffice workbook. A file whose producer does
apply the conditional style to its cells still reads through `style:map`, which is why both paths
exist.

Also worth knowing: the two halves NAME THE APPLIED STYLE DIFFERENTLY. `style:map` uses the escaped
`style:name` and calcext the `style:display-name`, so `ConditionalStyle_5f_1` and
`ConditionalStyle_1` are one style and a lookup by either alone silently resolves nothing.

- **Read**: every rule kind, from whichever half states it. calcext conditions are parsed in full
  (`>20`, `between(15,35)`, `contains-text("ap")` / `not-contains-text` / `begins-with` /
  `ends-with`, `formula-is(...)`, `duplicate` / `unique`, `top-elements(n)` / `bottom-percent(n)`,
  `above-average` / `below-average`), plus colour scales / data bars / icon sets. `<style:map>`
  gives value comparisons, `cell-content-is-between` and `is-true-formula`. A rule stated in both
  halves is counted once. The grammar came from LibreOffice's own xlsx to ods conversion of one
  Excel rule of each kind, not from a guess.
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

- **Data validation**: list rules drive a dropdown; whole-number, decimal, date, time, text-length
  and is-true-formula rules read, author and drive the invalid-value outline, with every operator
  (between / not-between / the comparisons). A condition in neither of those shapes is preserved
  (the `<table:content-validations>` block is untouched and the cell's content-validation-name is
  re-emitted) but surfaces no UI. Authoring does not add error/help messages.
- **Hyperlinks**: read the first `<text:a>`; author a whole-cell anchor (verified to survive
  LibreOffice, which requires the linked string cell to omit `office:string-value`). The anchor's
  other attributes (target-frame-name, show, style-name, visited-style-name) are carried across an
  edit. STILL A GAP: a link on part of a cell's text, or several links in one cell, is preserved
  while the cell is untouched but flattens to one whole-cell link once you edit that cell's value
  or its link, since the text it was anchored to is rebuilt.
- **Comments**: every annotation on a cell is kept through an edit, each with its own position,
  creator and date; the grid shows and edits the first, and removing it leaves the others. A note's
  lines are written one `<text:p>` each. `dc:date` is written.
  CAVEAT: LibreOffice itself models ONE note per cell and keeps the last on re-save, so a second
  annotation survives sheetedit but not a pass through LibreOffice.
- **CF colour scale / data bar / icon set for ODS**: read-only (from LibreOffice files); not
  authorable (no interoperable ODF form).
- **CF authoring** stays on the interoperable `<style:map>` cellIs subset. Reading now covers every
  rule kind (see above), but authoring a text or graphical rule would have to write calcext, which
  LibreOffice drops when it did not write it itself.

## Pivot tables (both formats)

Pivot tables are detected, modelled, surfaced, preserved and (v1) authorable, in either format:

- **Authoring**: the Insert-pivot dialog (source range from the selection, assign each column to
  Rows / Columns / Values(+function) / Report Filter, optional subtotals, live preview) builds the
  pivot on a new sheet. Supports any number of nested row and column fields, one or more value
  fields (sum / count / average / min / max), per-group subtotals, and report/page filters. An
  authored pivot can be refreshed or edited in place from its overlay tag menu. It emits
  the native definition
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
