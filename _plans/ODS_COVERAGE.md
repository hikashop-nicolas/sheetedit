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

## Not applicable to ODF

- **Power Query**: ODF has no equivalent to Power Query / the M language / DataMashup. The nearest
  concepts (database ranges, pivot tables, DDE/external-data links) do not store a portable
  transformation pipeline. A Power-Query workbook saved as .ods keeps only the cached result values.
