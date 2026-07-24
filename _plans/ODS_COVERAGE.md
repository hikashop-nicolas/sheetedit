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

## Known gaps (audited vs ODF 1.3)

- **Data validation**: only `cell-content-is-in-list` (dropdowns) is surfaced/authored. Other
  condition types (is-between, text-length, whole-number/date, is-true-formula) are not read, and
  authoring writes a minimal validation without `table:error-message` / `table:help-message` /
  `table:base-cell-address`. Untouched validated cells are preserved verbatim; only cells we
  re-emit lose non-list rules. Fix direction: pass the original validation elements through.
- **Hyperlinks**: we read the first `<text:a>` and author a whole-cell anchor. A cell with a link
  on only part of its text, or multiple links, is flattened to one whole-cell link on re-emit.
  Optional `text:a` attributes (target-frame, show, style-name) are not preserved.
- **Comments**: one note per cell; `<office:annotation-end>` (range comments) and 2nd+ annotations
  are dropped, and re-emitting a note loses its saved position/size and rich formatting. `dc:date`
  is not written.
- **CF condition grammar**: `is-true-formula(...)` and text-based conditions are not rendered.
- **CF colour scale / data bar / icon set for ODS**: read-only (from LibreOffice files); not
  authorable (no interoperable ODF form).

Overarching direction for the gaps: treat validations, hyperlink anchors, and annotations as
pass-through XML (parse for the UI, re-emit the original nodes when unchanged) instead of
reconstructing minimal versions.
