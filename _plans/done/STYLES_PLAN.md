# sheetedit cell styles: display + edit

## Problem

The editor renders a plain grid: it ignores the workbook's cell styling (fill
colours, fonts/bold, borders, text alignment) and column widths, so a styled
spreadsheet looks wrong. The cell style index (`Cell.style`, the xlsx `@s`) is
already read and preserved on save, but never resolved to actual formatting, and
there is no UI to change styles.

Two goals:
1. **Display** existing styles faithfully (read-only fidelity).
2. **Edit** styles via a toolbar (bold/italic, text colour, fill, borders,
   alignment, column width) and write them back into the file.

## Current state (what exists)

- `XlsxStyles` (src/index.ts ~198) parses only number formats: `customFmt`
  (numFmtId -> code) and `xfNumFmtIds` (cellXfs index -> numFmtId).
- `readSheetData` records each cell's `@s` into `Cell.style` and resolves only the
  number format.
- `renderGrid` (~1040) builds an HTML table of `<input>`s; fixed input width 96px;
  the only per-cell class is `.num`. No fonts/fills/borders/widths.
- `writeWorkbook` does surgical `<c>` edits and leaves `xl/styles.xml` untouched.
- ODS has its own model (`table:style-name`), also unstyled in the grid.

## The OOXML style model (what we must read/write for xlsx)

`xl/styles.xml`: `<fonts>`, `<fills>`, `<borders>`, `<cellXfs>` (+ numFmts). A cell's
`@s` indexes `<cellXfs>`; each `<xf>` references `fontId`/`fillId`/`borderId`/
`numFmtId` with `applyFont`/`applyFill`/`applyBorder`/`applyAlignment` flags and an
optional `<alignment>`.
- font: `<b/> <i/> <sz val> <color rgb="FFRRGGBB" | theme=n tint=x> <name val>`.
- fill: `<patternFill patternType="solid"><fgColor .../></patternFill>`.
- border: `<left|right|top|bottom style="thin|medium|..."><color/></...>`.
- Theme colours (`color theme="n" tint`) resolve against `xl/theme/theme1.xml`
  (`<a:clrScheme>`); needs a small theme reader + tint math, with the standard
  Office palette as a fallback.
- Column widths: worksheet `<cols><col min max width customWidth/></cols>` (width is
  in character units; convert to px). Row heights: `<row ht>`.

## Plan

### Phase 1 - Display xlsx styles (the reported bug)

- Extend `XlsxStyles` to also parse pools into resolved records:
  - fonts[]: { bold, italic, underline, color, sizePt, name }
  - fills[]: { color | null }
  - borders[]: { left, right, top, bottom: {style,color} | null }
  - xfs[]: { fontId, fillId, borderId, applyFont, applyFill, applyBorder, align }
- Add `xl/theme/theme1.xml` reader for theme colours + tint.
- Resolve per cell: `Cell.style` (s) -> xf -> a `CellStyle` { bold, italic, color,
  bg, borders, align } attached to the cell (computed at read time).
- Parse the worksheet `<cols>` into a per-column width map on the Sheet.
- `renderGrid`: apply CellStyle to each `<td>`/`<input>` (font-weight, font-style,
  color, background, border, text-align) and set column widths via a `<colgroup>`.
- Verify on the user's real workbook (colours, borders, bold, widths show).

### Phase 2 - Display ods styles

- Parse `content.xml`/`styles.xml` automatic styles: `table:style-name` ->
  `style:table-cell-properties` (background, border), `style:text-properties`
  (bold/colour), and `table:table-column` `style:column-width`. Map to the same
  CellStyle and column widths used by the renderer.

### Phase 3 - Toolbar to set styles + write-back

- Add a toolbar (bold, italic, text colour, fill colour, borders, horizontal
  align; column width via header drag or a number input). Selection model: a
  current cell plus a rectangular range.
- xlsx write-back = an OOXML style writer that manages the pools:
  - keep in-memory indexes of fonts/fills/borders/xfs;
  - applying a style derives the target xf from the cell's current xf plus the
    change, finds-or-creates the font/fill/border and the xf (deduped), and sets
    the cell `@s`;
  - on save, serialize the added pool entries into `styles.xml` (today it is
    preserved untouched; this becomes a careful, additive edit) and write changed
    `@s`. Column width changes edit/insert `<col>`; row height edits `<row ht>`.
- ods write-back = add automatic cell/column styles and set `table:style-name`.

## Risks / decisions

- Theme-colour + tint resolution is fiddly; Phase 1 handles rgb directly and maps
  theme indices via theme1.xml, falling back to the standard palette.
- The style writer must stay additive and dedupe pool entries, or files bloat and
  unrelated cells shift styling. This is the riskiest part; cover with round-trip
  tests (apply bold -> save -> reopen -> still bold, other cells unchanged).
- Range selection + a toolbar is a meaningful UI addition; keep it touch-friendly
  for the Android app (the toolbar should wrap/scroll on narrow screens).
- Scope order: ship Phase 1 (xlsx display) first to fix the visible regression,
  then ODS display, then the editing toolbar.

## What stays the same

Values, formulas + recalc, number-format display, in-place preservation of
untouched cells/sheets, and the public API (createSheetEditor / getBytes) are
unchanged. Styling is layered on: read adds resolved style info; the toolbar and
style writer are additive.
