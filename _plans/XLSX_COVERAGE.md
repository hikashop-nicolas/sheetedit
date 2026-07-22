# sheetedit xlsx coverage report

Audit of what the `.xlsx` path reads, edits and preserves, as of 2026-07-23 (Power Query
included). Tracks the gap list and the agreed priority order for closing it.

## Fully supported (read, edit in the grid, save)

- Cell content: text, number, boolean, error; inline and shared strings.
- Formulas + recalc: a large fast-formula-parser subset, recomputed in dependency order across
  sheets; legacy array formulas preserved; shared-formula groups de-shared safely on edit.
- Number formats: resolved from styles.xml and applied to the display (date, currency, percent,
  thousands); a typed value keeps the cell's format.
- Cell styling: bold, italic, underline (+ flavour), strikethrough, font family/size, text
  colour, fill colour, horizontal and vertical alignment, wrap, per-side borders.
- Furigana (Japanese phonetic ruby).
- Structure: insert/delete rows and columns (with formula and merge reference rewriting),
  column-width and row-height resize, merge/unmerge.
- View: frozen panes (sticky), hidden rows/columns.
- Multiple sheets (switch via tabs), find and replace across sheets, fill down/right, on-device
  formula assistant.
- Power Query: read, full editor (Applied Steps, preview, transform ribbon, Get Data,
  merge/append), refresh, and Load (existing table or new sheet).

## Preserved on save, but inert in the grid (round-trips, not rendered or editable)

Survive because untouched parts are kept byte-for-byte and the worksheet DOM is re-serialized
with its sibling elements intact:

- Charts, images, shapes, drawings
- Pivot tables and caches (not refreshed)
- Conditional formatting (not applied to the displayed cells)
- Data validation / dropdowns (rules kept, no dropdown UI)
- Hyperlinks (not shown or clickable)
- Comments, notes, threaded comments
- Sparklines, form controls, slicers
- Defined names (read for recalc, not user-editable), sheet/workbook protection, print
  settings, autofilter state, outline grouping, themes

## Gaps / not handled

- No sheet management UI: cannot add, rename, delete, reorder, hide, or colour sheets from the
  grid. A sheet is only created programmatically when Power Query loads to a new destination.
- No dynamic-array / spill computation; only legacy array formulas are preserved (not re-spilled).
- No editing of the preserved-only features above.
- Secondary ranges are not shifted on insert/delete: formula refs and merges are rewritten, but
  conditional-formatting, data-validation and hyperlink ranges are not, so they can go stale.
- No rich text within one cell: a multi-format string renders as plain text with a single style.
- Recalc is a subset: unsupported functions or circular refs yield an error value (cached value
  shown as fallback); exotic custom number-format codes may render slightly differently (SSF).

## Correctness caveats

- CF/DV/hyperlink range staleness after structural edits.
- Editing a shared-string cell rewrites it as an inline string (its sharedStrings entry can
  become unreferenced).
- Power Query load-to-new-sheet writes plain cells, not a live refreshable ListObject, and is
  verified only through sheetedit's own reader, not real Excel.

## Priority order for closing the gaps

1. Sheet management UI (add/rename/delete/reorder) - high value, createWorksheet already exists.
2. Hyperlinks (render + click) - common, low effort, data already in the sheet XML.
3. Data validation dropdowns (render + enforce) - common, moderate effort.
4. Shift CF/DV/hyperlink ranges on structural edits - a correctness fix.
5. Conditional formatting rendering - high visual payoff, higher effort.
6. Comments display - moderate.

## Progress

- (DONE) 1. Sheet management UI - add/rename/delete/reorder tabs, xlsx + ods (sheet-ops.ts)
- (DONE) 2. Hyperlinks - read external+internal, render blue link + open button, click opens/navigates
- (DONE) 3. Data validation dropdowns - list type: caret picker (inline + range values) + invalid-value outline
- (DONE) 4. Shift CF/DV/hyperlink/autofilter ranges (sqref/ref) + model validations on insert/delete
- (DONE) 5. Conditional formatting rendering - cellIs/text/top-bottom/average/dup dxf + colour scales + data bars (icon sets, expression, time-period round-trip only)
- (DONE) 6. Comments display - legacy + threaded comments: corner marker + hover popover (author + text)
