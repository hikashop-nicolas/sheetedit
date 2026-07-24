# sheetedit xlsx coverage report

Audit of what the `.xlsx` path reads, edits and preserves, as of 2026-07-24 (Power Query, charts
and pivot tables included). Tracks the gap list and the agreed priority order for closing it.

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
  merge/append), refresh, and Load (existing table or new sheet); can now author the first query
  in a query-less workbook (bootstraps the DataMashup payload).
- Charts: read, render (Chart.js + custom plugins), full create/edit dialog and write for every
  DrawingML chart type, all option tiers (axes, labels, trendlines, error bars, stock, of-pie,
  surface, styling), pseudo-3D, xlsx + ods round-trip.
- Pivot tables: read, outline on the grid, and create/edit/refresh. The insert dialog assigns
  columns to Rows / Columns / Values(+function) / Report Filter with subtotals and a live preview;
  nested row/column fields, multiple value fields (sum/count/average/min/max) and page filters are
  supported. Emits the native pivotCache + pivotTable parts with refreshOnLoad; edit/refresh work
  in place, for authored and file-read pivots alike (the spec is reconstructed on read). Verified
  through LibreOffice round-trips (xlsx + ods).

## Preserved on save, but inert in the grid (round-trips, not rendered or editable)

Survive because untouched parts are kept byte-for-byte and the worksheet DOM is re-serialized
with its sibling elements intact:

- Shapes and other drawings (pictures can now be moved/resized, and charts and pivot tables are
  fully editable, see above)
- Form controls, slicers, ActiveX
- Defined names (read for recalc, not user-editable), sheet/workbook protection, print
  settings, autofilter state, outline grouping, themes

Now rendered (were inert): hyperlinks (click), data-validation dropdowns, conditional
formatting (dxf / colour scales / data bars), comments and notes. See Progress below.

## Gaps / not handled

- Dynamic-array spill (MVP): a plain formula returning a 2-D result spills into its anchor + range
  (with a #SPILL! guard on collisions). Producers UNIQUE / SORT / FILTER / SEQUENCE are supplied
  (fast-formula-parser ships none); TRANSPOSE and bare range refs spill too. FILTER needs an array
  mask (no range=scalar broadcasting); multi-key SORT and exotic producers are best-effort.
- Pictures can be moved, resized and replaced (xlsx and ods; the anchor/frame is written back and
  the media part swapped). Editing non-picture shapes is still not supported (shapes are not
  rendered). Other preserved-only features above (form controls, slicers, protection) are not
  editable. The now-rendered features (hyperlinks, dropdowns, CF, comments) are read/followed, not
  authored.
- Pivot tables: nested fields, filters, subtotals, "show values as" (% of total/row/col + running
  total), calculated fields, calculated items, and a pivot chart (over the output) are all
  supported. Not attempted: byte-identical layout to Excel (both apps re-flow the body on open).
  Show-values-as / calculated fields / calculated items are honoured by Excel and sheetedit's
  display but ignored by LibreOffice's xlsx pivot rebuild; the calculated-item OOXML is emitted per
  spec but unverified in Excel (only that the file opens cleanly in LibreOffice).
- Conditional formatting: full authoring and rendering. Icon sets render (arrows/traffic-lights/
  symbols/ratings, bucketed by the cfvo thresholds); is-true-formula (expression) rules and cellIs
  cell-ref/formula operands evaluate through the workbook's formula engine; time-period rules (dates
  occurring) author + render against today. The authoring dialog covers every rule kind; ODS keeps
  the interoperable subset (cellIs incl. between) since the graphical/formula rules have no ODF form
  that survives a LibreOffice round-trip.
- Rich text within one cell renders each run's own style (bold/italic/colour/etc.) and is authored
  from the UI: select part of a cell's text while editing and apply a style to just that run. Runs
  are written to xlsx (`<r><rPr>`) and ods (`<text:span>`) and survive a LibreOffice round-trip.
  Retyping a cell's text clears its runs (offsets would shift).
- Recalc is a subset: unsupported functions or circular refs yield an error value (cached value
  shown as fallback); exotic custom number-format codes may render slightly differently (SSF).

## Correctness caveats

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
- (DONE) 5. Conditional formatting - full author + render: cellIs/text/top-bottom/average/dup dxf + colour scales + data bars + icon sets + is-true-formula + cell-ref/formula operands + time-period rules
- (DONE) 6. Comments display - legacy + threaded comments: corner marker + hover popover (author + text)
