# sheetedit: the nine gap-closing phases

Shipped between 2026-06 and 2026-07-24, each a self-contained increment: implement, unit-test,
browser-verify, round-trip through LibreOffice where a file format was touched, then bump omnitext.
Kept for the reasoning rather than the status; the living record is `../SHEETEDIT_GAPS_PLAN.md`.

## Phases

### Phase 1 - Interactive sort & filter (DONE)
Autofilter state already round-trips; add the UI. Click a column header (or a filter caret on the
autofilter range) to sort asc/desc, and a dropdown to filter by value (checkbox list of distinct
values) / basic conditions. Sorting reorders rows (rewriting formula + merge refs like row moves);
filtering hides non-matching rows (reuse the hidden-row mechanism) and writes the filterColumn state
to `<autoFilter>`. Scope: value filters + asc/desc sort; number/date condition filters if cheap.

### Phase 2 - Render embedded images in the grid (DONE)
Read `xl/media/*` + the drawing anchors (we already parse twoCellAnchor for charts) and float images
over the grid on a layer like the chart overlay (anchored, scroll-glued, clipped under the header).
Render `<xdr:pic>` (raster) as an `<img>`/canvas from the media bytes (data URI). Shapes/text boxes:
best-effort (render the picture fill if any; otherwise a placeholder rectangle). Preserve on save
(untouched). No image editing yet (move/resize/replace is a later nice-to-have).

### Phase 3 - Author hyperlinks, comments, data-validation rules (DONE)
We render/follow these; now let the user create/edit them.
- Hyperlinks: a small dialog on a cell (URL or in-workbook target + display text); writes the
  worksheet `<hyperlinks>` + the sheet rels (external) or the `location` (internal). Remove too.
- Comments: add/edit/delete a note on a cell (author + text); writes the legacy comments part
  (+ vmlDrawing) or a threaded comment; the existing hover popover renders it.
- Data validation: define a list validation on the selection (inline list or a range); writes
  `<dataValidations>`; the existing dropdown UI then drives it.

### Phase 4 - Conditional formatting: author + close render sub-gaps (DONE)
- Author: a dialog covering every CF rule kind on the selection (cellIs incl. between, text,
  top/bottom, above/below average, duplicate/unique, is-true-formula, time-period, colour scale,
  data bar, icon set); writes `<conditionalFormatting>` + dxf.
- Render: icon sets (icon per cfvo threshold); `cellIs` cell-ref/formula operands and is-true-formula
  rules evaluate through the workbook formula engine; time-period rules match a cell's date serial
  against a window derived from today. ODS keeps the interoperable cellIs subset.

### Phase 5 - Sparklines (in-cell mini charts) (DONE)
Read the `x14:sparklineGroups` from the worksheet extLst (line / column / win-loss + colours) and
render a tiny canvas inside the host cell (lazy, shared with the chart Chart.js or a hand-drawn
mini-renderer to stay light). Preserve on save. Render-only first; authoring later if wanted.

### Phase 6 - Rich text within a cell (DONE)
Read multi-run strings (`<is>`/sharedStrings `<r><rPr>...`) into a run model and render each run
with its own style in the cell. Authoring (2026-07-24): while a cell is being edited, selecting a
sub-range of its text and clicking a run-applicable style (bold/italic/underline/strike/colour/size/
font) formats just that range as a run; the whole style toolbar drives it (buttons preserve the
edit selection via a mousedown preventDefault). The runs are serialised to xlsx `<r><rPr>` and ODS
`<text:span>` + interned text styles, and survive a LibreOffice round-trip. A pure `core/richtext.ts`
splits/merges/normalises runs (toggle by selection state; drop richRuns when uniform with the base);
undo/redo now snapshots richRuns/phonetic. Retyping a cell's text still clears its runs (offsets
would shift) - documented; re-apply after the text is settled.

### Phase 7 - Dynamic arrays / spill (MVP) (DONE)
The formula engine preserves legacy array formulas but does not spill modern dynamic arrays. Add
spill for the common producers (e.g. UNIQUE, SORT, FILTER, SEQUENCE, single-arg array returns):
when a formula returns a 2-D result, write it into the anchor + spill range (guarding collisions,
showing #SPILL! on conflict), and mark the range as a spill so edits/reads stay consistent. Scope to
what fast-formula-parser can evaluate; exotic spills stay preserved-only.

### Phase 8 - Correctness fixes (DONE)
- Shared-string edit: kept as inline-string on write (decided against mutating the shared sst).
  Rewriting the shared entry would silently change every other cell that shares the string; an
  inline string isolates the edit and leaves the sst count/uniqueCount valid (a harmless orphan).
  Verified: editing one of two cells sharing a string leaves the twin unchanged.
- Fixed a bug from Phase 6/furigana: editing a cell now clears its stale richRuns/phonetic, so a
  re-typed value no longer renders the old per-run styling or ruby over the new text.

### Phase 9 - Pivot tables: read, author, edit, refresh (DONE)
Detect pivots (xlsx `pivotTable`+`pivotCache`, ods `data-pilot-table`), outline them on the grid,
and let the user create/edit/refresh them. A pure compute engine (core/pivot.ts, row/column axis
model with prefix-aware aggregation) feeds both writers; the insert dialog assigns columns to Rows /
Columns / Values(+function) / Report Filter with a subtotals toggle and a live preview. Supports
nested row/column fields, multiple value fields (sum/count/average/min/max), page filters and
subtotals. Emits the native structure with refreshOnLoad; edit rewrites in place and refresh
recomputes from source, for authored and file-read pivots alike (the authoring spec is reconstructed
on read). Verified via LibreOffice round-trips for every shape (xlsx + ods). See done/PIVOT_AUTHORING.md.

### Also shipped (toolbar / UX, 2026-07-24)
- The toolbar folds its authoring controls into a "⋯" overflow menu (icon + label rows) when width
  runs out; the style cluster only collapses into its "Aa" menu as a last resort.
- The on/off + mutually-exclusive style buttons (bold/italic/underline/strike/align/valign/wrap)
  show a pressed state reflecting the active cell.
