# sheetedit: closing the remaining xlsx/ods gaps

Status (2026-07-24): Phases 1-9 below are DONE. Charts and pivot tables are feature-complete for
the common cases (see CHARTS_SPEC_GAPS.md, PIVOT_AUTHORING.md); the "Remaining / not yet done"
section at the end is the current, honest backlog. Each phase was a shippable increment: implement,
unit-test, browser-verify the visible behaviour, LibreOffice/round-trip check where a file format is
touched, commit, then bump omnitext. House rules: dependency-light, framework-agnostic, in-place
surgical XML edits (preserve untouched parts byte-for-byte), lazy-load anything heavy.

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
- Author: a dialog to add common CF rules on the selection (cellIs comparison, colour scale, data
  bar, top/bottom, text-contains); writes `<conditionalFormatting>` + dxf.
- Render sub-gaps: icon sets (draw the icon per threshold), and allow `cellIs` operands that are
  cell references / simple formulas (evaluate via the formula engine), not just numeric literals.

### Phase 5 - Sparklines (in-cell mini charts) (DONE)
Read the `x14:sparklineGroups` from the worksheet extLst (line / column / win-loss + colours) and
render a tiny canvas inside the host cell (lazy, shared with the chart Chart.js or a hand-drawn
mini-renderer to stay light). Preserve on save. Render-only first; authoring later if wanted.

### Phase 6 - Rich text within a cell (DONE)
Read multi-run strings (`<is>`/sharedStrings `<r><rPr>...`) into a run model and render each run
with its own style in the cell (currently flattened to one style). Editing: keep it simple - on
edit, a cell with rich runs either preserves the runs if the text is unchanged, or collapses to the
predominant style (documented). Full in-cell rich editing is out of scope for this phase.

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
on read). Verified via LibreOffice round-trips for every shape (xlsx + ods). See PIVOT_AUTHORING.md.

### Also shipped (toolbar / UX, 2026-07-24)
- The toolbar folds its authoring controls into a "⋯" overflow menu (icon + label rows) when width
  runs out; the style cluster only collapses into its "Aa" menu as a last resort.
- The on/off + mutually-exclusive style buttons (bold/italic/underline/strike/align/valign/wrap)
  show a pressed state reflecting the active cell.

## Remaining / not yet done (the honest backlog)
- **Pivots, advanced**: calculated fields/items, "show values as" (% of total / running total),
  pivot charts, and byte-identical layout to Excel (both apps re-flow the body on open anyway).
- **Conditional formatting**: icon-set authoring, `is-true-formula` / text-period rule authoring,
  and cell-reference / formula operands for `cellIs` (numeric literals only today; these round-trip).
- **Sparklines**: authoring on ods (xlsx is author + render; ods is render-only).
- **Rich text**: full in-cell rich-text *editing* (per-run styling is rendered + preserved, not
  authored from the UI).
- **Images / shapes**: move / resize / replace editing (render-only + preserved today).
- **Dynamic arrays**: exotic spill producers are best-effort; only UNIQUE/SORT/FILTER/SEQUENCE +
  TRANSPOSE/bare-range spill are supplied.
- **Preserved-only, no plans to edit**: form controls / ActiveX / slicers interactivity, sheet /
  workbook protection, print settings, outline grouping, themes.
- **Recalc**: a large but partial function set; unsupported functions or circular refs yield an
  error value (the file's cached value is shown as a fallback; desktop apps recompute on open).
- **Data validation**: only list (dropdown) rules are authored; other condition types round-trip.
- **Correctness caveat**: an edited shared-string cell is rewritten as an inline string (its
  sharedStrings entry may become an unreferenced orphan); Excel-fidelity of authored pivots/charts
  is verified through LibreOffice, not real Excel.

## Working notes
- Reuse existing machinery: the chart overlay layer pattern (Phase 2), hidden-rows (Phase 1),
  the DV dropdown + CF renderer + comment popover (Phases 3-4), createWorksheet/schema-sync for
  new parts. Prefer extending src/adapters/xlsx/*.ts modules already present (hyperlink.ts,
  datavalidation-ish, condformat.ts, comments, tables.ts) over new subsystems.
- Every format-touching change: round-trip test + LibreOffice open; every visible change:
  browser-verify. Bump omnitext after each phase.
