# sheetedit: closing the remaining xlsx/ods gaps

Charts are now feature-complete (see CHARTS_SPEC_GAPS.md). This plan closes the rest of the gaps
from XLSX_COVERAGE.md, in priority order. Each phase is a shippable increment: implement, unit-test,
browser-verify the visible behaviour, LibreOffice/round-trip check where a file format is touched,
commit, then bump omnitext. Keep the house rules: dependency-light, framework-agnostic, in-place
surgical XML edits (preserve untouched parts byte-for-byte), lazy-load anything heavy.

## Phases

### Phase 1 - Interactive sort & filter
Autofilter state already round-trips; add the UI. Click a column header (or a filter caret on the
autofilter range) to sort asc/desc, and a dropdown to filter by value (checkbox list of distinct
values) / basic conditions. Sorting reorders rows (rewriting formula + merge refs like row moves);
filtering hides non-matching rows (reuse the hidden-row mechanism) and writes the filterColumn state
to `<autoFilter>`. Scope: value filters + asc/desc sort; number/date condition filters if cheap.

### Phase 2 - Render embedded images in the grid
Read `xl/media/*` + the drawing anchors (we already parse twoCellAnchor for charts) and float images
over the grid on a layer like the chart overlay (anchored, scroll-glued, clipped under the header).
Render `<xdr:pic>` (raster) as an `<img>`/canvas from the media bytes (data URI). Shapes/text boxes:
best-effort (render the picture fill if any; otherwise a placeholder rectangle). Preserve on save
(untouched). No image editing yet (move/resize/replace is a later nice-to-have).

### Phase 3 - Author hyperlinks, comments, data-validation rules
We render/follow these; now let the user create/edit them.
- Hyperlinks: a small dialog on a cell (URL or in-workbook target + display text); writes the
  worksheet `<hyperlinks>` + the sheet rels (external) or the `location` (internal). Remove too.
- Comments: add/edit/delete a note on a cell (author + text); writes the legacy comments part
  (+ vmlDrawing) or a threaded comment; the existing hover popover renders it.
- Data validation: define a list validation on the selection (inline list or a range); writes
  `<dataValidations>`; the existing dropdown UI then drives it.

### Phase 4 - Conditional formatting: author + close render sub-gaps
- Author: a dialog to add common CF rules on the selection (cellIs comparison, colour scale, data
  bar, top/bottom, text-contains); writes `<conditionalFormatting>` + dxf.
- Render sub-gaps: icon sets (draw the icon per threshold), and allow `cellIs` operands that are
  cell references / simple formulas (evaluate via the formula engine), not just numeric literals.

### Phase 5 - Sparklines (in-cell mini charts)
Read the `x14:sparklineGroups` from the worksheet extLst (line / column / win-loss + colours) and
render a tiny canvas inside the host cell (lazy, shared with the chart Chart.js or a hand-drawn
mini-renderer to stay light). Preserve on save. Render-only first; authoring later if wanted.

### Phase 6 - Rich text within a cell
Read multi-run strings (`<is>`/sharedStrings `<r><rPr>...`) into a run model and render each run
with its own style in the cell (currently flattened to one style). Editing: keep it simple - on
edit, a cell with rich runs either preserves the runs if the text is unchanged, or collapses to the
predominant style (documented). Full in-cell rich editing is out of scope for this phase.

### Phase 7 - Dynamic arrays / spill (MVP)
The formula engine preserves legacy array formulas but does not spill modern dynamic arrays. Add
spill for the common producers (e.g. UNIQUE, SORT, FILTER, SEQUENCE, single-arg array returns):
when a formula returns a 2-D result, write it into the anchor + spill range (guarding collisions,
showing #SPILL! on conflict), and mark the range as a spill so edits/reads stay consistent. Scope to
what fast-formula-parser can evaluate; exotic spills stay preserved-only.

### Phase 8 - Correctness fixes
- Editing a shared-string cell: keep it a shared string (add/reuse an sst entry) instead of
  rewriting inline, so the sst stays consistent.
- Anything small surfaced along the way (CF operand edge cases, DV enforcement on paste, etc.).

## Explicitly deferred (documented as preserved-only, not in this plan)
- Pivot tables (needs a pivot cache + refresh engine) - very large, low incremental value.
- Form controls / ActiveX / slicers interactivity - niche.
- Image move/resize/replace editing (render-only in Phase 2).
- Full in-cell rich-text editing (render-only in Phase 6).

## Working notes
- Reuse existing machinery: the chart overlay layer pattern (Phase 2), hidden-rows (Phase 1),
  the DV dropdown + CF renderer + comment popover (Phases 3-4), createWorksheet/schema-sync for
  new parts. Prefer extending src/adapters/xlsx/*.ts modules already present (hyperlink.ts,
  datavalidation-ish, condformat.ts, comments, tables.ts) over new subsystems.
- Every format-touching change: round-trip test + LibreOffice open; every visible change:
  browser-verify. Bump omnitext after each phase.
