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
on read). Verified via LibreOffice round-trips for every shape (xlsx + ods). See PIVOT_AUTHORING.md.

### Also shipped (toolbar / UX, 2026-07-24)
- The toolbar folds its authoring controls into a "⋯" overflow menu (icon + label rows) when width
  runs out; the style cluster only collapses into its "Aa" menu as a last resort.
- The on/off + mutually-exclusive style buttons (bold/italic/underline/strike/align/valign/wrap)
  show a pressed state reflecting the active cell.

## Remaining / not yet done (the honest backlog)
- **Pivots**: feature-complete. "Show values as", calculated fields, calculated items and a pivot
  chart are all DONE (see PIVOT_AUTHORING.md). Not attempted: byte-identical layout to Excel (both
  apps re-flow the body on open anyway). Caveat: show-values-as / calculated fields / calculated
  items are honoured by Excel and sheetedit's display but ignored by LibreOffice's xlsx pivot
  rebuild; the calculated-item OOXML is emitted per spec but unverified in Excel (only that the file
  opens cleanly in LibreOffice).
- **Conditional formatting**: feature-complete. Authoring covers every rule kind; icon sets,
  is-true-formula rules, `cellIs` cell-ref/formula operands and time-period rules all render (DONE).
  ODS authoring stays on the interoperable cellIs subset (the graphical/formula rules have no ODF
  form that survives a LibreOffice round-trip).
- **Sparklines**: author + render on both xlsx and ods. ODS authoring emits the LibreOffice
  calcext:sparkline-groups (rebuilt from the model after the rows on save; untouched groups kept
  verbatim), verified through a LibreOffice round-trip.
- **Rich text**: feature-complete. Per-run styling is rendered, authored (select a sub-range while
  editing, then bold/italic/underline/strike/colour/size/font) and written to xlsx + ods (DONE).
  Retyping a cell's text clears its runs (offsets shift); re-apply after the text is settled.
- **Images**: move, resize and replace authored on both xlsx and ods. Drag the picture to move,
  drag the corner handle to resize, double-click to pick a replacement file. xlsx writes back the
  two-cell drawing anchor (converting a oneCell/absolute anchor to a twoCellAnchor); ods rewrites the
  draw:frame's svg:x/y/width/height relative to its anchor cell (no re-parenting). Replace swaps the
  media part in place for the same extension, or writes a new part + retargets the rel/href +
  registers the content type / manifest entry for a different one. Verified through LibreOffice
  round-trips.
- **Shapes**: read + render + author + write on both xlsx and ods. Drawing shapes (rectangle,
  rounded rectangle, ellipse, triangle, diamond, parallelogram, pentagon, hexagon, star, arrow,
  line; unknown presets fall back to a rect but keep their preset name so they round-trip) render as
  an SVG overlay with fill / outline / centered text. Polygon shapes share one pure vertex helper
  (core/shape-geom.ts) between the SVG renderer and the ODF enhanced-path writer.
  Insert from the toolbar; move (drag) / resize (corner handle) / edit (double-click -> fill /
  outline / text) / delete (the corner x). xlsx writes xdr:sp in the drawing part (patched in place
  for edits, appended for new); ods writes draw:rect / draw:ellipse / draw:line into a table:shapes
  container with an interned graphic style. Verified through LibreOffice round-trips on both formats.
- **Dynamic arrays**: the whole Excel 365 shaping family spills - UNIQUE / SORT / SORTBY / FILTER /
  SEQUENCE / TRANSPOSE / RANDARRAY / TAKE / DROP / CHOOSEROWS / CHOOSECOLS / EXPAND / HSTACK /
  VSTACK / TOROW / TOCOL / WRAPROWS / WRAPCOLS / TEXTSPLIT, plus bare-range spill.
- **Slicers**: read + render + interactive filtering + write-back (xlsx). The three parts are read
  (xl/slicers/*.xml for the view, xl/slicerCaches/*.xml for the selection, and the drawing's
  graphicFrame sle:slicer extension for the anchor); item labels come from the pivot cache field's
  sharedItems, since the slicer stores only indices. The overlay renders a real slicer panel: click
  an item to narrow to it, ctrl/cmd-click to multi-select, and the ⊗ button clears the filter. The
  selection drives PivotSpec.itemFilters (a multi-select complement to the single-select pages) and
  the linked pivots recompute; on save the new selection is written back into the cache part as
  `s="1"` per selected item, leaving everything else byte-identical. Selections are mapped by item
  LABEL, since the cache's item order need not match the engine's.

  AUTHORING: "Insert slicer" on a pivot's tag menu creates one for any field the pivot groups by
  (those are the fields the cache writes sharedItems for). It writes the whole package Excel needs:
  the cache part and the view part, both content types (vnd.ms-excel.slicer+xml and
  .slicerCache+xml), a workbook relationship for the cache and a worksheet relationship for the
  view, both extLst registrations (x14:slicerCaches under {BBE1A952-AA13-448e-AADC-164F8A28A991} on
  the workbook, x14:slicerList under {A8765BA9-456A-4dab-B4F3-ACF838C121DE} on the worksheet), and a
  drawing graphicFrame with the sle:slicer extension. The pivotCacheId is resolved from
  workbook.xml's <pivotCaches> rather than assumed. GOTCHA: the worksheet extLst must be written
  AFTER any drawing work, because ensureSheetDrawing re-parses the sheet and replaces sheet.doc.

  KINDS: pivot slicers filter their pivots; TABLE slicers (an x15:tableSlicerCache in the cache's
  extLst, tableId + column, where column is a tableColumn @id mapped to an offset in the table
  range) bind to a ListObject column - their items are that column's distinct values and toggling
  one hides the table's non-matching rows; OLAP slicers are read from their own <x14:olap> captions
  and selections and rendered READ-ONLY, since no OLAP source is modelled to filter. Slicer STYLES
  are read and the built-in families (SlicerStyleLight/Dark/Other N) map to an Office theme accent
  used by the selected-item highlight.

  Not done: timelines (a separate timeline / timelineCache part type, not a slicer); custom slicer
  style definitions (the name round-trips, only built-in families are coloured); and creating table
  or OLAP slicers (creation covers pivot slicers).

  CAVEAT: there is no Excel here to verify against and LibreOffice drops slicers entirely, so all of
  this follows the MS-XLSX spec (URIs and content types cross-checked against excelize) and is
  verified in-app, by round-trip, and by asserting each registration is present - not against Excel.
- **Preserved-only, not authored yet**: form controls / ActiveX, sheet / workbook protection,
  print settings, outline grouping, themes.
- **Recalc**: fast-formula-parser ships many of its functions as empty stubs, so the gaps are filled
  in core/functions.ts (statistics, multi-criteria aggregates, MATCH / XLOOKUP / XMATCH / CHOOSE /
  LOOKUP, SUBTOTAL / AGGREGATE, text and SWITCH), core/financial.ts (the whole financial family),
  core/dynamic-arrays.ts (the Excel 365 shaping family) and core/reference-fns.ts (OFFSET /
  INDIRECT). A coverage probe over 81 real-world formulas went 18 -> 81 passing.
  - OFFSET / INDIRECT return a *reference*, not a value, so the range they name can feed an
    aggregate or spill (`SUM(OFFSET(A1,0,0,3,1))`). OFFSET needs its first argument's ref intact,
    which the parser only preserves for functions on its no-data-retrieve list, so allowRawRefs()
    adds it there on each parser instance.
  - LET has no runtime scope in the parser (arguments evaluate eagerly), so core/let-expand.ts
    expands it away textually before parsing: innermost first, token-aware so it never rewrites
    inside a string literal or an identifier used as a function call. The expansion also runs for
    the dependency parser, so a LET formula still recalculates when its sources change.
  Unsupported functions or circular refs still yield an error value with the file's cached value
  shown as a fallback, and the editor recalculates on edit rather than on open (like Excel).
- **Data validation**: all rule types author + read + validate on both formats. List rules show a
  dropdown; whole / decimal / date / time / text-length / custom-formula rules outline a cell whose
  value breaks the rule (custom formulas are round-tripped, not evaluated live). xlsx writes the
  type/operator/formula1/formula2; ods writes the LibreOffice-verified content-validation condition
  (`is-between(a,b)` with commas, `cell-content()OP a` unspaced), which round-trips through LibreOffice.
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
