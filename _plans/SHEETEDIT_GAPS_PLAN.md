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
  used by the selected-item highlight. USER-DEFINED styles are resolved for real: styles.xml's
  extLst carries an x14:slicerStyles group whose x14:slicerStyleElement entries index the x14 dxfs
  list living in its own ext of the same extLst, so a slicer naming a custom style gets that
  style's actual selected / unselected fill and text colours (the dxf's bgColor is the fill, as in
  any differential format). The dxfs list is found by element name rather than by ext URI, so a
  file that groups its extensions differently still resolves; the hovered and no-data variants are
  parsed past, since sheetedit shows every cache item and has no hover state.

  AUTHORING A TABLE SLICER: the toolbar's slicer button creates one for a column of the Excel table
  under the cursor. Excel keeps the two families in SEPARATE extensions, so a table slicer is
  registered under {46BE6895-7355-4a93-B00E-2C351335B9C9} on the workbook (an x15:slicerCaches
  container whose child slicerCache stays x14) and {3A4CF648-6AED-40f4-86FF-DC5316D8AED3} on the
  worksheet, and its cache carries an x15:tableSlicerCache (tableId + the tableColumn's @id) in its
  own extLst under {2F2917AC-EB37-4324-AD4E-5DD8C200BD13} instead of naming a pivot. Every slicer
  cache name is also written as a workbook defined name resolving to #N/A, the way Excel does. The
  URIs come from EPPlus's ExtLstUris, which also confirmed the two slicer-style GUIDs.

  Not done: creating OLAP slicers. An OLAP slicer's cache addresses cube levels by unique name
  through a connection sheetedit does not model, so there is nothing meaningful to point a new one
  at; reading and displaying them read-only stays the honest limit.

  CAVEAT: there is no Excel here to verify against and LibreOffice drops slicers entirely, so all of
  this follows the MS-XLSX spec (URIs and content types cross-checked against excelize) and is
  verified in-app, by round-trip, and by asserting each registration is present - not against Excel.
- **Timelines**: read + render + interactive date filtering + write-back (xlsx). A timeline is the
  date-range sibling of a slicer and uses its own two parts: xl/timelines/*.xml for the view
  (x15:timeline name / cache / caption / level) and xl/timelineCaches/*.xml for the state
  (x15:state with x15:selection startDate+endDate and x15:bounds), anchored by a drawing
  graphicFrame carrying the 2012 timeslicer extension. The overlay splits the bounds into periods
  at the view's granularity (level 0 years / 1 quarters / 2 months / 3 days), lights the selected
  ones, and lets you click a period, shift-click to extend the range, or ⊗ to clear it. The chosen
  range filters every linked pivot by turning it into PivotSpec.itemFilters on the pivot's date
  field (items whose date falls outside [startDate, endDate) are dropped), and on save the range
  goes back into the cache part's x15:selection - added when there was none, removed when cleared,
  everything else byte-identical. Pivot item labels now carry the source cell's FORMATTED text, so a
  date field shows dates rather than serials. Same caveat as slicers: spec-faithful and round-trip
  verified, not Excel-verified.
- **Outline grouping**: read + render + author + write on both formats. Row groups draw an Excel-style
  gutter left of the row numbers: a bar per level, a +/- button on each group's SUMMARY line (the row
  just past the run, or just before it when <outlinePr summaryBelow="0">), and 1/2/3 level buttons
  that show a depth and collapse everything deeper. Group / ungroup / collapse / expand / clear also
  sit in the row and column header context menus, so column groups are authored without a second
  gutter. Collapsing hides the run and marks the summary line, which is exactly what the file
  records; re-expanding leaves a deeper collapsed group collapsed, like Excel.
  xlsx writes @outlineLevel / @collapsed / @hidden per <row> and one <col> span per column (widths
  and other column attributes carried over), plus <sheetPr><outlinePr> only when the summary side is
  not the default. ODF nests grouped rows in <table:table-row-group>, one level of nesting per
  outline level, with table:display="false" on a group whose rows are all hidden. Column groups are
  read from <table:table-column-group> and preserved, but ODS authoring covers ROWS only (the column
  runs are kept verbatim rather than re-nested). Both directions verified through LibreOffice
  round-trips. The gutter measures row positions off the rendered grid rather than the declared row
  heights, because a row's content can be taller than its <row ht>.
  Fixed on the way: an ODS file with row groups duplicated its rows on every save, because the
  rebuild treated <table:table-row-group> as a structural child to keep.
- **Freeze panes**: read + render + author + write on both formats. The toolbar button opens Excel's
  little menu (freeze at the cursor / top row / first column / unfreeze) and the row and column
  header menus carry "freeze rows above" / "freeze columns to the left". xlsx writes
  <pane xSplit ySplit topLeftCell activePane state="frozen"> into the first <sheetView>, updating an
  existing pane in place and removing it (plus the selections' stale @pane) on unfreeze, leaving the
  rest of the view (zoom, gridlines, selection) untouched. ODF keeps this in settings.xml, not in
  content.xml: per sheet, Horizontal/VerticalSplitMode = 2 with the counts in *SplitPosition and
  PositionRight / PositionBottom; a workbook with no settings.xml gets a minimal one plus its
  manifest entry.
  VERIFICATION CAVEAT: LibreOffice's headless converter DROPS per-sheet view settings on every
  output it writes, including ods -> ods, so a freeze cannot be round-tripped through it. What it
  does prove: an xlsx -> xlsx pass keeps the written <pane> verbatim AND makes LibreOffice expand
  the per-pane <selection> entries, which it only does after parsing the frozen panes; and its xlsx
  export of the written ods picks up that sheet's scroll position, so it parsed the settings entry.
- **Split panes**: read + render + author + write, sharing the freeze model. Both kinds put a
  DIVIDER on the boundary that drags to move the split and double-clicks to remove it; "split at this
  cell" sits next to the freeze entries in the toolbar menu. xlsx states a split's offset in TWIPS
  (1/20 pt) rather than in line counts, so the same boundary is written differently per state and
  read back by walking the line sizes (topLeftCell short-circuits that when present). ODF mode 1 is
  read through PositionRight / PositionBottom, which name the trailing pane's first line and so avoid
  its pixel unit entirely.
  A ROW SPLIT IS TWO REAL VIEWPORTS. The grid was one scroll container with one virtualized table
  and sticky frozen cells; a split now adds a second container below it, and the renderer works on a
  Pane record ({scrollEl, tableEl, window, inputs, tds, header}) rather than on module-level state.
  Each pane keeps its own rendered window and its own cell elements, because a split can put the SAME
  cell on screen twice - so the cell lookups became inputAt / tdAt / tdsAt over the pane list, with
  the pane last pointed at searched first so Enter keeps the caret where the user is working. The
  panes share horizontal scroll, only the top one draws the column header, and nothing is sticky
  inside a split pane (the boundary IS the pane edge); a freeze still renders as one viewport with
  the sticky block, untouched.
  OVERLAYS FOLLOW THE SPLIT. All six floating layers (charts, images, shapes, slicers, timelines,
  pivot tags) were identical in shape - a box over the grid with an inner element translated by the
  scroll - so they now share core/ui/overlay-hosts.ts, which keeps one box PER pane and hands each
  object the inner of the pane that shows its anchor row. Charts move their Chart.js instance with
  the box rather than being rebuilt.
  GOTCHA, twice over: the host picker must decide from the RENDERED row element, not from
  yOfRow. The model's uniform row height is ~1px off the rendered one, which was enough to make the
  top pane claim a row it does not actually show; and headerH came from a querySelector("thead")
  that never matches (the header row is appended straight to the table), so all eight geom() sites
  were feeding a constant. There is now one headerH() helper measuring the corner cell.
  BOTH AXES. The pane list is now a 2x2 of quadrants (row band 0/1 x column band 0/1) inside a
  grid area of two column bands, each stacking its row bands. A pane knows whether it draws the
  column header (top band) and the row numbers (left band), and nothing is sticky inside any split
  pane since the bands ARE the panes. Sync is by band: same row band -> same scrollTop, same column
  band -> same scrollLeft.
  GOTCHA: a trailing band can only open past its boundary when the sheet is wider/taller than the
  band; on a narrow sheet it legitimately shows the leading lines too, exactly as Excel does. The
  snap that lines a fresh band up with the first line past the boundary stays PENDING until the
  reference line is actually measurable, because an early render can happen before the panes have
  any size.
  LIMITS: an object straddling the boundary is drawn once, in the pane showing more of its anchor
  row. The outline gutter is drawn for the upper-left pane.
  And an ODS split the user MOVES is written back as a frozen boundary, since ODF states the split
  position in a LibreOffice view-pixel unit that could not be determined here (its headless converter
  drops view settings, so there was no way to observe one); an untouched split is never rewritten.
  GOTCHA: the divider's drag handler must restore the bar IN PLACE on a no-op click rather than
  re-render, or the element is swapped out between the two clicks and dblclick never fires - the same
  trap the image and shape layers hit.
  GOTCHA: NOTHING about the dividers may come from the layout model. The bars were placed at
  headerH + yOfRow(n), which put the horizontal one ~10px above the real gridline (the model's
  uniform ROW_H disagrees with the rendered rows, and there is no <thead>, so every headerH read fell
  back to a constant), and their length came from a container rect that can still be stale, so the
  vertical one ran past the scrollbar into the sheet tabs. Both now measure the rendered
  th.rownum / th.colhead edges, and the bars live in a clipping box sized from the panes' CLIENT box
  so they can never reach a scrollbar. The same mismatch made a fresh split show a sliver of the row
  above, so the top viewport is trimmed to its last row's rendered bottom and the lower one is
  snapped to the first row past the boundary, once, at creation.
- **Preserved-only, not authored yet**: form controls / ActiveX.
- **Workbook themes** are read (palette + scheme fonts) and switchable. The trick is that cell styles
  keep the theme *reference* (index + tint) beside the resolved colour, so a switch re-resolves
  instead of having baked in whatever palette the file was read with. Cells share the style objects
  in the resolved pool, so mutating the pool updates untouched cells; cells restyled in-app hold
  their own object, which is why the switch walks both. Only theme1.xml is rewritten, patched in
  place so the unmodelled fmtScheme survives. xlsx-only: ODF has no indexed palette.
- **Styling** all lives in src/sheetedit.css, compiled to a string module by scripts/build-css.mjs
  because the package ships through plain tsc. Colours go through a --sheetedit-* token layer with
  light defaults and a dark set on prefers-color-scheme / [data-theme]. Tests fail on a literal
  colour in a rule, an undefined token, or a stale generated module. Inline styles are reserved for
  values that are data (a cell's own formatting, a measured overlay position).
- **Print settings** are read, rendered and authored in both formats. The model is stated in xlsx's
  terms (inches, a paper-size id, scale as a percentage) because that format is the more explicit of
  the two; the ods adapter converts to ODF's page-layout / master-page pair. The print area and the
  page breaks are drawn on the grid as cell borders rather than an overlay, so they follow scrolling,
  virtualization and split panes for free.
  - The two margin models differ: an xlsx top margin encloses the header block, an ODF page margin
    stops where it starts. Converted both ways and asserted on in the tests.
  - LibreOffice computes an OOXML margin from header CONTENT height plus the header's own
    margin-bottom, ignoring fo:min-height, and defaults that margin to 20mm when absent - which
    inflated every ods -> xlsx margin by ~0.5in. Fixed by stating the spacing the way LibreOffice
    states it (block height less one nominal text line, 0.1389in), verified at exactly 0.75in.
  - Header/footer text keeps the file's raw &-codes, so formatting codes round-trip; only the ODF
    write path drops font/size codes, since ODF states those as span styles.
  - **Printing** (core/print-render.ts) paginates the print area itself and hands the pages to
    window.print(): paper size, margins, scale or fit-to, manual breaks, repeated title lines and
    page order all decide where the breaks land, and none of that survives simply printing the grid.
    Pages are exact paper boxes with @page margin 0 so the margins and the header/footer bands are
    ours to place. The browser's own URL/date strip is a print-dialog setting we cannot control.
    Do NOT wait on requestAnimationFrame before print(): it is throttled in a background tab, which
    left the pages built and the dialog never opening. Force layout with an offsetHeight read.
- **Protection** is read, enforced and authored in both formats. The flags are stated the way OOXML
  states them (each one names a BLOCKED action, each with its own default), and the ods adapter
  inverts them, because ODF's loext flags are permissions. Two layers decide editability: the sheet
  carries protection AND the cell's style says locked, which is the default in both formats, so
  unlocking a range is what makes a protected sheet usable. Enforced at every write chokepoint
  (typing, paste, fill, clear, insert/delete lines, sort, formatting, sheet add/delete/rename/move),
  with a transient notice so a refusal is never silent.
  - No password is ever computed or verified: the formats store a hash, not encryption, so any tool
    can lift protection. A hash found in a file is preserved verbatim on re-save; unprotecting drops
    it, which the dialog says out loud.
  - ODF has no equivalent for the format / sort / autofilter / pivot flags, so those are dropped on
    an ods save and come back at their (blocked) defaults, which is the safe direction.
  - LibreOffice does not carry workbook structure protection through its OOXML export (its own
    ods -> ods keeps it), so that flag was verified within ODF only.
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
