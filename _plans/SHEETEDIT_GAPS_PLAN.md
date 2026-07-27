# sheetedit: closing the remaining xlsx/ods gaps

This is the living record: what shipped, what it cost, and what is still missing. Finished plans
for individual features live in `done/`; see `README.md`.

Status (2026-07-26): Phases 1-9 below are DONE. Charts and pivot tables are feature-complete for
the common cases (see done/CHARTS_SPEC_GAPS.md, done/PIVOT_AUTHORING.md); the "Remaining / not yet done"
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
on read). Verified via LibreOffice round-trips for every shape (xlsx + ods). See done/PIVOT_AUTHORING.md.

### Also shipped (toolbar / UX, 2026-07-24)
- The toolbar folds its authoring controls into a "⋯" overflow menu (icon + label rows) when width
  runs out; the style cluster only collapses into its "Aa" menu as a last resort.
- The on/off + mutually-exclusive style buttons (bold/italic/underline/strike/align/valign/wrap)
  show a pressed state reflecting the active cell.

## Named for later (nothing here is forgotten, and none of it is a mystery)

Each of these is understood; what stops it is stated, so picking one up starts from a known place.

| | |
|---|---|
| **rgColumnInfo: per-column widths** | A multi-column list renders its columns evenly. The count and the bound column are read; the per-column WIDTHS live in a ColumnInfo record whose field layout is not in the published MS-OFORMS index (only `cColumnInfo`, "the last column with a non-default width", is documented). Guessing a binary record is how the ScrollBar mask went wrong once already. Next step: find the record in the downloadable .docx of the specification rather than the HTML index, or measure it against a file Excel wrote with known widths. |
| **Frame / MultiPage / TabStrip** | Not a missing property but a missing STRUCTURE: the control's stream holds a whole embedded form - a ClassTable, a sites array, and a child stream per control. That is the "parent controls" half of MS-OFORMS, roughly the size of everything already done for the leaf controls. They are UserForm controls that only reach a worksheet through "More Controls", so this is real work for a rare case. |
| **Caption on a CommandButton / Label that has none** | The insert-a-missing-property rebuild is written for the MorphData family, which covers the checkbox, option button, toggle and text box. The other two would need the same emit-from-recorded-fields treatment for their own layouts. Nothing sets a caption from the UI yet either, so this is a UI gap before it is a format one. |
| **EMF text drawn at the wrong size (upstream)** | FIXED UPSTREAM, awaiting merge: [ChristopherVR/emf-converter#9](https://github.com/ChristopherVR/emf-converter/pull/9). The font's `lfHeight` was applied without the window/viewport mapping every coordinate goes through, so on a metafile with a non-identity map mode the text drew many times too large and painted over the picture. LibreOffice's own EMF export sets MM_ANISOTROPIC with window 2540x2540 against viewport 132x132 and a font of lfHeight -635, which comes out ~19x oversized. The PR passes the mapping factor the callers already hold (`gmh(rCtx, 1)` for EMF, `mh(1)` for WMF); upstream's 784 tests pass and three were added. Once it lands, bump the dependency and drop the caveat from the metafile note above. |
| **Third-party ActiveX** | Irreducible. OOXML says the content of such a control "shall be solely determined by the corresponding object", so the format belongs to whoever wrote the control. |

## Remaining / not yet done (the honest backlog)
- **Pivots**: feature-complete. "Show values as", calculated fields, calculated items and a pivot
  chart are all DONE (see done/PIVOT_AUTHORING.md). Not attempted: byte-identical layout to Excel (both
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
- **Preserved-only**: ActiveX (carried through a save, never rendered). The macro dialog now says so,
  since the control otherwise leaves an unexplained gap on the grid.
  - RESEARCHED 2026-07-26, and it corrects an earlier claim here that the `.bin` is untouchable.
    **[MS-OFORMS]** is a public, normative Microsoft Open Specification for exactly this binary
    persistence, so a Forms 2.0 control (which is what Excel's ActiveX toolbox inserts) is
    parseable on the same clean-room basis as MS-OVBA. A THIRD-PARTY COM control still is not: the
    OOXML spec says the content "shall be solely determined by the corresponding object", so its
    format belongs to whoever wrote the control.
  - The surprise is where the data sits. A real `activeX1.xml` (checked against one in the wild)
    carries ONLY `ax:classid` plus a relationship to the `.bin` when persistence is
    `persistStreamInit`. Caption, value, size, colours and the linked cell are all in the binary.
    So there is no cheap "read the XML and render it" half: the work IS the MS-OFORMS parser.
  - DONE for reading: CommandButton and every MorphData kind (checkbox, combo, text, list, option,
    toggle, label) give up their kind, caption and value. The parse self-checks against the stream's
    own `cb` field and returns the kind alone rather than half-right values when it does not land
    exactly. ScrollBar and SpinButton stay kind-only: their mask's bit order is unconfirmed.
  - Two bugs it uncovered, both worse than the missing feature. Every ActiveX control was read as a
    blank "label" because its part was fed to the formControlPr reader; and each was read TWICE,
    since Excel writes a control under both mc:Choice and mc:Fallback. A workbook with six controls
    drew twelve phantom labels.
  - Samples: found via a plain web search, not GitHub code search, which does not index binaries.
    Contextures' combo-box tutorial workbook has six. It is their copyright, so it is used locally
    and never committed; the committed fixtures are synthetic and were verified byte-identical to
    the real streams before being written down.
  - `linkedCell`, `listFillRange` and `macro` are read off `<controlPr>`, which is where Excel keeps
    the properties that are its own rather than the control's. A combo with both is fully live: it
    lists from the named range and writes the chosen TEXT to the linked cell, where a form control
    writes the item's position. Verified against the real file, whose three combos carry
    RegionList / MonthList / DayList and linked cells H9 / H5 / C7.
  - DONE: an MS-OFORMS writer for the Value. Same length patches IN PLACE so the stream stays
    byte-identical including padding; a length change rebuilds the ExtraDataBlock and carries the
    unmodelled trailing blocks (StreamData, TextProps, rgColumnInfo) across. It refuses on any
    control the reader would not vouch for, so a write never proceeds where a read would not.
    Identity confirmed on four real Excel streams, and a change survives editor -> save -> reread.
  - VERIFICATION CAVEAT, and it is a real one. LibreOffice does not surface ActiveX from xlsx, so
    unlike the VBA writer (which an independent engine was made to RUN) there is no outside judge.
    It reopens a rewritten workbook without complaint, and that only proves the package is not
    corrupt. Excel remains the untested case.
  - ScrollBar, SpinButton and Label are read too, from their own spec tables fetched rather than
    guessed. Three findings worth keeping: a LABEL IS NOT A MORPHDATA CONTROL, it has its own
    LabelPropMask, which is an easy and wrong assumption; ScrollBar and SpinButton do NOT share a
    mask (the spin has no LargeChange or ProportionalThumb, and fMousePointer moves to the end);
    and fPrevEnabled / fNextEnabled are mask bits with NO field behind them, so consuming bytes for
    them would push every later read out of place. The layouts are table-driven now, one table per
    family, each ending on the same cb check.
  - Writing covers every string a control carries (Value, Caption, GroupName), not just Value. The
    layout records each string's own place, so rewriting the middle of three moves the last
    correctly. Identity holds on all nine real streams for every property they carry.
  - SPEC AUDIT (2026-07-27), against [MS-OFORMS] section 2 property by property. What used to be
    read and thrown away is now kept and used:
    - **VariousPropertyBits**, the 32-bit field a dozen booleans share. Enabled, Locked, BackStyle,
      ColumnHeads, MatchRequired, Alignment, Editable, WordWrap, AutoSize, MultiLine. The bit
      positions are pinned by the spec's own file-format defaults (0x2C80081B for the MorphData
      family, 0x0080001B for a label), which is a real check on having read the table straight.
    - **DisplayStyle**, which is the ONLY thing separating an editable combo (3) from a drop-list
      one (7): they share a class id. Also the numeric properties around it - MaxLength,
      PasswordChar, BorderStyle, BorderColor, SpecialEffect, ScrollBars, ListRows, ListWidth,
      ColumnCount, BoundColumn, TextColumn, MultiSelect, MatchEntry, ListStyle, ShowDropButtonWhen,
      DropButtonStyle, MousePointer, Accelerator, PicturePosition - and SmallChange / LargeChange /
      Orientation / Delay / ProportionalThumb on the range controls.
    - **TextProps**, the font, which is its own versioned structure sitting after the control's cb.
      Name, size (twips), bold / italic / underline / strikeout, weight, paragraph alignment.
    - **The Image control**, which had no layout at all. Its mask puts fAutoSize and fPictureTiling
      in the BIT rather than in a field, like the scroll bar's fPrevEnabled.
    - **StreamData pictures**, sniffed to a MIME type and rendered from a data: URI. A metafile
      (WMF/EMF) is a drawing program rather than an image, so it is skipped rather than mislabelled.
  - What that buys on the page: a TextBox is an editor (a textarea when MultiLine, a password box
    when it names a PasswordChar, with its MaxLength), a ToggleButton is a button that stays down,
    an Image shows its picture, a list honours ListRows and MultiSelect, every control wears the
    file's own font and colours, a disabled control looks and behaves disabled, MousePointer is a
    CSS cursor and Accelerator is an access key. An OLE_COLOR naming a Windows SYSTEM palette entry
    is left unset on purpose: its colour is the desktop theme's, not the document's.
  - The writer can now ADD a property the control does not yet carry, which is what an empty text
    box needs: its Value bit is clear, so there is nothing to patch. That path re-emits the whole
    DataBlock from the fields the read recorded (splicing would not do, since a length word has to
    land 4-aligned and inserting one shifts everything after it) and reads its own output back
    before returning it. Same-length changes still patch in place, byte for byte.
  - **Multi-column lists** render as a grid, since that is what Excel draws and a `<select>` cannot
    be one. The columns come from the source range (which is where the items come from anyway) and
    BoundColumn decides which one the control reports - 0 means the row number, as in Excel.
- **Windows metafiles (WMF / EMF)** render now, which they never did: an `xl/media/*.emf` came
  through as a `data:image/emf` URI and a browser drew nothing at all. A metafile is not an image -
  it is a recorded list of GDI drawing calls - so showing one means replaying them onto a canvas.
  That is emf-converter (Apache-2.0, no dependencies), lazy-loaded so a workbook without one never
  pays for the code, and failing soft: no picture rather than a broken one. It covers sheet images,
  an ActiveX control's Picture, and anything else that arrives as a data URI.
  - GOTCHA, and it is ours to handle: the converter ignores a placeable WMF's own frame and renders
    into a square canvas at its 8192px cap - a 1.7MB PNG of mostly white for a picture two inches
    across. `metafileSize` reads the frame first (a WMF's placeable bounding box and units-per-inch,
    an EMF's rclFrame in hundredths of a millimetre) and passes it as an explicit cap, which puts
    both formats on the same sane size.
  - KNOWN DEFECT, fix submitted upstream (emf-converter#9): text inside a metafile whose map mode is
    not the identity draws at the file's LOGICAL height, so a label comes out many times too large
    and reads as a white block punched through the drawing. Everything else is drawn through the
    mapping helpers, which is why only text is affected. Diagnosed against LibreOffice's own
    rasterisation of the same files; the drawing is otherwise faithful, so this shipped rather than
    waiting. Bump the dependency and delete this note once the PR is merged and released.
- **Grid metrics** come from the workbook, not from constants, and getting this wrong showed up as
  an ActiveX bug rather than a layout one. A column's `width` is in character units of the NORMAL
  STYLE's font, so the conversion needs THAT font's maximum digit width (7px for Calibri 11, 9px for
  Calibri 14); hardcoding 7 ran a 14pt workbook's every column 28% narrow. An anchored object's
  position is the sum of the widths to its left, so the error compounded and the file's combo boxes
  drew on top of their own labels. `<sheetFormatPr defaultRowHeight/defaultColWidth>` was ignored
  too, which did the same thing vertically. GOTCHA: honouring a 21px row needed the cell editor to
  stop padding itself out to a height, since that put a 23px floor under every row.
- **Text spill**: a text too long for its column runs on over the empty cells beside it, and stops
  at the first that holds anything, as Excel does. The direction follows the alignment (a
  right-aligned label runs back over its own label column); numbers do not spill; rich text does,
  measured run by run. An edit re-renders rather than refreshing displays, since whether a neighbour
  is empty is exactly what a spill depends on. GOTCHA found here: the off-screen twin used to
  measure wrapped text was appended once and detached by the next render, after which every wrapped
  cell measured zero high. The editor is laid OVER its cell (position:absolute, inset) rather than
  sized by padding: a percentage height does not resolve inside a table cell, so height:100% left
  the focus ring stopping short of the gridlines on any row taller than the text.
- **Shape `<xdr:style>`**: a shape from Excel's gallery carries no fill or line of its own, only a
  colour plus an INDEX into the theme's fillStyleLst / lnStyleLst, and a fontRef for its text. That
  entry is usually a gradient of tints of the named colour. Ignoring it left every gallery shape
  unfilled, which on a dark grid is an invisible button with unreadable text. Resolved now, with
  DrawingML's lumMod / lumOff / tint / shade / satMod applied through HSL. The gradient is rendered
  AS a gradient: the overlay emits an SVG `<linearGradient>` def per shape and fills from it, with
  `<a:lin ang>` (60000ths of a degree, clockwise from east, which is already SVG's y-down sense)
  giving the direction. `shape.fill` still carries the first stop, so the property editor and the
  writer keep a single colour to work with. A shape's OWN `<a:gradFill>` is read the same way.
  GOTCHA: an SVG gradient id is document-global, so the ids are sequenced; a fixed one would make
  every shape on the page take the first shape's gradient.
- **A shape with a macro is a button**: `<xdr:sp macro>` is read, and clicking the shape runs it
  through the same path a form control's button uses. Three bugs found doing this. A drag started
  on ANY button, so a right-click left the shape following a pointer with nothing held down. A MOVE
  rewrote the shape's paint from the model's single colour, so dragging a gallery shape silently
  repainted it and a theme-styled one lost its look outright: only a restyle touches spPr now
  (`styleDirty`), a move rewrites the anchor and nothing else. And the macro name a shape carries is
  qualified with `!` (`[0]!Sub`), not `.`, which the control path stripped.
- **Worksheet.OLEObjects**: the ActiveX controls, reachable from a macro. `.Count`, by name or by
  1-based index, `For Each`, and per object `.Object`, `.Name`, `.Index`, `.LinkedCell`,
  `.ListFillRange`, `.TopLeftCell`. The control itself gives `.Value` / `.Text` / `.Caption` /
  `.ListCount` / `.ListIndex` / `.List(i)`, and Value and ListIndex are settable: the write goes to
  the persisted binary AND to the linked cell, since Excel keeps the two in step. Everything else
  refuses by name. The list comes from the host, because a listFillRange is usually a DEFINED NAME.
- **Structured (table) references**: `Table1[Units]`, `Table1[[#Headers],[Units]]`, `Table1[@Units]`,
  `Table1[[A]:[B]]`, and a bare `Table1`. The formula engine only knows A1, so these were invisible
  and every total written against a table showed the file's stored number under a "could not be
  evaluated" note - which is what Excel's own UI writes the moment a range becomes a table. They are
  rewritten to plain ranges before parsing, like LET. A selector that cannot be honestly resolved (a
  totals row, which the model does not track; a column that is not there) becomes a call to a
  function that does not exist, so the parse STOPS: leaving the text alone was worse, since the
  engine then read the bare table name and quietly produced a number from the wrong place.
- **ListObjects**: `Add(xlSrcRange, range, , xlYes)` writes the whole package (table part, worksheet
  rel, `<tableParts>` entry, content type), plus `.Name` / `.TableStyle` / `.Range` /
  `.HeaderRowRange` / `.DataBodyRange` and lookup by name or index. Verified by a LibreOffice
  round-trip of a workbook whose table a macro created. `Range.CurrentRegion` and `Range.AutoFit`
  came with it; AutoFit is a STATED APPROXIMATION (the longest text in the column, in the workbook's
  character units), since real AutoFit measures rendered glyphs.
- **Three VBA semantics bugs the real macro found**, all of the "plausible wrong answer" kind:
  - `Set ws = ActiveSheet` FOLLOWED the active sheet, because ActiveSheet was a live proxy. A macro
    that stored it, added a sheet, then looped over the stored sheet's controls walked the new,
    empty sheet instead. ActiveSheet is a property: vbalang grew lazy globals so evaluating the name
    produces a concrete Worksheet, and ordinary assignment then captures it.
  - A one-dimensional array written to a range filled DOWN a column. In Excel it is a row, which is
    what every macro writing a header row with `Array(...)` expects.
  - `MsgBox("delete it?", vbYesNo)` always answered vbOK, so the not-vbNo branch ran every time -
    and that branch is usually the destructive one. vbalang now refuses a question without an `ask`
    host hook; sheetedit supplies `confirm`, the only synchronous ask a page has.
  - `TypeName(ctl.Object)` reported the wrapper, not the control. It names the Forms 2.0 class now
    (ComboBox / CheckBox / CommandButton / ...), which is what a macro branches on.
- **Hidden sheets** are read, honoured (no tab is drawn) and authored on both formats, plus
  Worksheet.Visible in VBA with Excel's three states. ODF keeps this in the sheet's TABLE STYLE
  (`<style:table-properties table:display>`), NOT on `<table:table>`: the attribute on the element
  parses fine and LibreOffice ignores it entirely. Sheets share table styles, so hiding one clones
  its style rather than editing it, or its neighbours vanish too. Both directions LibreOffice-verified.
- **What a macro still cannot do**: Worksheet.ExportAsFixedFormat (a browser reaches a PDF only
  through the print dialog, where the user chooses it), Worksheet.Copy with no Before/After (that
  makes a new workbook in Excel), Workbook.Save/SaveAs/Close, and anything reaching outside the
  page. Worksheet.Copy copies the grid, not the drawing layer: a chart or image needs its own part
  copied and re-registered, and a half-copied drawing is worse than an absent one.
- **Form controls** are read, rendered and interactive: the linked cell is the point, so a checkbox
  writes TRUE/FALSE there, a dropdown the 1-based index, a spinner its value, each triggering a
  recalc. State is read from ctrlProps with the VML as fallback (files predating ctrlProps have only
  the VML) and written back to BOTH, since an older reader looks only at the VML. Controls can be
  created, edited, deleted, dragged and resized. A button runs the macro its `<x:FmlaMacro>` names,
  through the same path the macro viewer uses; one with nothing assigned is drawn disabled. A file
  written by older Excel has no `<controls>` element at all, only VML shapes, and those are read as
  controls too. Radios clear the rest of their group box.
  - The VML shape id is "_x0000_s1025" and the worksheet's @shapeId is the trailing number, so it
    must be matched from the END; a lazy match from the start stops at the first underscore.
  - LibreOffice validates the structure but NOT the checkbox state: a file that was never checked
    still reads back as checked, so state fidelity is verified by our own round-trip.
  - ctrlProps and the VML spell some object types differently ("CheckBox" vs "Checkbox") and the
    label must be nested <div><font> in the VML textbox or LibreOffice does not find it. Both were
    caught by round-tripping a CREATED control rather than only a hand-built one.
  - VBA: sheetedit reads, runs and writes macros now, through the vbalang library plus the Excel
    object model in `vba-excel.ts`. See `_plans/done/VBA_PLAN.md`. A macro compiled by our own
    writer runs in LibreOffice from a button we created, which is the one check a parser cannot
    make: it proves a module with no p-code cache in front of it still compiles elsewhere. Note
    XLSX cannot hold VBA at all; macros require XLSM/XLSB/XLS.
- **Undo covers the sheet-level settings** (protection, page setup, panes, outline grouping) and the
  workbook theme, not just cell edits. A cell edit records its own fields; these live on the sheet,
  so the step carries a before/after snapshot instead, via the history's existing undoExtra /
  redoExtra closures. Two things that bite:
  - Maps and sets must be CLONED into the snapshot, or undo hands back the very object it is meant
    to restore and the next edit corrupts it.
  - The dirty flags are FORCED true on restore rather than snapshotted. After a save the flag is
    clear; restoring that stale value would leave the undone change missing from the next save.
  - A theme switch is its own inverse: cells keep their theme references, so re-resolving against
    the previous palette restores every colour it changed.
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
    A job covers the active sheet, every sheet, or a selection (transient, never committed to the
    sheet's print area). Each range of a multi-range print area starts its own page, as Excel does.
    Page numbers run through the whole job. A browser applies ONE @page size per job, so a job whose
    sheets disagree on paper is flagged (mixedPaper) and uses the first sheet's rather than
    silently resizing the rest.
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
