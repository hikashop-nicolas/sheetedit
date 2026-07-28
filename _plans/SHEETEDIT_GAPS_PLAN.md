# sheetedit: the living record

What sheetedit does, what is honestly still missing, and the rules the work follows. Finished
plans live in `done/`; the part-by-part audits are `XLSX_COVERAGE.md` and `ODS_COVERAGE.md`.

Status (2026-07-28): the nine gap-closing phases are shipped (`done/PHASES_1_9.md`), and so are
charts, pivot tables, Power Query, macros, and ActiveX down to its container controls
(`done/ACTIVEX_PLAN.md`). What follows is what remains.

## Still open

Short, because most of what used to be here shipped. Each is stated as what a user would notice.

- **A part-of-cell hyperlink flattens on edit (.ods)**: ODF can anchor a link to part of a cell's
  text, or carry several in one cell. They survive untouched and survive an edit that leaves the
  text alone, but editing that cell's value drops them, since the text they were anchored to is
  gone. The grid also shows and follows only the first link of such a cell.
- **A second note per cell outlives sheetedit but not LibreOffice (.ods)**: every annotation is
  kept through an edit, and LibreOffice keeps one per cell and drops the rest on re-save.
- **An edited shared-string cell is rewritten inline (.xlsx)**: its `sharedStrings` entry can
  become an unreferenced orphan. The file is valid; it is slightly larger than Excel would write.
- **Power Query load-to-new-sheet** writes a real refreshable ListObject with its connection now,
  but whether Excel refreshes it on demand is untested here.
- **Pivot layout is not byte-identical to Excel's**, which neither app preserves anyway: both
  re-flow the body on open.

## What shipped, and what it cost

The record worth keeping: what each area does, the gotchas found on the way, and how far the
verification actually goes.

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

## Irreducible, and why

- **Third-party ActiveX**: OOXML says such a control's content is determined by the control itself,
  so its format belongs to whoever wrote it. Preserved untouched.
- **ODS graphical conditional formatting authoring**: colour scales, data bars and icon sets have
  no interoperable ODF form; they read from LibreOffice files and are written back as they came.
- **No Excel here**: slicers, timelines, the ActiveX writers and authored pivots are spec-faithful
  and verified by round-trip, by LibreOffice where it can judge, by reading our own output back,
  and now by the ECMA-376 schemas themselves (`npm run check:schema`). Excel itself remains the
  untested case, and every claim about it says so.

  What the schema check adds, and what it does not: it validates the parts the schemas describe
  (worksheets, workbook, styles, shared strings, tables, pivots, connections) and compares against
  the input, so only a violation we introduce is reported. It does not cover the parts that live in
  extensions - slicers, timelines, the x14/x15 registrations - which have no published XSD in that
  set, so those stay verified by round-trip and by asserting each registration is present.

## Working notes
- Reuse existing machinery: the chart overlay layer pattern (Phase 2), hidden-rows (Phase 1),
  the DV dropdown + CF renderer + comment popover (Phases 3-4), createWorksheet/schema-sync for
  new parts. Prefer extending src/adapters/xlsx/*.ts modules already present (hyperlink.ts,
  datavalidation-ish, condformat.ts, comments, tables.ts) over new subsystems.
- Every format-touching change: round-trip test + LibreOffice open; every visible change:
  browser-verify. Bump omnitext after each phase.
