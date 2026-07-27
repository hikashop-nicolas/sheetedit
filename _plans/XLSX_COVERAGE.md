# sheetedit xlsx coverage report

Audit of what the `.xlsx` path reads, edits and preserves, as of 2026-07-27. Tracks the gap list and
the priority order for closing it. The companion documents are `ODS_COVERAGE.md` for the ODF side and
`SHEETEDIT_GAPS_PLAN.md` for the phase-by-phase record of how each area got here.

## Fully supported (read, edit in the grid, save)

### Cells and formulas
- Cell content: text, number, boolean, error; inline and shared strings.
- Formulas + recalc: a large fast-formula-parser subset, recomputed in dependency order across
  sheets; legacy array formulas preserved; shared-formula groups de-shared safely on edit.
- The library ships many common functions as empty stubs; those are filled in `core/functions.ts`
  and `core/financial.ts` (statistics, SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/MINIFS, MATCH / XLOOKUP /
  XMATCH / CHOOSE / LOOKUP, SUBTOTAL / AGGREGATE, text, SWITCH, the financial family), so INDEX+MATCH
  and the other everyday idioms recalculate. OFFSET / INDIRECT return real references
  (`core/reference-fns.ts`) and LET is expanded before parsing (`core/let-expand.ts`).
- Dynamic arrays: a formula returning a 2-D result spills into its anchor + range, with a #SPILL!
  guard on collisions. The whole shaping family is supplied, since the parser has none of it:
  UNIQUE / SORT / SORTBY / FILTER / SEQUENCE / TRANSPOSE / RANDARRAY / TAKE / DROP / CHOOSEROWS /
  CHOOSECOLS / EXPAND / HSTACK / VSTACK / TOROW / TOCOL / WRAPROWS / WRAPCOLS / TEXTSPLIT, plus bare
  range refs. FILTER needs an array mask (no range=scalar broadcasting).
- Structured (table) references: `Table1[Units]`, `[[#Headers],[Units]]`, `[@Units]`, `[[A]:[B]]`,
  `#All` / `#Data`, and a bare `Table1`.
- Number formats resolved from styles.xml and applied to the display; a typed value keeps its format.

### Presentation
- Cell styling: bold, italic, underline (+ flavour), strikethrough, font family/size, text colour,
  fill colour, horizontal and vertical alignment, wrap, per-side borders. Furigana (Japanese ruby).
- Rich text within one cell: each run's own style renders, and is authored from the UI (select part
  of a cell's text while editing, then bold/italic/underline/strike/colour/size/font). Retyping a
  cell's text clears its runs, since the offsets would shift.
- Text spill: a text too long for its column runs over the empty cells beside it and stops at the
  first that holds anything, direction following the alignment. Numbers do not spill; rich text does.
- Conditional formatting: full authoring and rendering, every rule kind. dxf / colour scales / data
  bars / icon sets, is-true-formula (expression) rules, `cellIs` cell-ref and formula operands, and
  time-period rules.
- Workbook themes read (palette + scheme fonts) and switchable.
- Grid metrics come from the workbook rather than from constants: a column's `width` is in character
  units of the NORMAL STYLE's font, so the conversion uses that font's maximum digit width, and
  `<sheetFormatPr defaultRowHeight/defaultColWidth>` is honoured.

### Structure and view
- Insert/delete rows and columns (rewriting formula and merge references), column-width and
  row-height resize, merge/unmerge, fill down/right, find and replace across sheets.
- Sheet management: add / rename / delete / reorder; hidden sheets read, honoured and authored.
- Frozen panes and SPLIT panes: read, rendered, authored and written. A split is two real viewports
  (a 2x2 of quadrants when both axes are split), and every floating overlay follows the split.
- Outline grouping (rows and columns): read, rendered as an Excel-style gutter with +/- and level
  buttons, authored and written.
- Hidden rows/columns, and autofilter: interactive sort and per-column value filtering, persisted.
- Print settings: page setup, margins, header/footer, print area, repeated rows and page breaks, read,
  authored and drawn on the grid.
- Protection: sheet and workbook, read, enforced in the grid (locked cells become read-only), authored.
- Undo covers the sheet-level settings (protection, page setup, panes, outline grouping) as well as
  cell edits.

### Objects on the grid
- Charts: read, render (Chart.js + custom plugins), full create/edit dialog and write for every
  DrawingML chart type, all option tiers, pseudo-3D, xlsx + ods round-trip.
- Pivot tables: read, outline on the grid, create/edit/refresh, nested row/column fields, multiple
  value fields, page filters, subtotals, "show values as", calculated fields, calculated items, and a
  pivot chart.
- Images: rendered from `xl/media/*` and the drawing anchors; move (drag), resize (corner handle) and
  replace (double-click) are authored and written back to the anchor / media part.
- Shapes: read, rendered as an SVG overlay with fill / outline / centered text, and authored (insert,
  move, resize, edit, delete). A shape's `<xdr:style>` theme reference resolves through the theme's
  fill and line style lists, gradients included, rendered as real SVG gradients. A shape carrying a
  `macro` attribute is a button and runs it.
- Windows metafiles (WMF / EMF) render: a metafile is a recorded list of GDI drawing calls rather
  than an image, so it is replayed onto a canvas (emf-converter, lazy-loaded, failing soft). Covers
  sheet images and an ActiveX control's Picture.
- Sparklines: authored and rendered.
- Slicers: read, rendered as a real slicer panel, interactive (click / ctrl-click / clear), written
  back, and authorable. Pivot and TABLE slicers both; OLAP slicers render read-only.
- Timelines: read, rendered, interactive date filtering, written back.
- Comments and notes (legacy + threaded): corner marker + hover popover, authored.
- Hyperlinks: read, rendered, clickable, authored. Data-validation dropdowns: all rule types author,
  read and validate; list rules show a picker.

### Controls and macros
- Form controls: read, rendered and interactive; the linked cell is the point, so a checkbox writes
  its state where the file says.
- ActiveX (Forms 2.0): read from the binary through an [MS-OFORMS] parser, rendered as live controls
  and WRITTEN back. Covers CommandButton, the MorphData family (checkbox, combo, text, list, option,
  toggle), Label, Image, ScrollBar and SpinButton. Properties honoured include VariousPropertyBits
  (Enabled, Locked, BackStyle, ColumnHeads, MatchRequired, Alignment, Editable, WordWrap, AutoSize,
  MultiLine), DisplayStyle (the only thing separating an editable combo from a drop-list, since they
  share a class id), the numeric family (MaxLength, PasswordChar, BorderStyle/Color, SpecialEffect,
  ScrollBars, ListRows, ListWidth, ColumnCount, BoundColumn, TextColumn, MultiSelect, MatchEntry,
  ListStyle, ShowDropButtonWhen, DropButtonStyle, MousePointer, Accelerator, PicturePosition,
  SmallChange, LargeChange, Orientation, Delay, ProportionalThumb), TextProps (the font), and
  StreamData pictures. Writing covers every string a control carries and can ADD a property the
  control does not yet have, for the MorphData family AND the flat layouts, so an unlabelled
  button or label can be given a caption (double-click it). Multi-column lists render as a grid, at the per-column widths the file states (rgColumnInfo).
- VBA macros: read, run and written (vbalang). `Worksheet.OLEObjects` and `ListObjects` are on the
  object model, and `ListObjects.Add` writes the whole package.
- Power Query: read, full editor (Applied Steps, preview, transform ribbon, Get Data, merge/append),
  refresh, and Load; can author the first query in a query-less workbook.

## Preserved on save, but inert in the grid

Survive because untouched parts are kept byte-for-byte and the worksheet DOM is re-serialized with
its sibling elements intact:

- Third-party (non Forms 2.0) ActiveX controls. Irreducible: OOXML says such a control's content
  "shall be solely determined by the corresponding object", so the format belongs to whoever wrote it.
- Drawing types not modelled above (SmartArt, WordArt, embedded OLE objects other than ActiveX).
- Defined names (read for recalc, not user-editable).
- External data connections other than Power Query.

## Gaps / not handled

- **ActiveX Frame / MultiPage / TabStrip**: a missing STRUCTURE rather than a missing property. The
  control's stream holds a whole embedded form (a ClassTable, a sites array, a child stream per
  control), which is the parent-controls half of MS-OFORMS, about the size of all the leaf-control
  work. They only reach a worksheet through "More Controls".
- **System-palette OLE colours** are left unset on purpose: such a colour is the desktop theme's,
  not the document's.
- **Pivot layout**: byte-identical layout to Excel is not attempted (both apps re-flow on open).
- **Recalc is a subset**: unsupported functions or circular refs yield an error value (the cached
  value is shown as a fallback); exotic custom number-format codes may render slightly differently.

## Correctness caveats

- Editing a shared-string cell rewrites it as an inline string (its sharedStrings entry can become
  unreferenced).
- Power Query load-to-new-sheet writes plain cells, not a live refreshable ListObject, and is verified
  only through sheetedit's own reader.
- Show-values-as / calculated fields / calculated items are honoured by Excel and by sheetedit's own
  display, but ignored by LibreOffice's xlsx pivot rebuild. The calculated-item OOXML is emitted per
  spec but unverified in Excel.
- **No Excel here.** Slicers, timelines and the ActiveX writer follow the specification and are
  verified in-app, by round-trip, and by asserting each registration is present. LibreOffice drops
  slicers and does not surface ActiveX from xlsx, so for those there is no outside judge: it reopening
  a rewritten workbook proves only that the package is not corrupt.
- LibreOffice's headless converter drops per-sheet view settings on every output it writes, so frozen
  and split panes cannot be round-tripped through it (what an xlsx pass does prove is in the plan).

## Priority order for closing the remaining gaps

The original six-item list (sheet management, hyperlinks, data validation, range shifting,
conditional formatting, comments) is DONE in full; the current order is:

1. Frame / MultiPage / TabStrip, if a real file ever needs it. This is the parent-controls half of
   MS-OFORMS: a ClassTable, a sites array and a child stream per control, roughly the size of all
   the leaf-control work, for controls that only reach a worksheet through "More Controls".

DONE since this list was written: the ODS parity gaps, Power Query load-to-new-sheet as a real
ListObject, the per-column widths of a multi-column ActiveX list (rgColumnInfo), and ActiveX
caption authoring.
