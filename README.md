# sheetedit

A standalone, framework-agnostic, client-side **spreadsheet editor** for `.xlsx`
(OOXML) and `.ods` (OpenDocument). Both are zips of XML; sheetedit reads the cells
into an editable grid, **preserves formulas and recalculates** them as you edit, and
exports a valid workbook, **keeping styles, number formats, charts, pivot tables and other
sheets** intact. No server, no upload.

It also **creates and edits charts and pivot tables**, **runs and edits a workbook's VBA macros**,
and opens and edits its **Power Query** definitions in a built-in, Excel-style query editor,
refreshing them on-device (see [Charts](#charts), [Pivot tables](#pivot-tables),
[Macros](#macros-vba), [Power Query](#power-query)).

**[▶ Live demo](https://hikashop-nicolas.github.io/sheetedit/)** - open a `.xlsx` or
`.ods`, edit cells and formulas, and download the result, entirely in your browser.

```ts
import { createSheetEditor } from "sheetedit";

const editor = createSheetEditor(containerEl, workbookBytes, {
  onChange: () => console.log("edited"),
});

// later, to save:
const editedBytes = await editor.getBytes(); // a valid .xlsx or .ods
```

Runtime dependencies: [`fflate`](https://github.com/101arrowz/fflate) (zip) and
[`fast-formula-parser`](https://github.com/LesterLyu/fast-formula-parser) (formula
engine) and [`vbalang`](https://github.com/hikashop-nicolas/vbalang) (the VBA engine,
extracted from this project), all MIT. The Power Query engine
([`mlang`](https://github.com/hikashop-nicolas/mlang)) and the on-device formula
assistant ([`localml`](https://github.com/hikashop-nicolas/localml)) are lazy-loaded
only when their features are used, so they stay out of the base bundle.

## Formulas and recalculation

- Type `=B2*2` or `=SUM(C2:C3)` into a cell; the result is computed and shown.
- When a cell changes, every formula that depends on it (directly or transitively) is
  recomputed in dependency order, across sheets.
- Formula cells display their computed value, and the formula itself when focused.
- Existing formulas in the file are preserved on save with refreshed cached values.

## Number formats

Cells are shown using their number format: dates, currency, percentages and
thousands separators render the way the file specifies, while the cell still edits as
its raw underlying value (focus a `$1,234.50` cell and you edit `1234.5`). For `.xlsx`
the format code is resolved from `xl/styles.xml` and applied with SSF (the SheetJS
formatter, bundled with the formula engine); for `.ods` the producer's formatted text
in `<text:p>` is used. A typed value keeps the cell's format, and a recomputed formula
result is re-formatted.

You can also use the pure functions directly:

```ts
import { readWorkbook, recalc, setCellInput, writeWorkbook } from "sheetedit";

const wb = readWorkbook(bytes);
setCellInput(wb.sheets[0], 2, 2, "5"); // row 2, col 2 (B2)
recalc(wb);
const out = writeWorkbook(wb); // re-zips, preserving other parts
```

## Power Query

`.xlsx` workbooks made with Power Query embed their query definitions (the M `Section1.m`
inside the DataMashup blob). sheetedit reads them and gives you a full **Power Query
Editor**, entirely in the browser:

- **Applied Steps + live preview** - a query is a `let` expression, so its steps are listed
  like Excel's; selecting a step evaluates the query up to that point and shows the result
  (row-capped) in a preview grid with per-column type and quality bars.
- **Transform ribbon** - GUI transforms that generate M and append a step: choose/remove/
  rename columns, filter, sort, keep/remove rows, remove duplicates, change type, replace
  values, split column, group by, unpivot, transpose, add custom/index columns, and
  merge/append with other queries.
- **Get Data** - create a new query from a workbook table, a web CSV/JSON URL, or blank.
- **Load** - write a query's result into its destination table, or onto a new sheet.
- **Refresh** - a quick panel refreshes existing queries into their tables, on open or on
  demand. The formula bar and a raw-M view let you hand-edit steps.

Evaluation is a **clean-room M engine** ([`mlang`](https://github.com/hikashop-nicolas/mlang),
~640 standard-library functions), lazy-loaded so the base editor bundle never carries it.
Workbook data is served through `Excel.CurrentWorkbook`; browser-reachable connectors
(`Web.Contents`, `OData.Feed`, `File.Contents`) work, while queries that read databases or
other native sources can only be refreshed in Excel and are reported as such. Power Query is
`.xlsx`-only (OpenDocument has no equivalent payload). The editor is fully responsive: on
narrow screens the ribbon goes icon-only and the side panels become drawers.

### Loading a query's result

Loading a query writes its result as a **real Excel table** named after the query, with the
connection Excel recognises it by, rather than as loose cells. That is what makes the result a
thing rather than a copy of some values: structured references reach it, and loading the same query
again finds that table and refreshes it in place instead of adding another sheet.

## Charts

sheetedit renders the charts already in a workbook (xlsx DrawingML and ods embedded chart
objects) as live [Chart.js](https://www.chartjs.org/) canvases floating over the grid, anchored
to their cells and updated as you edit the underlying data. You can also **create** charts: the
Insert-chart button opens a dialog (type, data range, options, live preview), and a selected
chart can be moved, resized, retyped or deleted. Created and edited charts are **saved back** to
both formats (DrawingML for xlsx, an embedded chart object for ods) so Excel and LibreOffice open
and re-edit them natively; charts you only view round-trip untouched. Chart.js is loaded lazily,
only when a workbook has (or gains) a chart.

Covered types: column, bar, line, area, pie, doughnut, scatter, bubble and radar. Not covered
(rendered approximately or skipped): surface, stock, waterfall and other types with no Chart.js
equivalent, plus deep formatting (trendlines, per-point styling). Chart rendering is an
approximation, not pixel-identical to Excel/LibreOffice.

## Pivot tables

sheetedit reads the pivot tables in a workbook (xlsx `pivotTable` + `pivotCache`, ods
`data-pilot-table`), outlines each one on the grid, and lets you **create, edit and refresh** them:

- **Insert** opens a dialog: pick the source range, drag each column into Rows / Columns / Values
  (with an aggregation) / Report Filter, toggle subtotals, all with a **live preview** of the
  result, then it lands on a new sheet.
- **Fields**: any depth of nested row and column fields, one or more value fields
  (sum / count / average / min / max), report/page filters, and per-group subtotals.
- **Show values as**: express a value field as a % of the grand / column / row total, or a running
  total, per field.
- **Calculated fields**: add a value field defined by a formula over the source columns' sums
  (e.g. `Revenue - Cost`).
- **Calculated items**: add a synthetic member to a row/column field from a formula over that
  field's items (e.g. on Region: `North + South`); it appears as an extra row/column.
- **Pivot chart**: from the tag menu, chart the pivot's output (grand totals excluded); the chart is
  bound to the output cells, so it updates when the pivot recomputes.
- **Edit / refresh** from the pivot's tag menu: refresh recomputes the output from the current
  source; edit reopens the dialog prefilled and rewrites the pivot in place. This works both for
  pivots you create and for pivots opened from a file (the definition is reconstructed on read).
- **Slicers** (xlsx): the slicers in a workbook are rendered as real panels over the grid. Click an
  item to narrow to it, ctrl/cmd-click to multi-select, ⊗ to clear; the linked pivots recompute and
  the new selection is written back into the slicer cache on save. Table slicers filter their
  table's rows instead. "Insert slicer" on a pivot's tag menu creates one for any field the pivot
  groups by, and the toolbar's slicer button creates one for a column of the table under the
  cursor. Slicer styles are honoured, including workbook-defined custom ones.
- **Timelines** (xlsx): the date sibling of a slicer, shown as a strip of periods at the timeline's
  granularity (years / quarters / months / days). Click a period to filter the linked pivots to it,
  shift-click to extend the range, ⊗ to clear; the range is written back into the timeline cache.
- Written as the **native** structure in both formats, with `refreshOnLoad` set so Excel and
  LibreOffice re-render the pivot from the source on open; verified through LibreOffice round-trips.

Caveats: "show values as", calculated fields and calculated items are honoured by Excel (standard
OOXML) and by sheetedit's own display, but LibreOffice's xlsx pivot rebuild ignores them (it shows
the raw aggregate). Calculated items are emitted per the OOXML spec but could not be verified in
Excel here, only that the file still opens cleanly in LibreOffice. Byte-identical layout to Excel is
not attempted (both apps re-flow the body from the definition on open).

## Freeze and split panes

Pin a header row and/or the first columns so they stay put while the rest of the sheet scrolls. The
toolbar's freeze button offers **freeze at this cell**, **freeze top row**, **freeze first column**,
**split at this cell** and **unfreeze**; the row and column header menus also carry **freeze rows
above** / **freeze columns to the left**.

Either kind puts a **divider** on the boundary: drag it to move the split, double-click it to remove
that one. Freeze and split render the same way here (the leading pane stays put); the difference is
what the file records, and a split stays a split where the format can say so.

Existing boundaries are read and rendered, and a change is written back: `.xlsx` as
`<pane state="frozen">` (line counts) or `<pane state="split">` (an offset in twips) in the sheet
view, `.ods` into the view settings in `settings.xml` (created, with its manifest entry, when the
file has none).

A **split gives you real viewports**: a row split makes two, a column split makes two side by side,
and a split on both axes makes Excel's four. Each pane scrolls the axis its boundary cuts and shares
the other with its neighbours, so panes in the same row band scroll vertically together and panes in
the same column band scroll horizontally together. The column header is drawn by the top band and the
row numbers by the left one. A freeze stays a single viewport with the leading lines pinned.

Floating objects (charts, images, shapes, slicers, timelines and the pivot tags) follow the split:
each one is drawn in whichever pane currently shows the row it is anchored to.

Limitations worth knowing: an object straddling the boundary is drawn once,
in the pane showing more of its anchor row, so it is clipped there rather than continuing across.
The outline gutter is drawn for the upper pane. And on
`.ods`, a split you *move* is written back as a frozen boundary, because ODF states a split's
position in a LibreOffice view-pixel unit that could not be pinned down here; an untouched split
round-trips unchanged.

## Outline grouping

Row and column groups are read, rendered and authored in both formats. Grouped rows get an
Excel-style gutter left of the row numbers: a bar per level, a **+/-** button on each group's summary
line, and **1 / 2 / 3** buttons that show one depth and collapse everything deeper. **Group**,
**Ungroup**, **Collapse**, **Expand** and **Clear outline** are also in the row and column header
context menus, which is how column groups are managed.

Collapsing hides the group's lines and marks its summary line, exactly as the file records it, so the
state round-trips: `.xlsx` writes `outlineLevel` / `collapsed` / `hidden` per row and a `<col>` span
per column, `.ods` nests the rows in `<table:table-row-group>` (with `table:display="false"` for a
collapsed one). Verified through LibreOffice round-trips both ways. ODS **column** groups are read
and preserved but not authored.

## Undo

Ctrl+Z covers cell edits and, since they are just as easy to trigger by accident, the sheet-level
settings too: protecting a sheet, page setup and print area, freeze and split panes, outline
grouping, and a workbook theme switch. A theme switch is exactly reversible because the cells keep
their references to the palette rather than the colours it resolved to.

Row and column insert/delete clear the history, because they shift every recorded position and
steps from before them cannot replay safely.

## Form controls

The checkboxes, dropdowns, spinners and buttons a workbook can put over the grid are read, drawn and
made to work. Their point is the **linked cell**: a checkbox writes TRUE/FALSE into it, a dropdown
writes the 1-based index of the chosen item, a spinner writes its number, and formulas read that
cell. So ticking a box on a rendered sheet recalculates it, the way it would in Excel.

A dropdown's items come from its source range, so it stays in step with the sheet. A **button runs
the macro it names** (see [Macros](#macros-vba)); one with nothing assigned is drawn disabled and
says so. **Radio buttons** clear the rest of their group box, as they must to be radios at all.
Labels and group boxes are drawn as-is.

Controls can be **created, edited, moved and resized**, not only read. The toolbar's control button
inserts a checkbox, dropdown or spinner at the selection and opens its settings; the same dialog
edits an existing control's label, linked cell, item source and assigned macro, or deletes it. Drag
the grip on a control's left edge to move it and the corner handle to resize it; the face keeps its
own clicks, so a checkbox stays tickable. Creating one builds all three parts Excel expects (the
`ctrlProps` part, a VML shape, the worksheet entry) plus the content types and relationships that
make them findable, since missing any one of them makes Excel drop the control silently.

Files written by older Excel carry **no worksheet `<controls>` element at all**, only VML shapes.
Those are read as controls too, or the buttons in such a file would be invisible.

**ActiveX controls are read and drawn as well.** Their state lives in a persisted binary beside the
sheet, which is not opaque: [MS-OFORMS] specifies it, so a Forms 2.0 control gives up its kind,
caption, value, size, colours and (for a scroll bar or spin button) its bounds, on the same
published-spec basis as everything else here. Command buttons, check boxes, option buttons, text
boxes, combo and list boxes, toggle buttons, labels, scroll bars and spin buttons are all read. A button runs the
handler its name implies (`CommandButton1` runs `CommandButton1_Click` from the sheet's own code
module). A control with a **linked cell** is live: an ActiveX combo lists the items its
`listFillRange` names, resolving a defined name, and writing the chosen **text** to that cell,
which is where the two control families differ (a form control writes the item's position). A control's own persisted state is
written back too, so a checkbox toggled here is still toggled when the file is reopened.

A control's **caption can be set** as well as its value: double-click a button or a label to name
it. That is offered only where the write is known to work, which is tested by rewriting the caption
the control already has, so an affordance never appears where saving would fail. A multi-column list
is drawn as the grid Excel draws, at **the per-column widths the file states**.

**Frame, MultiPage and TabStrip are read, drawn and written too.** These are a different shape from
the rest: a container control persists as a *storage*, so its binary is a compound file holding its
own properties, a table of what it contains, an object stream with every child's properties, and
(for a MultiPage) a storage per page. The container is drawn as the captioned box it is, with its
children at their recorded positions, and changing one writes back into that storage. The option
buttons in a container are one group, as they are in Excel, so choosing one clears the rest. A
TabStrip is not a container at all - it has tabs, not children - and shows its tabs with the one
the file selects.

That writer is deliberately conservative. Where the new text is the same length as the old, the
bytes are patched **in place**, so the stream stays byte-identical down to its padding; only a
change of length rebuilds the block, carrying the parts sheetedit does not model (picture streams,
font properties) across untouched. Writing into a container also corrects the container's own
bookkeeping, since its children sit end to end and one that changes length moves the rest. It
refuses outright on a control whose layout the reader would not vouch for, and reads its own output
back before returning it: a write must not proceed where a read would not.

**Verified by round-trip, not by a second implementation.** Writing an unchanged value back returns
the original bytes exactly, on every real stream tested, and a changed one reads back correctly
through the whole save. But LibreOffice does not surface ActiveX from `.xlsx` at all, so unlike the
VBA writer, which an independent engine could be made to *run*, there is no outside judge here. It
opens a rewritten file without complaint, and that is all that proves. A third-party ActiveX control is genuinely opaque,
since OOXML says its content is determined by the control itself, and it is preserved untouched.

State lives in two places in an `.xlsx` and both are read and written: the modern `ctrlProps` part
and the legacy VML drawing that positions the control. Files predating `ctrlProps` carry everything
in the VML, so it is the fallback rather than an afterthought, and a state change is mirrored into
both so an older reader sees it too. Everything else in either part is left untouched.

`.ods` keeps controls in `office:forms` with a `draw:control` frame, a different model that is
preserved rather than rendered. Verified against LibreOffice, which reads the kind, label, linked
cell and source range back from what we write; its own import does not carry a checkbox's checked
state (a file that was never checked reads back as checked), so that part is verified by sheetedit's
own round-trip rather than through it.

## Macros (VBA)

A `.xlsm`'s macros are **read, run and edited**, through
[`vbalang`](https://github.com/hikashop-nicolas/vbalang) (the VBA engine, extracted from this
project) plus an Excel object model that maps onto sheetedit's own workbook. Nothing reaches outside
the page: a browser tab is where it all happens.

The toolbar's macro button lists the modules and their source, with a **Run** button for each
procedure that can run on its own. The source box is **editable**, and saving it compiles the text
back into `vbaProject.bin`, verified by reading it back before it is accepted. A **button on the
sheet** runs the macro its `<x:FmlaMacro>` names, which is how the workbook's author meant it to be
run.

`Workbook_Open` and `Worksheet_Change` run **only if you say so**, from a checkbox in the macro
dialog that lasts for the session and is never persisted. A workbook that runs code the moment it
opens is the whole reason Excel grew its own "enable content" bar.

Two rules the whole feature is built on:

- **A run is one undo step**, including the parts that touch no cell (hidden rows, a rename,
  protection). One undo puts the workbook back.
- **Refuse rather than approximate.** Anything unmodelled stops the run naming what was missing,
  rather than evaluating to `Empty` and carrying on. A macro that half-runs leaves a workbook in a
  state its author never intended, and the user then saves it. That covers both what cannot be done
  in a page (`Shell`, `CreateObject`, `Application.Quit`, `Workbook.SaveAs`) and what simply is not
  built yet. A step budget stops a runaway loop from hanging the tab, and `MsgBox` output is shown
  after the run instead of blocking on a dialog.

What the object model covers: `Range` (values, formulas, `Offset`/`Resize`/`Cells`, `Find`, `Sort`,
`AutoFilter`, `SpecialCells`, `Copy`/`Cut`/`PasteSpecial`, `Replace`, `RemoveDuplicates`,
`TextToColumns`, `AdvancedFilter`, `Clear`, `Font`/`Interior`), `Worksheet` and `Worksheets`
(including `Visible`, `Add`, `Copy`, `Move`, `Delete`, `Protect`), `Workbook`, and `Application`
with `WorksheetFunction` delegating to sheetedit's own formula engine rather than growing a second
implementation of `SUM`. A `Range` can cover several areas, which is what makes `SpecialCells` and
`Union` behave.

**Verified in LibreOffice, not in Excel.** A workbook whose macro source was compiled by sheetedit's
own writer opens and *runs* in LibreOffice, driven from a button sheetedit created: the macro reads
cells, computes, and writes results back, accented text included. That matters more than a parse
check, because rewriting source means dropping the compiled p-code cache (so the workbook can never
show one macro and run another), and only an engine that compiles the source can say whether that
still works. There is no Excel on the machine this was built on, and Excel is stricter, so this
raises the confidence rather than settling it.

## Hidden sheets

A sheet the workbook hides draws no tab, in either format, and can be hidden or shown from the tab
menu. The last visible sheet cannot be hidden. `Worksheet.Visible` works from a macro too, with
Excel's three states including "very hidden", which is the one only a macro can reach.

ODF keeps this somewhere other than where it looks like it should: in the sheet's **table style**
(`<style:table-properties table:display>`), not on `<table:table>`. The attribute written on the
element parses fine and LibreOffice ignores it entirely.

## Sorting, filtering and the rest

Beyond the sections above, sheetedit also reads, renders and authors: **sort and value filters** on
an autofilter range, **conditional formatting** (every rule kind, icon sets included),
**data validation** (list dropdowns plus the whole/decimal/date/time/length/custom family, with the
rule's own help and error messages, shown in the grid as the cell's tooltip),
**rich text** runs within a cell, **sparklines**, **images** (move, resize, replace),
**drawing shapes**, **slicers** and **timelines** (interactive, and written back), **hyperlinks**,
**comments**, **furigana**, and **dynamic arrays** (the Excel 365 spilling family: `UNIQUE`, `SORT`,
`FILTER`, `SEQUENCE`, `TEXTSPLIT` and the rest).

Both formats author every conditional-format kind, which for `.ods` means writing LibreOffice's
`calcext` extension. That works, contrary to a long-standing note here saying it did not: two
placement rules decide it and both fail silently. The block must sit INSIDE `<table:table>`, and the
fill it applies must be a named style in `styles.xml`. Written anywhere else, LibreOffice ignores
the rule while sheetedit's own reader still finds it, so a round-trip test passes and the file looks
right here and blank there.

## Workbook themes

A workbook's theme is a named palette of twelve colours plus a heading and a body font. Cells do not
store a theme colour, they store a **reference** to one, which is why switching the theme recolours
everything that used it while leaving cells given an explicit colour completely alone.

sheetedit reads the palette and the scheme fonts from `xl/theme/theme1.xml`, keeps each reference
next to the colour it resolved to, and offers the built-in palettes (Office, Office 2007-2010,
Berlin, Slice, Ion, Grayscale) behind the toolbar's theme button. A file whose own theme is not one
of those keeps it at the top of the list, so switching away and back loses nothing.

Switching rewrites **only** `theme1.xml`; `styles.xml` is untouched, exactly as Excel does it, and
the part is patched in place so `<a:fmtScheme>` (the effect and fill styles behind chart and shape
presets, which sheetedit does not model) survives. A workbook that shipped without a theme gets a
complete one, registered in the content types and the workbook relationships so other readers
actually find it. Verified through LibreOffice: it resolves the switched cells to the new palette.

`.ods` has no equivalent palette that cells reference by index, so this is xlsx-only.

## Dark and light

The widget ships both. Colours resolve through `--sheetedit-*` custom properties, defaulting to
light and following `prefers-color-scheme` for dark, with `data-theme="light"` or `"dark"` on the
root element overriding the OS. That is the same signal Omnitext sets, so an embedded editor flips
with its host and needs no wiring.

A host can also redefine any subset of the properties on any ancestor to restyle the editor
directly. The grid follows the mode too: a paper-white grid inside a dark app is the thing that
actually looks broken.

## Print settings

sheetedit prints, and stores the settings for whatever else opens the file: Excel, LibreOffice, or a
PDF export. Two parts of the setup are visible on the grid, which is what makes the rest checkable:
the **print area** is outlined in green, and **page breaks** are drawn as blue dashed lines that
continue into the row and column headers.

The toolbar's printer button holds **Print**, **Page setup** (orientation, paper size, scaling or
fit-to-width, margins, print gridlines and headings, centring, and the header/footer regions), plus
**Set** / **Clear print area** from the selection and **Reset page breaks**. The row and column
header menus carry **Insert** / **Remove page break** and **Repeat these rows (or columns) on every
page**.

**Print** lays the print area out as pages and hands them to the browser. It asks what the job
covers first: the **active sheet**, **all sheets**, or the **selection only** (offered when there is
one; it never touches the sheet's stored print area). The pagination is real, not a screenshot of
the grid: columns and rows are split by the paper size and margins, manual breaks are honoured,
fit-to-width shrinks (never enlarges), title rows and columns repeat on every page, and the pages
come out in the order the setup asks for. A print area with several ranges puts each range on its
own page, as Excel does, and fit-to sizes from the widest so the ranges stay comparable. Each page carries its
own header and footer with the field codes resolved, so `&P of &N` numbers correctly and
`firstPageNumber` is respected. Merged cells span, hidden rows and columns are skipped, and each
cell keeps its own formatting.

Pages are emitted at exactly the paper size with `@page { margin: 0 }`, so the margins, header and
footer are placed from the file's own values rather than the browser's. Page numbers run through the
whole job, so `&N` is the job's total rather than one sheet's.

Two things outside our control: the browser's own header and footer strip (the URL and date), which
you turn off in the print dialog for an exact page; and the fact that a browser applies **one**
`@page` size to a whole job, so printing sheets that disagree on paper uses the first sheet's and
says so rather than silently resizing the rest. Printing always uses light colours, whatever mode
the editor is in.

Header and footer text keeps the file's own field codes (`&P` page, `&N` total pages, `&D` date,
`&T` time, `&A` sheet, `&F` file), so a header sheetedit did not author round-trips exactly,
formatting codes included.

`.xlsx` writes `<printOptions>`, `<pageMargins>`, `<pageSetup>`, `<headerFooter>` and the row/column
break lists, with the print area and repeated rows as the sheet-scoped `_xlnm.Print_Area` and
`_xlnm.Print_Titles` names. `.ods` writes a `<style:page-layout>` plus `<style:master-page>` pair in
`styles.xml` (reached through the table's own style), `table:print-ranges`, a
`<table:table-header-rows>` group, and `fo:break-before="page"` on the row or column style.

Two conversions needed care, and both were pinned down against LibreOffice rather than guessed:

- **Margins mean different things.** An xlsx top margin encloses the header block; an ODF page margin
  stops where that block starts. The adapters convert both ways, so a 0.75in top margin with a 0.3in
  header becomes a 0.3in ODF page margin over a 0.45in header block, and back.
- **LibreOffice derives an OOXML margin from the header's content height plus its own spacing,
  ignoring `fo:min-height`**, and defaults that spacing to 20mm when absent, which inflated the
  margin by about half an inch on every `.ods` conversion. Writing the spacing the way LibreOffice
  writes it in its own files makes the margin survive exactly.

## Sheet and workbook protection

Protection is read, **enforced in the grid** and authored, in both formats. On a protected sheet a
locked cell is selectable and copyable but refuses typing, and every write path that could go around
that (paste, fill handle, clear, insert/delete rows and columns, sort, formatting, and adding,
deleting, renaming or moving sheets) is gated too. A refused action shows a short notice rather than
doing nothing silently.

Both formats default every cell to **locked**, so protection alone freezes the whole sheet: unlocking
the input range is what makes it usable. The toolbar's lock button offers **Protect sheet** (with the
allowances: select locked cells, format, insert / delete rows and columns, sort, autofilter),
**Unprotect sheet**, **Lock** / **Unlock cells** for the selection, and **Protect workbook
structure**. Cell locking is only offered while the sheet is unprotected, as in Excel.

**This is not a security feature.** Neither format encrypts anything; they store at most a password
*hash*, and any spreadsheet tool can lift the flag. sheetedit therefore never computes or checks a
password. A hash already in the file is preserved verbatim when you re-save, and removing protection
drops it without asking for it, which the dialog states plainly.

`.xlsx` writes `<sheetProtection>` (only the flags that differ from their spec defaults, as Excel
does), `<workbookProtection lockStructure="1"/>`, and per-cell `<protection locked="0"/>` in the style
pool. `.ods` writes `table:protected` with LibreOffice's `<loext:table-protection>` permission flags,
`table:structure-protected`, and `style:cell-protect`. The two are inverses (OOXML names what is
blocked, ODF what is allowed) and the adapters convert. ODF has no equivalent for the format / sort /
autofilter / pivot flags, so those fall back to their blocked defaults on an `.ods` round-trip.
Verified through LibreOffice both ways; LibreOffice itself drops workbook structure protection on its
OOXML export, so that flag was checked within ODF only.

## How preservation works

- **`.xlsx`**: only the `<c>` cell elements you changed are rewritten in the
  worksheet DOM; everything else (styles, number formats, merges, charts, other
  sheets, untouched cells) is left byte-for-byte. New string cells are written as
  inline strings, so the shared-string table is never disturbed.
- **`.ods`**: the table body is regenerated, but untouched cells are cloned verbatim
  (dates, currency, formats survive), column definitions are kept, and every other
  part of the archive is preserved. The `mimetype` entry stays first and stored, as
  ODF requires.

## Scope / honest limitations

- Edits cell values and formulas across all sheets. Styles are preserved; number
  formats are both preserved and applied to the displayed value (see Number formats).
- The formula engine implements a large subset of spreadsheet functions; an
  unsupported function or a circular reference yields an error value in that cell.
  Desktop apps recompute on open, so cached values are a convenience, not authority.
- `.ods` formulas typed in the grid are translated from A1 to ODF syntax on save; the
  common arithmetic, range and function cases are handled.
- Not a full spreadsheet application, but the gap is narrower than it was: shapes and macros are
  both authored now. What it is not is a desktop suite.
- **Nothing here is verified against Excel itself**, which is not available on the machine this was
  built on. Authored pivots, slicers, timelines and rewritten macro projects are checked against the
  published specs, sheetedit's own round-trips, and LibreOffice, which drops slicers entirely and
  ignores calculated pivot fields on rebuild. Where a caveat matters it is stated in the section
  concerned rather than left implied.
- Power Query editing is `.xlsx`-only. A query loaded onto a new sheet is written as a real table
  with its connection, and loading it again refreshes that table in place; whether Excel refreshes
  it on demand is untested here, for want of an Excel.

## Develop

```
npm install
npm run dev       # standalone demo (open a workbook, edit, download)
npm run build     # compile the library to dist/ (tsc)
npm test          # vitest round-trip + recalc tests (jsdom)
npm run test:e2e  # Cypress end-to-end tests (Chrome) against the built demo
npm run typecheck # the native TypeScript compiler, about a second on the whole project
npm run check:schema   # validate what the writer emits against the ECMA-376 schemas
npm run check:openpyxl # have an independent reader confirm an authored workbook
```

**What `check:schema` actually asks.** Not "is this file schema-perfect", because real workbooks
are not: Excel writes `xml:space` on `<t>`, which the schema forbids outright, and
`mc:AlternateContent` inside worksheets, which the worksheet schema does not mention. Run a
validator over any genuine file and it complains. So the check reads each demo workbook, edits a
cell, writes it back, and reports only the complaints the **original did not already draw** - the
ones this project introduced. It needs `xmllint` and downloads the ECMA-376 Part 4 (transitional)
schemas once into `.cache/`, which is not committed: they are ECMA's, not ours.

**What `check:openpyxl` adds.** Round-trips prove sheetedit agrees with itself; LibreOffice judges
what it supports and drops slicers entirely; the schemas judge structure but not meaning. So a
workbook is authored with a table, two conditional-format rules, a data validation with all four of
its messages, a hyperlink, a comment, a frozen pane and a merge, and **openpyxl** - a separate
implementation, in another language, sharing no code with this one - is asked whether it reads back
what was meant. It needs `pip install openpyxl`.

Both checks earned their place immediately: the schema one found an attribute written where it was
not needed, and the openpyxl one found that the corpus was setting model fields directly instead of
going through the authoring APIs, so the file it produced had neither the freeze nor the merge in
it.

Regenerate the e2e fixtures with `node cypress/gen-fixture.mjs`.

**Two TypeScripts, on purpose.** `typecheck` runs TypeScript 7, the native compiler, aliased as
`tsgo` and called by path. `typescript` itself stays on 6 because TypeScript 7's package is a
launcher for a binary and exposes no JS compiler API, which Cypress needs to compile its specs:
with 7 under that name, every e2e spec fails to bundle. Keeping both gives the fast typecheck and a
working test suite.

License: MIT.
