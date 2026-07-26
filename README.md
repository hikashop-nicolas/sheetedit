# sheetedit

A standalone, framework-agnostic, client-side **spreadsheet editor** for `.xlsx`
(OOXML) and `.ods` (OpenDocument). Both are zips of XML; sheetedit reads the cells
into an editable grid, **preserves formulas and recalculates** them as you edit, and
exports a valid workbook, **keeping styles, number formats, charts, pivot tables and other
sheets** intact. No server, no upload.

It also **creates and edits charts and pivot tables**, and opens and edits a workbook's
**Power Query** definitions in a built-in, Excel-style query editor, refreshing them on-device
(see [Charts](#charts), [Pivot tables](#pivot-tables), [Power Query](#power-query)).

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
engine), both MIT. The Power Query engine
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

A dropdown's items come from its source range, so it stays in step with the sheet. A **button** runs
a macro, which a browser cannot and should not execute, so it is drawn disabled with a tooltip
saying why rather than looking live and doing nothing. Labels and group boxes are drawn as-is.

Controls can be **created and edited**, not only read. The toolbar's control button inserts a
checkbox, dropdown or spinner at the selection and opens its settings; the same dialog edits an
existing control's label, linked cell and item source, or deletes it. Creating one builds all three
parts Excel expects (the `ctrlProps` part, a VML shape, the worksheet entry) plus the content types
and relationships that make them findable, since missing any one of them makes Excel drop the
control silently.

State lives in two places in an `.xlsx` and both are read and written: the modern `ctrlProps` part
and the legacy VML drawing that positions the control. Files predating `ctrlProps` carry everything
in the VML, so it is the fallback rather than an afterthought, and a state change is mirrored into
both so an older reader sees it too. Everything else in either part is left untouched.

**Macros and ActiveX are preserved, not run.** A `.xlsm`'s `vbaProject.bin` and any ActiveX part
survive a save byte-for-byte, because untouched parts are never rebuilt. Running them is a different
proposition: it needs a VBA interpreter *and* the Excel object model, since macros exist to
manipulate the workbook. That is a project in its own right, not a missing flag, and a partial
implementation would be worse than none: a macro that half-runs can leave a workbook in a state its
author never intended.

`.ods` keeps controls in `office:forms` with a `draw:control` frame, a different model that is
preserved rather than rendered. Verified against LibreOffice, which reads the kind, label, linked
cell and source range back from what we write; its own import does not carry a checkbox's checked
state (a file that was never checked reads back as checked), so that part is verified by sheetedit's
own round-trip rather than through it.

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
- Not a full spreadsheet application (no drawing/shape editing, no macros). This is a lightweight,
  embeddable in-browser editor for cell, formula, chart, pivot and query content.
- Power Query editing is `.xlsx`-only, and a query result loaded onto a brand-new sheet is
  written as plain cells (not a live, Excel-refreshable table). Loading onto an existing
  destination table refreshes it in place.

## Develop

```
npm install
npm run dev       # standalone demo (open a workbook, edit, download)
npm run build     # compile the library to dist/ (tsc)
npm test          # vitest round-trip + recalc tests (jsdom)
npm run test:e2e  # Cypress end-to-end tests (Chrome) against the built demo
```

Regenerate the e2e fixtures with `node cypress/gen-fixture.mjs`.

License: MIT.
