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
- Written as the **native** structure in both formats, with `refreshOnLoad` set so Excel and
  LibreOffice re-render the pivot from the source on open; verified through LibreOffice round-trips.

Caveats: "show values as", calculated fields and calculated items are honoured by Excel (standard
OOXML) and by sheetedit's own display, but LibreOffice's xlsx pivot rebuild ignores them (it shows
the raw aggregate). Calculated items are emitted per the OOXML spec but could not be verified in
Excel here, only that the file still opens cleanly in LibreOffice. Byte-identical layout to Excel is
not attempted (both apps re-flow the body from the definition on open).

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
