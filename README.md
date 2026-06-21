# sheetedit

A standalone, framework-agnostic, client-side **spreadsheet editor** for `.xlsx`
(OOXML) and `.ods` (OpenDocument). Both are zips of XML; sheetedit reads the cells
into an editable grid, **preserves formulas and recalculates** them as you edit, and
exports a valid workbook, **keeping styles, number formats, charts and other sheets**
intact. No server, no upload.

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
engine), both MIT.

## Formulas and recalculation

- Type `=B2*2` or `=SUM(C2:C3)` into a cell; the result is computed and shown.
- When a cell changes, every formula that depends on it (directly or transitively) is
  recomputed in dependency order, across sheets.
- Formula cells display their computed value, and the formula itself when focused.
- Existing formulas in the file are preserved on save with refreshed cached values.

You can also use the pure functions directly:

```ts
import { readWorkbook, recalc, setCellInput, writeWorkbook } from "sheetedit";

const wb = readWorkbook(bytes);
setCellInput(wb.sheets[0], 2, 2, "5"); // row 2, col 2 (B2)
recalc(wb);
const out = writeWorkbook(wb); // re-zips, preserving other parts
```

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

- Edits cell values and formulas across all sheets. Styles and number formats are
  preserved, not displayed: a date stored as a serial number shows the raw number in
  the grid (and is preserved on save).
- The formula engine implements a large subset of spreadsheet functions; an
  unsupported function or a circular reference yields an error value in that cell.
  Desktop apps recompute on open, so cached values are a convenience, not authority.
- `.ods` formulas typed in the grid are translated from A1 to ODF syntax on save; the
  common arithmetic, range and function cases are handled.
- Not a full spreadsheet application (no charts editing, pivot tables, formatting UI).
  This is a lightweight, embeddable in-browser editor for cell and formula content.

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
