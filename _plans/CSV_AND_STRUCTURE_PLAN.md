# sheetedit: CSV/TSV support + richdoc-style structure

Goal: open .csv/.tsv in the real spreadsheet editor (grid, formula bar, keyboard,
undo, row/column ops) with every unsupported control removed and a "Convert to
XLSX" button instead; and reorganize the source like richdoc (per-format adapter
directories with read/write split) for long-term maintainability.

## Phase 1: structure refactor (no behavior change)

Mirror richdoc's core/ + adapters/ layout:

```
src/core/        model.ts  recalc.ts  structure.ts  history.ts  workbook.ts  i18n.ts  editor.ts
src/core/ui/     toolbar.ts  formulabar.ts  floatbar.ts
src/adapters/xlsx/  read.ts  write.ts  styles.ts  index.ts
src/adapters/ods/   read.ts  write.ts  styles.ts  shared.ts  index.ts
src/adapters/csv/   (new, Phase 2)
src/index.ts     unchanged public re-export surface
```

Split rules:
- xlsx.ts (760 lines) -> read.ts (readXlsx + cell/style-pool parsing),
  write.ts (writeXlsx, writeXlsxCell, col-width/row-height/merge setters),
  styles.ts (setXlsxCellStyle + pool helpers). mergeCellStyle (shared with ods)
  moves to core.
- ods.ts (700 lines) -> read.ts (readOds, readOdsTable, parseOdsRow, style
  parsing), write.ts (writeOds, makeOdsCell, setters), shared.ts (ODS
  namespaces, odfToA1 / a1ToOdf, length/color helpers).
- Pure moves + import fixes, one commit. index.ts keeps exporting the same
  names, so omnitext, the tests and the demo compile unchanged. Suite must be
  green with zero test-logic edits (only import paths in test files).

## Phase 2: CSV/TSV adapter in sheetedit

### Model and round-trip

- Workbook.kind gains "csv". One sheet, delimiter stored on the workbook
  (sniffed comma/semicolon/tab/pipe, same quote-aware consistency scorer as
  omnitext's csv sniffer; a .tsv filename hint forces tab).
- Span preservation, like omnitext's CSV model: the reader keeps each physical
  row's raw text + terminator. The writer emits untouched rows byte-for-byte
  and re-serializes only rows whose cells were edited (quoting only where
  needed). Consequences:
  - open-then-save is byte-identical;
  - "007" or "1.10" in untouched rows can never be coerced (edited cells keep
    the exact typed text: sheetedit stores values as strings already).
- Structure ops (insert/delete rows/cols) reuse structure.ts; affected rows
  are marked dirty, reference rewriting works as in xlsx/ods since formulas
  are plain A1.
- readWorkbook(bytes): not a zip / not CFB -> UTF-8 decode -> csv path.
  createSheetEditor gains options.formatHint?: "csv" | "tsv" so hosts can
  route explicitly (extension beats sniffing).

### Formulas

CSV cells starting with "=" are read as formulas, recalculated by the existing
engine, and saved back as formula text. That matches how Excel and LibreOffice
treat CSVs. (Alternative considered: treat "=" as literal text. Rejected
unless you prefer it: it would show stale text where other apps compute.)

### UI in csv mode

- styled=false already hides bold/italic/colours/align/borders/merge; also
  hide the sheet tab bar (single sheet) and the number-format-dependent bits.
  Kept: undo/redo, + Row / + Col, row/column header menus, formula bar,
  fill/keyboard/copy-paste behaviors.
- New toolbar button "Convert to XLSX" (en/fr/ja strings): builds a minimal
  real workbook from the model (values, formulas, column widths) via a small
  adapters/csv/convert.ts, then hands it to options.onConvert?(bytes, name).
  Without a host callback it downloads name.xlsx (standalone demo behavior).

### Tests

Sniffer cases; byte-exact unedited round-trip (quoting quirks, CRLF, ragged
rows, no trailing newline); edited-row-only rewriting; formula recalc + save;
structure ops on csv; convert-to-xlsx output reopens as a valid xlsx with the
same values/formulas; editor-level: csv mode shows no style cluster and shows
the convert button.

## Phase 3: omnitext integration

- csv/tsv formats offer the "sheet" editor and it becomes their default view;
  the existing Table and Text views stay available in the switcher (the table
  editor also still serves .xls).
- The sheet adapter bridges text formats: mount = serialize model -> UTF-8
  bytes -> createSheetEditor(bytes, { formatHint }); save = a new synchronous
  editor.getText() (csv only) so omnitext's text pipeline (encoding menu, .gz,
  history diffs) keeps working untouched.
- Convert to XLSX inside omnitext: onConvert opens the produced bytes as a new
  unsaved document named <base>.xlsx (window-level event the shell listens
  for), leaving the original CSV document as-is.
- Live verification on the built app, then audit/shipped-log updates.

## Out of scope (unchanged limits)

- Render cap (5000 rows) applies to CSV like everything else.
- .xls stays on the lossy SheetJS table path.
- Multi-sheet CSV concepts do not exist; the tab bar stays hidden.

## Decisions (approved 2026-07-08)

1. sheetedit is the default view for csv/tsv in omnitext (Table/Text remain).
2. CSV formulas are computed and saved as formula text.
3. Convert to XLSX opens the result as a new unsaved document.
