# Plans

Two kinds of document live here, and the split is what the `done/` folder is for.

## Living (this folder)

Kept current as the code changes. Read these to know where sheetedit stands.

| | |
|---|---|
| `SHEETEDIT_GAPS_PLAN.md` | The main record. "Still open" is the honest backlog, "What shipped, and what it cost" is the reasoning behind each area, and every entry carries the gotchas and verification caveats that came with it. |
| `XLSX_COVERAGE.md` | What the `.xlsx` path reads, edits and preserves, part by part. |
| `ODS_COVERAGE.md` | The same for `.ods`, including which half of a file states what. |

## Done (`done/`)

Finished pieces of work, kept for the reasoning rather than the status. Each one records why a
thing was built the way it was, and what bit on the way, which is the part worth keeping once the
checkboxes are all ticked.

| | |
|---|---|
| `PHASES_1_9.md` | The nine gap-closing phases: sort/filter, images, hyperlinks and comments, conditional formatting, sparklines, rich text, dynamic arrays, correctness fixes, pivot tables. |
| `ACTIVEX_PLAN.md` | ActiveX end to end: the MS-OFORMS parser, the property audit, per-column widths, caption authoring, and the container controls (Frame / MultiPage / TabStrip). |
| `VBA_PLAN.md` | Reading, running and writing macros. Stages 0-5, all shipped; the engine now lives in the vbalang library. |
| `CHARTS_PLAN.md` | Chart reading, the Chart.js overlay, authoring, and both writers. |
| `CHARTS_SPEC_GAPS.md` | Closing the gap to the DrawingML chart schema. |
| `PIVOT_AUTHORING.md` | Pivot tables from read-only to authoring. |
| `POWERQUERY_UI_PLAN.md` | The Power Query editor UI on top of `mlang/steps`. |
| `CSV_AND_STRUCTURE_PLAN.md` | CSV/TSV support and the adapter structure refactor. |
| `STYLES_PLAN.md` | Cell styles: display, then editing, then write-back. |

## House rules, which every one of them followed

Dependency-light and framework-agnostic. In-place surgical XML edits, so untouched parts stay
byte-for-byte. Every format-touching change gets a round-trip test and a LibreOffice open; every
visible change gets browser-verified; then bump omnitext.

Two rules that came from being wrong, and are worth stating plainly:

- **Read the specification, do not infer it.** Three separate "no form exists" claims in these
  documents turned out to be false, each from not finding something rather than from checking. The
  ActiveX column-width record was in the downloadable spec but absent from the HTML index; a
  time-period conditional format has its own calcext element; validation messages exist in both
  formats and were carried by neither.
- **A writer reads its own output back** before returning it, and refuses any input the reader
  would not vouch for. A refusal leaves the file alone, which is always better than half a write.
