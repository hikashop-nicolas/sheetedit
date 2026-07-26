# Plans

Two kinds of document live here, and the split is what the `done/` folder is for.

## Living (this folder)

Kept current as the code changes. Read these to know where sheetedit stands.

| | |
|---|---|
| `SHEETEDIT_GAPS_PLAN.md` | The main record. Its phases are all shipped; the "Remaining / not yet done" section at the end is the honest backlog, and every entry carries the gotchas and verification caveats that came with it. |
| `XLSX_COVERAGE.md` | What the `.xlsx` path reads, edits and preserves, part by part. |
| `ODS_COVERAGE.md` | The same for `.ods`. |

## Done (`done/`)

Finished pieces of work, kept for the reasoning rather than the status. Each one records why a
thing was built the way it was, and what bit on the way, which is the part worth keeping once the
checkboxes are all ticked.

| | |
|---|---|
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
