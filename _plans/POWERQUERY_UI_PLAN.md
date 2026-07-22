# Power Query editor UI in sheetedit

Goal: turn today's read/refresh/edit-raw-M panel into an Excel-style **Power Query Editor**:
create and build queries with a live preview grid, an Applied Steps pipeline, and GUI
transforms that generate M, then Load the result into the workbook. The mlang engine is
complete (~640 functions, section evaluation, qdeff read/write); this plan is almost entirely
**UI + a thin step-level API on mlang**. No new evaluator semantics.

Reference for the target UX: Excel's Power Query Editor (queries pane, preview grid, Query
Settings with Applied Steps, the Home/Transform/Add-Column ribbon, the formula bar, Close &
Load). We reproduce its shape, not its chrome.

## What exists today (done)

- `src/core/ui/query-panel.ts`: lists section members, shows/edits the whole `Section1.m`
  (highlighted textarea), refreshes a query into its existing destination table, auto-refresh
  on load, file attach for `File.Contents`, host connectors (Excel.CurrentWorkbook, Web.*,
  OData.Feed, File/Folder).
- `src/adapters/xlsx/tables.ts`: list workbook tables, match a query to its table, write a
  result back resizing the table part (`applyQueryResult`).
- mlang: `evaluateSection(sectionM, host)` gives member names + `run(name)`; `mlang/qdeff`
  reads/writes the DataMashup blob.

## The key insight (makes this tractable)

A Power Query is a `let` expression. **Applied Steps = the ordered `let` bindings**; the final
`in` names the last step. "Preview at step N" = evaluate the member with its `in` re-pointed to
step N. A GUI transform = append one binding `#"New Step" = Table.Fn(#"Prev Step", args)` and
re-point `in`. mlang already evaluates `let` and returns tables, so the editor is: parse a
member into steps, re-point/append/remove bindings (string or AST level), and re-run for the
preview. Every common transform maps to a `Table.*` function mlang already ships.

## Cross-lib addition: a step-level API on mlang

mlang owns the parser (@microsoft/powerquery-parser) and the evaluator, so step manipulation
belongs there, not duplicated in sheetedit. Proposed new module `mlang/steps` (lazy, like
`mlang/qdeff`):

- `parseMemberSteps(sectionM, name): { steps: {name, expression}[], tail: string }` - decompose
  a member's `let` into ordered steps (a non-`let` member is a single implicit step).
- `previewExpression(sectionM, name, uptoStep?): string` - the M to evaluate for the preview at
  a step (re-points `in`); sheetedit runs it through the existing host bindings and grabs the
  table. A row cap is applied by the caller (wrap in `Table.FirstN`).
- `withStep(...)` editors, all returning a new `sectionM`: `appendStep`, `insertStep`,
  `removeStep`, `renameStep`, `replaceStepExpression`, `reorderStep`, and `addMember` /
  `removeMember` / `renameMember` for the query list. These edit the AST and re-serialize (the
  qdeff writer already round-trips), preserving the rest of the section.

This is the only engine-side work. It is mechanical (AST in, M out) and oracle-neutral (no new
evaluation behaviour), so it needs unit tests but no PQTest fixtures.

## sheetedit UI: new module `src/core/ui/pq-editor/`

A full-window overlay (Excel opens a separate window; we use an in-app modal filling the
viewport, themed by the same `--sheetedit-*` variables). Layout:

- **Left - Queries pane**: every section member; select to edit; context actions add / rename /
  duplicate / reference / delete / enable-load. "New query" opens Get Data.
- **Center - Preview grid**: the selected step's result as a read-only table. A small dedicated
  MTable renderer (virtualized, reusing `virtual.ts` math), not the editing grid. Column
  headers show name + type; footer shows "N columns, M preview rows (capped at K)".
- **Right - Query Settings**: query name + **Applied Steps** list (from `parseMemberSteps`).
  Click a step to preview at that point; right-click to rename / delete / move / edit formula;
  a gear on a GUI-generated step reopens its dialog.
- **Top - transform ribbon** (grouped, each button appends a step and refreshes the preview).
- **Formula bar**: the selected step's M, editable; commit re-parses and re-previews (reuses
  `m-highlight.ts`).
- **Close & Load / Load To**: run the query and write the result in (see Phase C).

mlang and the pq-editor module are lazy-imported; the base editor bundle is unchanged.

## Status (2026-07-23)

Phases A, B and C are DONE and shipped (sheetedit 2770335 / 9fd034a / 7102e35), on top of the
`mlang/steps` API (mlang 6399889). Browser-verified on demo/pq-sales.xlsx. Remaining: Phase D
polish and Load-To a brand-new sheet/table (deferred, needs an Excel-reopen test). See the
per-phase notes below.

## Phases (each independently shippable)

**Phase A (DONE) - Editor shell + Applied Steps + live preview.** The overlay, queries pane, preview
grid renderer, Applied Steps from `parseMemberSteps`, formula bar. Selecting a step or editing
its formula re-previews through the existing host bindings (preview row-capped). No transform
buttons yet, but you can already read any query as a stepped pipeline and hand-edit steps with
live results. Delivers most of the perceived "Power Query" value on top of what exists.

**Phase B (DONE) - Transform actions (M generators).** The ribbon. Each button opens a tiny dialog and
appends a step via the `mlang/steps` editors, all mapping to shipped functions:
Choose/Remove Columns, Filter Rows, Sort, Change Type, Rename Columns, Keep/Remove Top/Bottom
N, Remove Duplicates, Use First Row as Headers, Replace Values, Split Column, Group By, Add
Custom Column (raw M), Add Index Column, Add Conditional Column, Unpivot/Pivot. Plus step
delete / rename / reorder. This is the bulk of the work but each transform is small and
independently testable (generate M, assert the preview).

**Phase C (DONE, except new-destination Load-To) - Get Data + Load To.** Create new queries: from a workbook table/range
(`Excel.CurrentWorkbook`), from an attached CSV/JSON/Excel file, from a Web URL, or blank. Load
To: an existing table (reuse `applyQueryResult`) or a **new sheet + table** (extend `tables.ts`
to create a destination table part and register the query-to-table link). "Connection only"
(no load) is allowed. This is where a fresh query first reaches the grid.

**Phase D (remaining) - polish.** Merge/Append queries (reference other members), lightweight column
profiling (row count, type, empty/error quality bar), query dependency view, `.ods` parity
(ODF has no DataMashup, so queries are xlsx-only for now; the panel hides for `.ods`), full
i18n, and a Cypress round-trip: build a query in the UI, Load, save, reopen (and refresh) in
Excel.

## Open decisions (for review, not blockers)

1. **v1 scope.** Recommend shipping **A + B** first (build and edit queries over data already in
   the workbook, with live preview and the core transforms), then C (new sources / load to a new
   table), then D. This puts "mlang properly used" in the user's hands fastest without the
   trickier new-destination write-back.
2. **Full-window overlay vs expanding the current panel.** Recommend the full-window overlay:
   three panes + ribbon + preview do not fit a popover, and it matches Excel's mental model.
3. **Preview cost.** The replay-sync engine re-runs the `let` per step selection. Fine for
   typical sheets; cap the preview (e.g. 1000 rows, like Excel) and only re-run on step change
   or an explicit "refresh preview", not per keystroke.
4. **Where step decomposition lives.** Recommend the `mlang/steps` module (mlang owns the AST)
   over a second parser dependency in sheetedit.

## Risks / honesty

- **Biggest UI in the family.** A + B is multi-week at a steady pace; the per-transform
  structure keeps it shippable incrementally (each transform is a self-contained unit).
- **Step round-trip fidelity.** Editing the AST and re-serializing must preserve untouched
  steps byte-for-reason (comments, `#"quoted"` names, `in` target). Unit-test every editor
  against real members, and keep the raw-M textarea as an escape hatch.
- **New-destination write-back (Phase C).** Creating a sheet + table part that Excel still
  accepts and can refresh is the delicate part, same class of risk as the existing resize
  write-back; gate it behind the Excel-reopen Cypress test.
- **Preview vs load divergence.** Preview is row-capped and load is not; make the cap explicit
  in the UI so a user never thinks the capped preview is the full result.
