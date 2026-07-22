# Charts in sheetedit: view, create, and save (xlsx + ods)

Goal: sheetedit reads the charts already in a workbook and renders them live over the grid, lets
you create new charts from a data range through an interface, and writes them back so both Excel
(xlsx DrawingML) and LibreOffice/OpenDocument (ods embedded chart object) open and re-edit them
natively. Rendering uses Chart.js (MIT); everything else (parsing, the create UI, and the two
writers) is ours.

## The shape of the problem

A chart touches three layers, and the two file formats disagree on all of them:

- **xlsx**: a chart is `xl/charts/chartN.xml` (DrawingML `c:chartSpace`), anchored by
  `xl/drawings/drawingN.xml` (a `twoCellAnchor` + `graphicFrame` referencing the chart), wired
  into the sheet by a `<drawing r:id>` in `sheetN.xml` and rels, and declared in
  `[Content_Types].xml`. Data ranges are formula refs (`Sheet1!$B$2:$B$10`) with cached values.
- **ods**: a chart is a whole embedded OpenDocument sub-document in its own zip folder
  (`Object 1/content.xml` with `chart:chart`, plus its `styles.xml`, `meta.xml`), listed in
  `META-INF/manifest.xml`, and referenced from the sheet's `content.xml` by a `draw:frame` +
  `draw:object xlink:href="./Object 1"`. Data is a `chart:series` with
  `chart:values-cell-range-address` (and usually an internal fallback data table).

So we need a **normalized chart model** that both readers produce, both writers consume, the
Chart.js renderer maps from, and the create UI builds. That model is the spine of the feature.

### Normalized model (`src/core/chart-model.ts`)

```ts
interface ChartModel {
  id: string;
  kind: "column" | "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "bubble" | "radar";
  stacked?: boolean;
  title?: string;
  legend?: { show: boolean; pos: "top" | "bottom" | "left" | "right" };
  categories?: Ref;                 // shared x labels (not for scatter/bubble/pie-per-series)
  series: {
    name?: Ref | string;
    values: Ref;                    // y values; for scatter/bubble also xValues/sizes
    xValues?: Ref; sizes?: Ref;
    color?: string;
  }[];
  axes?: { x?: AxisOpt; y?: AxisOpt };
  anchor: { fromCol: number; fromRow: number; fromColOff: number; fromRowOff: number;
            toCol: number; toRow: number; toColOff: number; toRowOff: number }; // EMU-ish, grid coords
  original?: { parts: string[] };   // for a chart we read but have not edited (see write policy)
  dirty?: boolean;                  // created or edited -> written from the model
}
type Ref = { ref?: string; cache?: (string | number)[] }; // live sheet range, cache fallback
```

`sheet.charts?: ChartModel[]` holds them per sheet.

## Rendering (Chart.js over the grid)

- Chart.js is a **lazy import**, loaded only when a workbook has (or gains) a chart, like
  mlang/localml, so the base bundle is unaffected.
- Each chart is a floating canvas positioned from its anchor and the grid geometry (column
  widths, row heights, scroll, frozen panes) that the editor already owns. An overlay layer sized
  to the scroll content holds the chart canvases; they reposition on scroll / resize / column
  drag, and clip when scrolled out of view.
- Data is resolved from the **live sheet** through `parseA1Ref` + the cell model, with the cached
  values as a fallback, so a cell edit re-renders the chart (recompute the Chart.js dataset, no
  file rewrite).
- `kind`/`stacked` map onto Chart.js types: column/bar -> `bar` (horizontal via `indexAxis`),
  line, area -> `line` with `fill`, pie/doughnut, scatter, bubble, radar. Legend, title, stacking
  and a secondary y-axis map to Chart.js options.

## Create / edit interface

- An **Insert chart** toolbar button (grid toolbar, xlsx/ods only) opens a dialog seeded from the
  current selection: data range, a chart-type picker (icon grid), and options (title, legend,
  "first row/column is labels"). A live Chart.js preview updates as options change. Insert builds
  a `ChartModel` anchored near the selection, pushes it to `sheet.charts`, overlays it, marks
  dirty.
- **Select a chart** (click) shows move/resize handles (drag updates the anchor) and a small
  panel: change type, edit title, edit data range, toggle legend, delete. Every edit sets
  `dirty` and re-renders.

## Write policy (preserve unless touched)

Mirrors sheetedit's cell philosophy:

- A chart the user **only viewed** keeps its original parts byte-for-byte; the writers skip it, so
  exotic chart features we do not model survive the round trip, and Excel re-renders it from the
  (possibly edited) cells anyway.
- A chart that was **created or edited** (`dirty`) is written from the `ChartModel`, replacing or
  adding its parts. We accept that a heavily-featured original, once edited in our UI, is
  re-emitted as the subset we model (documented limitation).

## xlsx writer (`src/adapters/xlsx/chart-write.ts`)

For each dirty chart: emit `xl/charts/chartN.xml` (per-kind DrawingML template), ensure a
`xl/drawings/drawingN.xml` with the `twoCellAnchor` + `graphicFrame`, add the sheet
`<drawing r:id>` + `worksheets/_rels/sheetN.xml.rels` entry, the drawing->chart rel, and the
`[Content_Types].xml` overrides (`drawingml.chart+xml`, `drawing+xml`). Reuses the part-
registration approach from `sheet-create.ts`. Round-trip through our own reader; validate the
emitted XML against the ECMA-376 chart schema in CI.

## ods writer (`src/adapters/ods/chart-write.ts`)

For each dirty chart: create the embedded object folder (`Object N/content.xml` with
`chart:chart` + an internal data table, plus minimal `styles.xml`), add its `META-INF/manifest.xml`
entries, and insert a `draw:frame` + `draw:object` into the sheet `content.xml` at the anchor.
Keep `mimetype` first/stored (already handled by writeOds). Round-trip through our reader;
validate against the ODF schema where feasible.

## Readers

- `src/adapters/xlsx/chart-read.ts`: resolve chart parts via the sheet -> drawing -> chart rels,
  parse `c:chartSpace` (kind, series `c:ser` with `c:cat`/`c:val` refs + caches, title, legend,
  axes) and the anchor -> `ChartModel`.
- `src/adapters/ods/chart-read.ts`: walk `draw:frame`/`draw:object`, load the referenced embedded
  `content.xml`, parse `chart:chart` -> `ChartModel`.

## Phases (each independently shippable)

0. **Spike / de-risk.** Read ONE real bar chart from an xlsx, map to Chart.js, overlay it at its
   anchor, and keep it glued to the cells while scrolling and resizing a column. Proves the three
   riskiest pieces at once (DrawingML parse, the overlay/anchor math, Chart.js mapping).
1. **View xlsx charts.** chart-read.ts + the overlay layer + live data + the common kinds. Read
   only; nothing written.
2. **View ods charts.** chart-read.ts for the embedded object -> same model -> same renderer.
   Viewing now works for both formats.
3. **Create + edit UI.** Insert dialog (range/type/options/live preview), move/resize/delete,
   change-type/title/range. Produces/edits `ChartModel`; no file writing yet (charts exist in the
   session).
4. **xlsx writer.** Persist created/edited charts to DrawingML; reopen-in-our-reader round trip +
   schema validation. This is where a created chart first survives save/reload.
5. **ods writer.** Persist to the embedded chart object + manifest; round trip + validation.
6. **Polish.** More kinds (combo, stock-ish), data labels, dual axis, axis number formats, series
   colors/theme, i18n, responsive (charts in the mobile layout), Cypress round trip.

## Verification (important: no Excel available)

The make-or-break for the writers is that Excel and LibreOffice open the result without a
"repair" prompt, and there is no Excel here. Strategy: (a) round-trip through sheetedit's own
reader every time (structural correctness); (b) validate emitted parts against the committed
ECMA-376 (xlsx) and ODF (ods) schemas in CI; (c) where a headless LibreOffice is available, open
the written ods/xlsx and assert no error. Residual risk (real-Excel acceptance) is called out per
phase, same honesty as the Power Query load-to note.

## Open decisions (for review)

1. **Home: in sheetedit, or a standalone chart-codec lib?** The model <-> DrawingML/ODS codec is
   reusable and could be its own lib (family pattern), but the overlay, data resolution and UI are
   sheetedit-bound. Recommend building in sheetedit first; extract the pure codec later only if it
   earns it (avoid premature split).
2. **v1 kind coverage.** Recommend column/bar (+stacked), line (+stacked), area, pie, doughnut,
   scatter, bubble, radar. Defer surface, stock, waterfall, treemap, box-and-whisker (no native
   Chart.js type).
3. **How much create-UI in v1.** Recommend insert + move/resize/delete + change type/title/range;
   defer deep formatting (per-point colors, effects, trendlines).
4. **Verification depth for writers** given no Excel: schema validation + our-reader round trip as
   the CI gate, headless LibreOffice for ods if installable. Confirm that is acceptable.

## Risks / honesty

- **Two full write paths.** DrawingML and the ODS embedded object share almost nothing; the
  writers are the bulk of the effort, and getting a file that opens cleanly in real apps is the
  delicate part (gated behind the strongest checks we can run locally).
- **Anchoring over a virtualized + frozen + variable-size grid** is fiddly; the spike de-risks it.
- **Chart.js fidelity is approximate**, not pixel-identical to Excel/LibreOffice; exotic types and
  deep formatting are out of v1.
- **Edited charts lose un-modeled features** (write policy above) - preserved only while untouched.
- **Bundle size**: Chart.js is lazy-loaded so it never weighs on non-chart workbooks.

## Progress

- (DONE) Phase 0/1: xlsx chart reader (chart-read.ts) + Chart.js overlay (chart-overlay.ts) glued
  to the grid, scroll-synced, live data from cells. Model in core/chart-model.ts, Sheet.charts.
  Browser-verified: grouped column chart renders, scrolls with cells, and updates on a cell edit.
- (DONE) Phase 2: ods chart reader (adapters/ods/chart-read.ts) - embedded object -> same model -> same overlay; browser-verified
- (DONE) Phase 3: create/edit UI - insert dialog (type/range/options/live preview), select + edit toolbar (change type, delete), move/resize handles; chart-build.ts + chart-insert.ts. Browser-verified.
- (DONE) Phase 4: xlsx writer (adapters/xlsx/chart-write.ts) - DrawingML chart part + anchor/drawing/rels/content-type registration for created charts; rewrite part + anchor for edited ones. Round-trip unit-tested through the reader; NOT real-Excel-verified.
- (DONE) Phase 5: ods writer (adapters/ods/chart-write.ts) - embedded Object N/ (chart:chart + internal data table) + manifest + draw:frame in content.xml. Round-trip unit-tested; anchor absolute-approximate (default-width sheets).
- (pending) Phase 6: polish
