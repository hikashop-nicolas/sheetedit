# Charts: full spec coverage (DrawingML + ODF)

Goal: close the gap between our chart support and the ECMA-376 DrawingML chart schema (Part 1,
§21.2) and the ODF chart schema, for read + render + write round-trip. Renderer is Chart.js
(some features need its plugins or a custom draw); the write policy stays "preserve an untouched
chart verbatim, re-emit an edited/created chart from our model", so every feature we model is one
less thing lost when a chart is edited.

## Tiers

### Tier 1 - quick wins (this pass)

- [x] Axis min/max: read c:valAx/c:catAx c:scaling c:max/c:min -> model.axes; apply in Chart.js scales.
- [x] dispBlanksAs (gap|zero|span): read chart-level; plot nulls as gaps/zeros/spanned; write it back.
- [x] Doughnut holeSize, bar gapWidth + overlap: read/apply (cutout, categoryPercentage/barPercentage)/write.
- [x] percentStacked: true 100% stacking (normalise per-category to %) on render; grouping on write.
- [x] ODS bar-vs-column + stacked/percentage: read the plot-area chart-properties style (chart:vertical,
      chart:stacked, chart:percentage) so ODS bar charts and stacked charts render right.
- [x] Per-point colours (c:dPt): read per-point spPr srgbClr onto series.pointColors; use on render;
      write c:dPt. (Makes read pie/bar charts match the file's colours.)

### Tier 2 - moderate

- [x] Smooth lines (c:smooth) read/render/write.
- [x] Series markers (c:marker symbol/size) on line/scatter read/render/write.
- [x] Series theme colours (schemeClr) + gradient/pattern fills -> a resolved CSS colour (approx).
- [x] Number formats on value axis + data labels (use the source cell number format / c:numFmt via SSF).
- [x] Per-series data labels + label position (c:dLbls at ser, c:dLblPos); showCatName/showSerName/
      showPercent. (Chart-level and per-series; rendered via the datalabels plugin.)
- [x] ODS secondary axis (chart:attached-axis) read + write; ODS data labels (chart:data-label-*) read/write.
- [x] dateAx (date axis) read + write; rendered on a proportional linear time axis (points spaced by
      date, date-formatted ticks) for line/area, no date-adapter dependency. Multi-level categories
      (c:multiLvlStrRef) read + preserved through edit/write (innermost level renders).
- [x] Legend entry deletion (c:legendEntry delete) read/render/write; legend overlay preserved.
- [x] Pie firstSliceAng (rotation) + pie explosion (c:explosion) read/render/write.

### Tier 3 - needs a plugin or custom draw

- [x] Trendlines (c:trendline: linear/exp/log/poly/power/movingAvg) via a dependency-free plugin
      (regressions computed in-house); equation + R-squared display.
- [x] Error bars (c:errBars) via a dependency-free plugin (fixedVal/percentage/stdDev/stdErr/cust).
- [x] stockChart (high-low-close / open-high-low-close): rendered as candlesticks / HLC bars by a
      plugin over invisible scale-carrier series.
- [x] ofPieChart (pie-of-pie / bar-of-pie): main pie + "Other" aggregate, with the breakout drawn
      as a secondary pie or bar in a reserved strip.
- [x] surfaceChart / surface3DChart: rendered as a heatmap (series x categories, value -> colour).
- [x] 3D chart scene: rendered 2D; on write an edited 3D chart re-emits the 3D chart-type element
      (bar3DChart etc.) + a minimal view3D so it stays 3D in Excel.

### Tier 4 - styling fidelity

- [x] Fonts / colours / bold-italic on the title, legend and axes (labels + titles) via txPr /
      rich-run rPr; series line width + dash (a:ln / a:prstDash); plot-area and chart-area
      background fills (spPr). Read + render (Chart.js font/color, borderWidth/borderDash, a
      background plugin) + write. Also fixed axis titles being read but never written.
- [x] Base palette from the colours part (colors1.xml via chartColorStyle) or the theme accents,
      used for series/slices without an explicit colour.

## Preservation guarantees (keep true throughout)

- An untouched chart is written back byte-for-byte (all of the above survives regardless).
- Each modelled feature is one the edited/created path no longer loses. As tiers land, the
  "editing a rich chart degrades it" caveat shrinks.

## Progress

- (DONE) Tier 1 - all six items shipped, xlsx+ods round-trip tested, LibreOffice-validated (chart:percentage read), 100% stacked browser-verified. Insert dialog gained Stacked + 100% stacked options.
- (DONE) Tier 2 - smooth lines, series markers, schemeClr/theme colour resolution, axis+label number formats (SSF), pie rotation, rich data labels (content + position, chart + per-series), pie explosion, legend-entry deletion + overlay, ODS secondary axis + data labels, proportional date axis (line/area), multi-level categories (read + preserved on write). xlsx round-trip tested; data-labels + explosion + smooth/marker + date-axis browser-verified; LibreOffice-validated (pie + percentage labels, and the date-axis line chart, survive the xlsx->ods convert). Both earlier caveats closed: date axes now render on a proportional time axis with date-formatted ticks, and editing a multi-level-category chart preserves all levels.
- (DONE) Tier 3 - trendlines, error bars, stock/candlestick, pie-of-pie/bar-of-pie, surface heatmap, and 3D-on-write all shipped. Rendering uses five dependency-free Chart.js plugins in src/core/ui/chart-plugins.ts (regression math, error whiskers, candlesticks, heatmap, of-pie secondary plot); read + write cover c:trendline, c:errBars, c:stockChart, c:ofPieChart, c:surfaceChart/surface3DChart, and the 3D chart-type elements + c:view3D. Each browser-verified and LibreOffice-validated. Also fixed a scatter/bubble buildChart bug (firstColLabels made x == y).
- (DONE) Tier 4 - element text styling (title/legend/axis fonts, colours, bold/italic), series line width + dash, plot/chart-area background fills, and a base palette from colors1.xml or the theme accents. Read + render + write; each round-trip-tested and browser-verified; LibreOffice keeps the title styling. Axis titles now round-trip (were read-only before).

All four tiers of CHARTS_SPEC_GAPS are complete. 3-D column and pie/doughnut now render with a pseudo-3D (isometric) look matching Excel; bar/line/area 3-D still render flat (and re-emit their 3-D element + view3D on write). Remaining niceties out of scope of this doc: a live time-scale for date axes (currently proportional-linear) and multi-level category axis display (innermost level shown).
