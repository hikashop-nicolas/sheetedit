# Charts: full spec coverage (DrawingML + ODF)

Goal: close the gap between our chart support and the ECMA-376 DrawingML chart schema (Part 1,
§21.2) and the ODF chart schema, for read + render + write round-trip. Renderer is Chart.js
(some features need its plugins or a custom draw); the write policy stays "preserve an untouched
chart verbatim, re-emit an edited/created chart from our model", so every feature we model is one
less thing lost when a chart is edited.

## Tiers

### Tier 1 - quick wins (this pass)

- [ ] Axis min/max: read c:valAx/c:catAx c:scaling c:max/c:min -> model.axes; apply in Chart.js scales.
- [ ] dispBlanksAs (gap|zero|span): read chart-level; plot nulls as gaps/zeros/spanned; write it back.
- [ ] Doughnut holeSize, bar gapWidth + overlap: read/apply (cutout, categoryPercentage/barPercentage)/write.
- [ ] percentStacked: true 100% stacking (normalise per-category to %) on render; grouping on write.
- [ ] ODS bar-vs-column + stacked/percentage: read the plot-area chart-properties style (chart:vertical,
      chart:stacked, chart:percentage) so ODS bar charts and stacked charts render right.
- [ ] Per-point colours (c:dPt): read per-point spPr srgbClr onto series.pointColors; use on render;
      write c:dPt. (Makes read pie/bar charts match the file's colours.)

### Tier 2 - moderate

- [ ] Smooth lines (c:smooth) read/render/write.
- [ ] Series markers (c:marker symbol/size) on line/scatter read/render/write.
- [ ] Series theme colours (schemeClr) + gradient/pattern fills -> a resolved CSS colour (approx).
- [ ] Number formats on value axis + data labels (use the source cell number format / c:numFmt via SSF).
- [ ] Per-series and per-point data labels + label position (c:dLbls at ser/dPt, c:dLblPos); showCatName/
      showSerName/showPercent.
- [ ] ODS secondary axis (chart:attached-axis) read; ODS data labels (chart:data-label-*) read/write.
- [ ] dateAx (date axis) read + a time scale; multi-level categories (c:multiLvlStrRef).
- [ ] Legend entry deletion (c:legendEntry delete); legend overlay.
- [ ] Pie firstSliceAng (rotation), pie explosion (c:explosion).

### Tier 3 - needs a plugin or custom draw

- [ ] Trendlines (c:trendline: linear/exp/log/poly/power/movingAvg) via a Chart.js plugin or custom.
- [ ] Error bars (c:errBars) via chartjs-plugin-error-bars or custom.
- [ ] stockChart (high-low-close / open-high-low-close) via a candlestick plugin or custom draw.
- [ ] ofPieChart (pie-of-pie / bar-of-pie) - approximate (a pie + a secondary breakdown).
- [ ] surfaceChart / surface3DChart - no 2D equivalent; render a heatmap-ish fallback or leave a
      placeholder that preserves the original on save.
- [ ] 3D chart scene (view3D/floor/walls): keep rendering 2D, but on WRITE of an edited 3D chart,
      emit the 3D chart-type element (bar3DChart etc.) + a minimal view3D so it stays 3D in Excel
      instead of silently flattening.

### Tier 4 - styling fidelity

- [ ] Fonts / fills / borders / effects on title, axes, legend, plot area and series (txPr/spPr
      throughout) mapped onto Chart.js element options where possible.
- [ ] The chart style + colours parts (style1.xml / colors1.xml) - read a base palette from them.

## Preservation guarantees (keep true throughout)

- An untouched chart is written back byte-for-byte (all of the above survives regardless).
- Each modelled feature is one the edited/created path no longer loses. As tiers land, the
  "editing a rich chart degrades it" caveat shrinks.

## Progress

- (in progress) Tier 1
