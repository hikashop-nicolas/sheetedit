import { t } from "../i18n";
import { parseA1Ref, type Workbook } from "../model";
import { buildChart, defaultAnchor, type Rect } from "../chart-build";
import { chartConfig, loadChartJs, type ChartCtor } from "./chart-overlay";
import { resolveLabels, seriesName } from "../chart-data";
import { CHART_PALETTE, type ChartKind, type ChartModel } from "../chart-model";

// The chart create/edit UI: one dialog (type, range, options, appearance, live preview) used both
// to build a new ChartModel and to edit an existing one, plus a small floating toolbar on a
// selected chart (Edit, Delete).

export interface ChartUiDeps {
  wrap: HTMLElement;
  gridScroll: HTMLElement;
  getWorkbook: () => Workbook;
  activeSheetName: () => string;
  onCreate: (model: ChartModel) => void;
  onDelete: (model: ChartModel) => void;
  onChange: (model: ChartModel) => void;
  boxRect: (id: string) => DOMRect | null;
}

const KINDS: { kind: ChartKind; key: string }[] = [
  { kind: "column", key: "chartColumn" }, { kind: "bar", key: "chartBar" }, { kind: "line", key: "chartLine" },
  { kind: "area", key: "chartArea" }, { kind: "pie", key: "chartPie" }, { kind: "doughnut", key: "chartDoughnut" },
  { kind: "scatter", key: "chartScatter" }, { kind: "radar", key: "chartRadar" },
];
const isScatter = (k: ChartKind): boolean => k === "scatter" || k === "bubble";
const isPie = (k: ChartKind): boolean => k === "pie" || k === "doughnut";
const isCartesian = (k: ChartKind): boolean => k === "column" || k === "bar" || k === "line" || k === "area";
// Which option groups apply to which chart kinds (a group is hidden when it does not apply).
const canStack = (k: ChartKind): boolean => k === "column" || k === "bar" || k === "area";
const canCombo = (k: ChartKind): boolean => isCartesian(k);
const hasLine = (k: ChartKind): boolean => k === "line" || k === "area" || k === "scatter" || k === "radar";
const hasAxes = (k: ChartKind): boolean => isCartesian(k) || k === "scatter";
const DASHES: { val: string; key: string }[] = [{ val: "solid", key: "chartDashSolid" }, { val: "dash", key: "chartDashDash" }, { val: "dot", key: "chartDashDot" }];

const STYLE_ID = "sheetedit-chart-style";
function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-chart-resize { position:absolute; right:-3px; bottom:-3px; width:12px; height:12px; cursor:nwse-resize; z-index:2;
      background:var(--sheetedit-accent, #6e7bff); border:2px solid #fff; border-radius:3px; opacity:0; }
    .sheetedit-chartbox.sel .sheetedit-chart-resize { opacity:1; }
    .sheetedit-chartbox { cursor:move; }
    .sheetedit-chart-modal { position:fixed; inset:0; z-index:60; display:flex; align-items:center; justify-content:center; background:rgba(0,0,0,.45); padding:16px; }
    .sheetedit-chart-card { display:flex; flex-direction:column; width:min(860px,100%); max-height:90vh; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e6e6e6);
      border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 14px 44px rgba(0,0,0,.5); font:13px system-ui,sans-serif; overflow:hidden; }
    .sheetedit-chart-head { display:flex; align-items:center; justify-content:space-between; padding:14px 16px; border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
    .sheetedit-chart-head h3 { margin:0; font-size:15px; }
    .sheetedit-chart-x { font:inherit; font-size:18px; line-height:1; width:28px; height:28px; border:none; border-radius:6px; background:transparent; color:var(--sheetedit-muted, #aab2bf); cursor:pointer; }
    .sheetedit-chart-x:hover { background:var(--sheetedit-btn, #3a3f47); color:inherit; }
    .sheetedit-chart-body { display:flex; gap:18px; padding:16px; overflow:hidden; min-height:0; }
    .sheetedit-chart-opts { flex:1 1 auto; min-width:0; overflow-y:auto; padding-right:4px; }
    .sheetedit-chart-side { flex:0 0 340px; display:flex; flex-direction:column; }
    .sheetedit-chart-sec { margin-bottom:14px; }
    .sheetedit-chart-sec[hidden] { display:none; }
    .sheetedit-chart-card h4 { margin:0 0 8px; font-size:11px; text-transform:uppercase; letter-spacing:.05em; color:var(--sheetedit-muted, #aab2bf); border-bottom:1px solid var(--sheetedit-border, #1c1f24); padding-bottom:4px; }
    .sheetedit-chart-types { display:flex; flex-wrap:wrap; gap:6px; }
    .sheetedit-chart-type { font:inherit; font-size:12px; padding:6px 11px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px;
      background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); cursor:pointer; }
    .sheetedit-chart-type.sel { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:#fff; }
    .sheetedit-chart-field { display:flex; flex-direction:column; gap:4px; margin-bottom:10px; }
    .sheetedit-chart-field[hidden] { display:none; }
    .sheetedit-chart-field > span { color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
    .sheetedit-chart-field input[type=text], .sheetedit-chart-field select, .sheetedit-chart-row select, .sheetedit-chart-row input[type=number] {
      font:inherit; background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); padding:6px 8px; }
    .sheetedit-chart-checks { display:flex; flex-wrap:wrap; gap:12px 16px; font-size:13px; }
    .sheetedit-chart-checks label { display:flex; align-items:center; gap:6px; }
    .sheetedit-chart-row { display:flex; flex-wrap:wrap; align-items:center; gap:12px 16px; margin-top:8px; }
    .sheetedit-chart-row label { display:flex; align-items:center; gap:6px; font-size:12px; color:var(--sheetedit-muted, #aab2bf); }
    .sheetedit-chart-row input[type=number] { width:64px; }
    .sheetedit-chart-two { display:flex; gap:12px; }
    .sheetedit-chart-two > * { flex:1; }
    .sheetedit-chart-swatches { display:flex; flex-wrap:wrap; gap:10px 14px; }
    .sheetedit-chart-swatch { display:flex; align-items:center; gap:6px; font-size:12px; }
    .sheetedit-chart-swatch input[type=color] { width:26px; height:22px; padding:0; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:4px; background:none; cursor:pointer; }
    .sheetedit-chart-fill { display:flex; align-items:center; gap:6px; }
    .sheetedit-chart-preview { flex:1 1 auto; min-height:240px; background:#fff; border-radius:6px; padding:8px; }
    .sheetedit-chart-foot { display:flex; justify-content:flex-end; gap:8px; padding:12px 16px; border-top:1px solid var(--sheetedit-border, #1c1f24); }
    .sheetedit-chart-btn { font:inherit; font-size:13px; padding:7px 16px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px;
      background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); cursor:pointer; }
    .sheetedit-chart-btn.primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:#fff; }
    @media (max-width:720px) {
      .sheetedit-chart-body { flex-direction:column; overflow-y:auto; }
      .sheetedit-chart-side { flex-basis:auto; order:-1; }
      .sheetedit-chart-preview { min-height:170px; height:170px; }
      .sheetedit-chart-opts { overflow:visible; }
    }
    .sheetedit-chart-editbar { position:fixed; z-index:30; display:flex; align-items:center; gap:6px; padding:5px 7px;
      background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
    .sheetedit-chart-editbar[hidden] { display:none; }
    .sheetedit-chart-editbar button { font:inherit; font-size:12px; background:var(--sheetedit-btn, #3a3f47);
      color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:5px; padding:4px 10px; cursor:pointer; }
  `;
  document.head.appendChild(s);
}

function parseRange(text: string, fallbackSheet: string): { sheet: string; rect: Rect } | null {
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(text.trim());
  const sheet = m ? (m[1] ?? m[2]) : fallbackSheet;
  const body = (m ? m[3] : text).replace(/\$/g, "");
  const [a, b] = body.split(":");
  const p1 = parseA1Ref(a ?? "");
  const p2 = b ? parseA1Ref(b) : p1;
  if (!p1 || !p2) return null;
  return { sheet, rect: { r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) } };
}

/** The bounding range + header flags of an existing chart, to pre-fill the edit dialog. */
function deriveRange(model: ChartModel, fallbackSheet: string): { text: string; firstRow: boolean; firstCol: boolean } {
  const refs: string[] = [];
  const push = (r?: { ref?: string }): void => { if (r?.ref) refs.push(r.ref); };
  push(model.categories);
  for (const s of model.series) { push(s.values); push(s.xValues); if (typeof s.name === "object") push(s.name); }
  let sheet = "", r1 = Infinity, c1 = Infinity, r2 = -Infinity, c2 = -Infinity;
  for (const ref of refs) {
    const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(ref);
    if (m) sheet = m[1] ?? m[2];
    const body = (m ? m[3] : ref).replace(/\$/g, "");
    const [a, b] = body.split(":");
    for (const part of [a, b].filter(Boolean)) { const p = parseA1Ref(part!); if (p) { r1 = Math.min(r1, p.row); c1 = Math.min(c1, p.col); r2 = Math.max(r2, p.row); c2 = Math.max(c2, p.col); } }
  }
  const sh = sheet || fallbackSheet;
  const text = isFinite(r1) ? `${sh}!${cellA1(c1, r1)}:${cellA1(c2, r2)}` : `${sh}!A1`;
  return { text, firstRow: model.series.some((s) => typeof s.name === "object" && !!s.name.ref), firstCol: !!model.categories?.ref };
}

export function setupChartUi(deps: ChartUiDeps): { openInsert(rect: Rect): void; showEdit(model: ChartModel): void; hideEdit(): void; teardown(): void } {
  injectStyles();

  // ---- Create/edit dialog (editModel omitted -> create) ----
  function openDialog(sel: Rect | null, editModel?: ChartModel): void {
    const wb = deps.getWorkbook();
    const sheetName = deps.activeSheetName();
    const derived = editModel ? deriveRange(editModel, sheetName) : null;
    let kind: ChartKind = editModel && KINDS.some((x) => x.kind === editModel.kind) ? editModel.kind : "column";
    const origRange = derived?.text ?? (sel ? `${sheetName}!${cellA1(sel.c1, sel.r1)}:${cellA1(sel.c2, sel.r2)}` : `${sheetName}!A1`);
    const origFirstRow = derived ? derived.firstRow : true;
    const origFirstCol = derived ? derived.firstCol : true;
    const line0 = editModel?.series.find((s) => s.lineWidth != null || s.dash);
    const state = {
      firstRowHeader: origFirstRow, firstColLabels: origFirstCol,
      showLegend: editModel ? (editModel.legend?.show ?? true) : true,
      legendPos: editModel?.legend?.pos ?? "bottom",
      dataLabels: !!editModel?.dataLabels,
      stacked: !!editModel?.stacked, percent: !!editModel?.percent,
      comboLine: !!editModel?.series.some((s) => s.type === "line"),
      comboSecondary: !!editModel?.series.some((s) => s.secondaryAxis),
      title: editModel?.title ?? "",
      titleBold: !!editModel?.titleStyle?.bold,
      titleColor: editModel?.titleStyle?.color ?? "",
      xTitle: editModel?.axes?.x?.title ?? "", yTitle: editModel?.axes?.y?.title ?? "",
      seriesColors: [] as (string | undefined)[],
      sliceColors: [] as (string | undefined)[],
      lineWidth: line0?.lineWidth ?? 0,
      dash: line0?.dash ?? "solid",
      plotFill: editModel?.plotFill ?? "", areaFill: editModel?.areaFill ?? "",
    };
    if (editModel) {
      editModel.series.forEach((s, i) => { if (s.color) state.seriesColors[i] = s.color; });
      (editModel.series[0]?.pointColors ?? []).forEach((c, j) => { if (c) state.sliceColors[j] = c; });
    }

    const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string): HTMLElementTagNameMap[K] => { const e = document.createElement(tag); if (cls) e.className = cls; return e; };
    const section = (titleKey: string, ...children: (HTMLElement | null)[]): HTMLDivElement => {
      const s = el("div", "sheetedit-chart-sec");
      const head = el("h4"); head.textContent = t(titleKey);
      s.append(head, ...children.filter((c): c is HTMLElement => !!c));
      return s;
    };
    const field = (labelKey: string, input: HTMLElement): HTMLLabelElement => {
      const l = el("label", "sheetedit-chart-field");
      const sp = el("span"); sp.textContent = t(labelKey);
      l.append(sp, input);
      return l;
    };
    const textInput = (val: string, onInput: (v: string) => void): HTMLInputElement => {
      const i = el("input"); i.type = "text"; i.value = val;
      i.addEventListener("input", () => onInput(i.value));
      return i;
    };

    const modal = el("div", "sheetedit-chart-modal");
    const card = el("div", "sheetedit-chart-card");

    // Header
    const head = el("div", "sheetedit-chart-head");
    const h = el("h3"); h.textContent = editModel ? t("chartEdit") : t("chartInsert");
    const closeX = el("button", "sheetedit-chart-x"); closeX.type = "button"; closeX.textContent = "×";
    head.append(h, closeX);

    // Type strip
    const types = el("div", "sheetedit-chart-types");
    const typeBtns = KINDS.map(({ kind: k, key }) => {
      const b = el("button", "sheetedit-chart-type" + (k === kind ? " sel" : ""));
      b.type = "button";
      b.textContent = t(key);
      b.addEventListener("click", () => { kind = k; typeBtns.forEach((x) => x.classList.toggle("sel", x === b)); updateVisibility(); renderSwatches(); redraw(); });
      types.appendChild(b);
      return b;
    });

    // Data: range + header flags
    const rangeInput = textInput(origRange, () => { updateVisibility(); renderSwatches(); redraw(); });
    const checks = el("div", "sheetedit-chart-checks");
    type BoolKey = "firstRowHeader" | "firstColLabels";
    let firstColField: HTMLLabelElement;
    const mkCheck = (label: string, get: () => boolean, set: (v: boolean) => void, onToggle?: () => void): HTMLLabelElement => {
      const l = el("label");
      const cb = el("input"); cb.type = "checkbox"; cb.checked = get();
      cb.addEventListener("change", () => { set(cb.checked); onToggle?.(); renderSwatches(); redraw(); });
      const sp = el("span"); sp.textContent = label;
      l.append(cb, sp);
      return l;
    };
    const boolCheck = (label: string, key: BoolKey, onToggle?: () => void): HTMLLabelElement => mkCheck(label, () => state[key], (v) => { state[key] = v; }, onToggle);
    const firstRowChk = boolCheck(t("chartFirstRow"), "firstRowHeader");
    firstColField = boolCheck(t("chartFirstCol"), "firstColLabels");
    checks.append(firstRowChk, firstColField);

    // Title & legend
    const titleInput = textInput(state.title, (v) => { state.title = v; redraw(); });
    const boldChk = mkCheck(t("chartTitleBold"), () => state.titleBold, (v) => { state.titleBold = v; });
    const titleColour = fillControl(t("chartTitleColour"), state.titleColor, (v) => { state.titleColor = v; redraw(); });
    const titleRow = el("div", "sheetedit-chart-row"); titleRow.append(boldChk, titleColour);
    const legendChk = mkCheck(t("chartLegend"), () => state.showLegend, (v) => { state.showLegend = v; }, () => updateVisibility());
    const legendSel = el("select");
    for (const [val, key] of [["top", "chartPosTop"], ["bottom", "chartPosBottom"], ["left", "chartPosLeft"], ["right", "chartPosRight"]]) { const o = el("option"); o.value = val; o.textContent = t(key); legendSel.appendChild(o); }
    legendSel.value = state.legendPos;
    legendSel.addEventListener("change", () => { state.legendPos = legendSel.value as typeof state.legendPos; redraw(); });
    const legendPosField = field("chartLegendPos", legendSel);
    const dataLabelsChk = mkCheck(t("chartDataLabels"), () => state.dataLabels, (v) => { state.dataLabels = v; });
    const legendRow = el("div", "sheetedit-chart-row"); legendRow.append(legendChk, dataLabelsChk);

    // Layout (stacking) - column/bar/area
    const stackChk = mkCheck(t("chartStacked"), () => state.stacked, (v) => { state.stacked = v; });
    const percentChk = mkCheck(t("chartPercent"), () => state.percent, (v) => { state.percent = v; });
    const layoutRow = el("div", "sheetedit-chart-checks"); layoutRow.append(stackChk, percentChk);
    const layoutSec = section("chartLayoutSection", layoutRow);

    // Combo - cartesian, >=2 series
    const comboLineChk = mkCheck(t("chartComboLine"), () => state.comboLine, (v) => { state.comboLine = v; });
    const comboSecChk = mkCheck(t("chartComboSecondary"), () => state.comboSecondary, (v) => { state.comboSecondary = v; });
    const comboRow = el("div", "sheetedit-chart-checks"); comboRow.append(comboLineChk, comboSecChk);
    const comboSec = section("chartComboSection", comboRow);

    // Line width/dash - line/area/scatter/radar
    const lwInput = el("input"); lwInput.type = "number"; lwInput.min = "0"; lwInput.step = "0.5"; lwInput.value = String(state.lineWidth || "");
    lwInput.addEventListener("input", () => { state.lineWidth = Number(lwInput.value) || 0; redraw(); });
    const lwLbl = el("label"); lwLbl.append(document.createTextNode(t("chartLineWidth")), lwInput);
    const dashSel = el("select");
    for (const { val, key } of DASHES) { const o = el("option"); o.value = val; o.textContent = t(key); dashSel.appendChild(o); }
    dashSel.value = state.dash;
    dashSel.addEventListener("change", () => { state.dash = dashSel.value; redraw(); });
    const dashLbl = el("label"); dashLbl.append(document.createTextNode(t("chartDash")), dashSel);
    const lineRow = el("div", "sheetedit-chart-row"); lineRow.append(lwLbl, dashLbl);
    const lineSec = section("chartLineSection", lineRow);

    // Axes - cartesian + scatter
    const xTitleInput = textInput(state.xTitle, (v) => { state.xTitle = v; redraw(); });
    const yTitleInput = textInput(state.yTitle, (v) => { state.yTitle = v; redraw(); });
    const axesTwo = el("div", "sheetedit-chart-two"); axesTwo.append(field("chartXTitle", xTitleInput), field("chartYTitle", yTitleInput));
    const axesSec = section("chartAxesSection", axesTwo);

    // Series / slice colours
    const swatches = el("div", "sheetedit-chart-swatches");
    const coloursHead = el("h4");
    const coloursSec = el("div", "sheetedit-chart-sec"); coloursSec.append(coloursHead, swatches);

    // Background fills
    const bgRow = el("div", "sheetedit-chart-row");
    bgRow.append(fillControl(t("chartPlotFill"), state.plotFill, (v) => { state.plotFill = v; redraw(); }), fillControl(t("chartAreaFill"), state.areaFill, (v) => { state.areaFill = v; redraw(); }));
    const bgSec = section("chartBgSection", bgRow);

    // Assemble options column
    const opts = el("div", "sheetedit-chart-opts");
    opts.append(
      section("chartTypeSection", types),
      section("chartDataSection", field("chartRange", rangeInput), checks),
      section("chartTitleSection", field("chartTitle", titleInput), titleRow, legendRow, legendPosField),
      layoutSec, comboSec, lineSec, axesSec, coloursSec, bgSec,
    );

    // Preview column
    const side = el("div", "sheetedit-chart-side");
    const preview = el("div", "sheetedit-chart-preview");
    const canvas = el("canvas");
    preview.appendChild(canvas);
    side.append(preview);

    const body = el("div", "sheetedit-chart-body");
    body.append(opts, side);

    // Footer
    const foot = el("div", "sheetedit-chart-foot");
    const cancel = el("button", "sheetedit-chart-btn"); cancel.textContent = t("chartCancel");
    const submit = el("button", "sheetedit-chart-btn primary"); submit.textContent = editModel ? t("chartApply") : t("chartInsertBtn");
    foot.append(cancel, submit);

    card.append(head, body, foot);
    modal.appendChild(card);
    deps.wrap.appendChild(modal);

    // Show only the sections/fields that apply to the current kind + series count.
    function updateVisibility(): void {
      const count = baseModelRaw()?.series.length ?? 0;
      layoutSec.hidden = !canStack(kind);
      comboSec.hidden = !(canCombo(kind) && count >= 2);
      lineSec.hidden = !hasLine(kind);
      axesSec.hidden = !hasAxes(kind);
      firstColField.hidden = isScatter(kind);
      legendPosField.hidden = !state.showLegend;
      coloursHead.textContent = isPie(kind) ? t("chartSliceColours") : t("chartSeriesColours");
    }
    closeX.addEventListener("click", () => close());

    // Build a model from the current state; reuse the edited chart's series unless the data
    // structure changed (range / header flags / scatter-vs-category), so read-in richness survives.
    function baseModel(forId: string): ChartModel | null {
      const parsed = parseRange(rangeInput.value, sheetName);
      if (!parsed) return null;
      const structChanged = !editModel || rangeInput.value.trim() !== origRange.trim() || state.firstRowHeader !== origFirstRow || state.firstColLabels !== origFirstCol || isScatter(kind) !== isScatter(editModel.kind);
      let m: ChartModel;
      if (editModel && !structChanged) {
        m = JSON.parse(JSON.stringify(editModel)) as ChartModel;
        m.kind = kind;
      } else {
        m = buildChart(parsed.sheet, kind, parsed.rect, { firstRowHeader: state.firstRowHeader, firstColLabels: state.firstColLabels }, forId, editModel?.anchor ?? defaultAnchor(parsed.rect));
        if (editModel) { m.id = editModel.id; m.anchor = editModel.anchor; m.original = editModel.original; m.palette = editModel.palette; }
      }
      applyState(m);
      return m;
    }
    function applyState(m: ChartModel): void {
      m.title = state.title || undefined;
      m.titleStyle = state.titleBold || state.titleColor ? { bold: state.titleBold || undefined, color: state.titleColor || undefined } : undefined;
      m.legend = { show: state.showLegend, pos: state.legendPos };
      m.dataLabels = state.dataLabels || undefined;
      m.stacked = canStack(m.kind) ? (state.stacked || state.percent || undefined) : undefined;
      m.percent = canStack(m.kind) ? (state.percent || undefined) : undefined;
      if (canCombo(m.kind) && m.series.length >= 2) {
        const last = m.series[m.series.length - 1];
        last.type = state.comboLine ? "line" : undefined;
        last.secondaryAxis = state.comboSecondary || undefined;
      }
      if (isPie(m.kind)) {
        if (m.series[0] && state.sliceColors.some((c) => c)) m.series[0].pointColors = state.sliceColors.map((c) => c || undefined);
      } else {
        state.seriesColors.forEach((c, i) => { if (m.series[i]) m.series[i].color = c || undefined; });
      }
      const width = hasLine(m.kind) && state.lineWidth > 0 ? state.lineWidth : undefined;
      const dash = hasLine(m.kind) && state.dash && state.dash !== "solid" ? state.dash : undefined;
      m.series.forEach((s) => { s.lineWidth = width; s.dash = dash; });
      const useAxes = hasAxes(m.kind);
      const xa = useAxes && state.xTitle ? { ...(m.axes?.x ?? {}), title: state.xTitle } : m.axes?.x;
      const ya = useAxes && state.yTitle ? { ...(m.axes?.y ?? {}), title: state.yTitle } : m.axes?.y;
      m.axes = xa || ya ? { x: xa, y: ya } : undefined;
      m.plotFill = state.plotFill || undefined;
      m.areaFill = state.areaFill || undefined;
    }

    // A colour input per series, or per slice for pie/doughnut (labelled by category). Blank input
    // value shows the current effective colour; editing it makes that colour explicit.
    function renderSwatches(): void {
      const probe = baseModelRaw();
      swatches.textContent = "";
      const swatch = (colour: string, label: string, onPick: (v: string) => void): void => {
        const wrap = document.createElement("label");
        wrap.className = "sheetedit-chart-swatch";
        const inp = document.createElement("input");
        inp.type = "color";
        inp.value = colour;
        inp.addEventListener("input", () => { onPick(inp.value); redraw(); });
        const name = document.createElement("span");
        name.textContent = label;
        wrap.append(inp, name);
        swatches.appendChild(wrap);
      };
      const pal = (i: number): string => probe?.palette?.[i] ?? CHART_PALETTE[i % CHART_PALETTE.length];
      if (isPie(kind)) {
        const cats = resolveLabels(wb, probe?.categories);
        const pts = probe?.series[0]?.pointColors ?? [];
        cats.forEach((label, j) => swatch(state.sliceColors[j] ?? pts[j] ?? pal(j), label || `#${j + 1}`, (v) => { state.sliceColors[j] = v; }));
      } else {
        (probe?.series ?? []).forEach((s, i) => swatch(state.seriesColors[i] ?? s.color ?? pal(i), seriesName(wb, s.name) ?? `#${i + 1}`, (v) => { state.seriesColors[i] = v; }));
      }
    }
    // A model without the per-series colour application (used to probe series names/count).
    function baseModelRaw(): ChartModel | null {
      const parsed = parseRange(rangeInput.value, sheetName);
      if (!parsed) return null;
      if (editModel && rangeInput.value.trim() === origRange.trim() && state.firstRowHeader === origFirstRow && state.firstColLabels === origFirstCol && isScatter(kind) === isScatter(editModel.kind)) {
        const m = JSON.parse(JSON.stringify(editModel)) as ChartModel; m.kind = kind; return m;
      }
      return buildChart(parsed.sheet, kind, parsed.rect, { firstRowHeader: state.firstRowHeader, firstColLabels: state.firstColLabels }, "probe", defaultAnchor(parsed.rect));
    }

    let chart: InstanceType<ChartCtor> | null = null;
    let previewSeq = 0;
    function redraw(): void {
      const model = baseModel("preview");
      if (!model) return;
      const seq = ++previewSeq;
      void loadChartJs().then((Ctor) => {
        if (seq !== previewSeq) return;
        chart?.destroy();
        chart = new Ctor(canvas.getContext("2d")!, chartConfig(model, wb));
      });
    }
    const close = (): void => { chart?.destroy(); modal.remove(); };
    cancel.addEventListener("click", close);
    modal.addEventListener("mousedown", (e) => { if (e.target === modal) close(); });
    submit.addEventListener("click", () => {
      const model = baseModel(editModel?.id ?? uniqueId());
      if (!model) return;
      model.dirty = true;
      close();
      if (editModel) { Object.assign(editModel, model); deps.onChange(editModel); }
      else deps.onCreate(model);
    });
    updateVisibility();
    renderSwatches();
    redraw();
  }

  // ---- Edit toolbar on a selected chart: Edit + Delete ----
  const editbar = document.createElement("div");
  editbar.className = "sheetedit-chart-editbar";
  editbar.hidden = true;
  const editBtn = document.createElement("button");
  editBtn.textContent = t("chartEdit");
  const delBtn = document.createElement("button");
  delBtn.textContent = t("chartDelete");
  editbar.append(editBtn, delBtn);
  deps.wrap.appendChild(editbar);
  let editing: ChartModel | null = null;

  const positionEditbar = (): void => {
    if (!editing) return;
    const r = deps.boxRect(editing.id);
    const wr = deps.wrap.getBoundingClientRect();
    const gr = deps.gridScroll.getBoundingClientRect();
    if (!r) { editbar.hidden = true; return; }
    editbar.hidden = false;
    editbar.style.left = `${Math.max(4, r.left - wr.left)}px`;
    const above = r.top - editbar.offsetHeight - 4;
    editbar.style.top = `${(above >= gr.top ? above : r.top + 4) - wr.top}px`;
  };
  editBtn.addEventListener("click", () => { if (editing) openDialog(null, editing); });
  delBtn.addEventListener("click", () => { if (editing) { const m = editing; hideEdit(); deps.onDelete(m); } });
  deps.gridScroll.addEventListener("scroll", positionEditbar, { passive: true });

  function showEdit(model: ChartModel): void { editing = model; positionEditbar(); }
  function hideEdit(): void { editing = null; editbar.hidden = true; }

  return { openInsert: (rect) => openDialog(rect), showEdit, hideEdit, teardown() { editbar.remove(); deps.gridScroll.removeEventListener("scroll", positionEditbar); } };
}

/** A "colour + clear" control: a colour input plus a small clear (X) that resets to no colour. */
function fillControl(label: string, initial: string, onChange: (v: string) => void): HTMLElement {
  const wrap = document.createElement("label");
  wrap.className = "sheetedit-chart-fill";
  const sp = document.createElement("span");
  sp.style.cssText = "font-size:12px;color:var(--sheetedit-muted,#aab2bf)";
  sp.textContent = label;
  const inp = document.createElement("input");
  inp.type = "color";
  inp.style.cssText = "width:26px;height:22px;padding:0;border:1px solid var(--sheetedit-btn,#3a4047);border-radius:4px;background:none;cursor:pointer";
  if (initial) inp.value = initial;
  inp.addEventListener("input", () => onChange(inp.value));
  const clear = document.createElement("button");
  clear.type = "button";
  clear.textContent = "×";
  clear.title = t("chartNone");
  clear.style.cssText = "font:inherit;font-size:13px;line-height:1;padding:2px 6px;border:1px solid var(--sheetedit-btn-border,#4a4f57);border-radius:4px;background:var(--sheetedit-btn,#3a3f47);color:inherit;cursor:pointer";
  clear.addEventListener("click", () => onChange(""));
  wrap.append(sp, inp, clear);
  return wrap;
}

let idSeq = 0;
const uniqueId = (): string => `chart-new-${++idSeq}`;
const cellA1 = (col: number, row: number): string => `${colLetters(col)}${row}`;
function colLetters(c: number): string { let s = ""; while (c > 0) { s = String.fromCharCode(65 + ((c - 1) % 26)) + s; c = Math.floor((c - 1) / 26); } return s; }
