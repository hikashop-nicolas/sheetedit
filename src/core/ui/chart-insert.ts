import { t } from "../i18n";
import { parseA1Ref, type Workbook } from "../model";
import { buildChart, defaultAnchor, type Rect } from "../chart-build";
import { chartConfig, loadChartJs, type ChartCtor } from "./chart-overlay";
import { seriesName } from "../chart-data";
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
const isCartesian = (k: ChartKind): boolean => k === "column" || k === "bar" || k === "line" || k === "area";
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
    .sheetedit-chart-modal { position:fixed; inset:0; z-index:60; display:flex; align-items:center; justify-content:center; background:rgba(0,0,0,.45); }
    .sheetedit-chart-card { width:min(560px,94%); max-height:88%; overflow:auto; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e6e6e6);
      border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 14px 44px rgba(0,0,0,.5); padding:16px; font:13px system-ui,sans-serif; }
    .sheetedit-chart-card h3 { margin:0 0 12px; font-size:15px; }
    .sheetedit-chart-card h4 { margin:14px 0 8px; font-size:12px; text-transform:uppercase; letter-spacing:.04em; color:var(--sheetedit-muted, #aab2bf); border-bottom:1px solid var(--sheetedit-border, #1c1f24); padding-bottom:4px; }
    .sheetedit-chart-types { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:12px; }
    .sheetedit-chart-type { font:inherit; font-size:12px; padding:6px 11px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px;
      background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); cursor:pointer; }
    .sheetedit-chart-type.sel { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:#fff; }
    .sheetedit-chart-field { display:flex; flex-direction:column; gap:4px; margin-bottom:10px; }
    .sheetedit-chart-field > span { color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
    .sheetedit-chart-field input[type=text], .sheetedit-chart-field select, .sheetedit-chart-row select, .sheetedit-chart-row input[type=number] {
      font:inherit; background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); padding:6px 8px; }
    .sheetedit-chart-checks { display:flex; flex-wrap:wrap; gap:14px; margin-bottom:12px; font-size:13px; }
    .sheetedit-chart-checks label { display:flex; align-items:center; gap:6px; }
    .sheetedit-chart-row { display:flex; flex-wrap:wrap; align-items:center; gap:14px; margin-bottom:10px; }
    .sheetedit-chart-row label { display:flex; align-items:center; gap:6px; font-size:12px; color:var(--sheetedit-muted, #aab2bf); }
    .sheetedit-chart-row input[type=number] { width:70px; }
    .sheetedit-chart-swatches { display:flex; flex-wrap:wrap; gap:12px; margin-bottom:10px; }
    .sheetedit-chart-swatch { display:flex; align-items:center; gap:6px; font-size:12px; }
    .sheetedit-chart-swatch input[type=color] { width:26px; height:22px; padding:0; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:4px; background:none; cursor:pointer; }
    .sheetedit-chart-fill { display:flex; align-items:center; gap:6px; }
    .sheetedit-chart-preview { height:220px; background:#fff; border-radius:6px; padding:6px; margin:12px 0; }
    .sheetedit-chart-actions { display:flex; justify-content:flex-end; gap:8px; }
    .sheetedit-chart-btn { font:inherit; font-size:13px; padding:6px 14px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px;
      background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); cursor:pointer; }
    .sheetedit-chart-btn.primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:#fff; }
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
      lineWidth: line0?.lineWidth ?? 0,
      dash: line0?.dash ?? "solid",
      plotFill: editModel?.plotFill ?? "", areaFill: editModel?.areaFill ?? "",
    };
    if (editModel) editModel.series.forEach((s, i) => { if (s.color) state.seriesColors[i] = s.color; });

    const modal = document.createElement("div");
    modal.className = "sheetedit-chart-modal";
    const card = document.createElement("div");
    card.className = "sheetedit-chart-card";
    const h = document.createElement("h3");
    h.textContent = editModel ? t("chartEdit") : t("chartInsert");

    const types = document.createElement("div");
    types.className = "sheetedit-chart-types";
    const typeBtns = KINDS.map(({ kind: k, key }) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-chart-type" + (k === kind ? " sel" : "");
      b.textContent = t(key);
      b.addEventListener("click", () => { kind = k; typeBtns.forEach((x) => x.classList.toggle("sel", x === b)); renderSwatches(); redraw(); });
      types.appendChild(b);
      return b;
    });

    const field = (labelKey: string, input: HTMLElement): HTMLLabelElement => {
      const l = document.createElement("label");
      l.className = "sheetedit-chart-field";
      const sp = document.createElement("span");
      sp.textContent = t(labelKey);
      l.append(sp, input);
      return l;
    };
    const rangeInput = document.createElement("input");
    rangeInput.type = "text";
    rangeInput.value = origRange;
    rangeInput.addEventListener("input", () => { renderSwatches(); redraw(); });
    const titleInput = document.createElement("input");
    titleInput.type = "text";
    titleInput.value = state.title;
    titleInput.addEventListener("input", () => { state.title = titleInput.value; redraw(); });

    const checks = document.createElement("div");
    checks.className = "sheetedit-chart-checks";
    type BoolKey = "firstRowHeader" | "firstColLabels" | "showLegend" | "dataLabels" | "stacked" | "percent" | "comboLine" | "comboSecondary";
    const mkCheck = (label: string, key: BoolKey): HTMLLabelElement => {
      const l = document.createElement("label");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = state[key];
      cb.addEventListener("change", () => { state[key] = cb.checked; if (key === "firstRowHeader" || key === "firstColLabels") renderSwatches(); redraw(); });
      const sp = document.createElement("span");
      sp.textContent = label;
      l.append(cb, sp);
      return l;
    };
    checks.append(mkCheck(t("chartFirstRow"), "firstRowHeader"), mkCheck(t("chartFirstCol"), "firstColLabels"), mkCheck(t("chartLegend"), "showLegend"), mkCheck(t("chartDataLabels"), "dataLabels"), mkCheck(t("chartStacked"), "stacked"), mkCheck(t("chartPercent"), "percent"), mkCheck(t("chartComboLine"), "comboLine"), mkCheck(t("chartComboSecondary"), "comboSecondary"));

    // Legend position
    const legendSel = document.createElement("select");
    for (const [val, key] of [["top", "chartPosTop"], ["bottom", "chartPosBottom"], ["left", "chartPosLeft"], ["right", "chartPosRight"]]) { const o = document.createElement("option"); o.value = val; o.textContent = t(key); legendSel.appendChild(o); }
    legendSel.value = state.legendPos;
    legendSel.addEventListener("change", () => { state.legendPos = legendSel.value as typeof state.legendPos; redraw(); });

    // Appearance: title bold/colour, series colours, line width/dash, backgrounds.
    const appHead = document.createElement("h4");
    appHead.textContent = t("chartAppearance");
    const titleRow = document.createElement("div");
    titleRow.className = "sheetedit-chart-row";
    const boldLbl = document.createElement("label");
    const boldCb = document.createElement("input"); boldCb.type = "checkbox"; boldCb.checked = state.titleBold;
    boldCb.addEventListener("change", () => { state.titleBold = boldCb.checked; redraw(); });
    boldLbl.append(boldCb, document.createTextNode(t("chartTitleBold")));
    const titleColorWrap = fillControl(t("chartTitleColour"), state.titleColor, (v) => { state.titleColor = v; redraw(); });
    titleRow.append(boldLbl, titleColorWrap);

    const swHead = document.createElement("h4");
    swHead.textContent = t("chartSeriesColours");
    const swatches = document.createElement("div");
    swatches.className = "sheetedit-chart-swatches";

    const lineRow = document.createElement("div");
    lineRow.className = "sheetedit-chart-row";
    const lwLbl = document.createElement("label");
    const lwInput = document.createElement("input"); lwInput.type = "number"; lwInput.min = "0"; lwInput.step = "0.5"; lwInput.value = String(state.lineWidth || "");
    lwInput.addEventListener("input", () => { state.lineWidth = Number(lwInput.value) || 0; redraw(); });
    lwLbl.append(document.createTextNode(t("chartLineWidth")), lwInput);
    const dashLbl = document.createElement("label");
    const dashSel = document.createElement("select");
    for (const { val, key } of DASHES) { const o = document.createElement("option"); o.value = val; o.textContent = t(key); dashSel.appendChild(o); }
    dashSel.value = state.dash;
    dashSel.addEventListener("change", () => { state.dash = dashSel.value; redraw(); });
    dashLbl.append(document.createTextNode(t("chartDash")), dashSel);
    lineRow.append(lwLbl, dashLbl);

    const bgRow = document.createElement("div");
    bgRow.className = "sheetedit-chart-row";
    bgRow.append(fillControl(t("chartPlotFill"), state.plotFill, (v) => { state.plotFill = v; redraw(); }), fillControl(t("chartAreaFill"), state.areaFill, (v) => { state.areaFill = v; redraw(); }));

    const axHead = document.createElement("h4");
    axHead.textContent = t("chartAxesSection");
    const xTitleInput = document.createElement("input"); xTitleInput.type = "text"; xTitleInput.value = state.xTitle;
    xTitleInput.addEventListener("input", () => { state.xTitle = xTitleInput.value; redraw(); });
    const yTitleInput = document.createElement("input"); yTitleInput.type = "text"; yTitleInput.value = state.yTitle;
    yTitleInput.addEventListener("input", () => { state.yTitle = yTitleInput.value; redraw(); });

    const preview = document.createElement("div");
    preview.className = "sheetedit-chart-preview";
    const canvas = document.createElement("canvas");
    preview.appendChild(canvas);

    const actions = document.createElement("div");
    actions.className = "sheetedit-chart-actions";
    const cancel = document.createElement("button");
    cancel.className = "sheetedit-chart-btn";
    cancel.textContent = t("chartCancel");
    const submit = document.createElement("button");
    submit.className = "sheetedit-chart-btn primary";
    submit.textContent = editModel ? t("chartApply") : t("chartInsertBtn");
    actions.append(cancel, submit);

    card.append(h, types, field("chartRange", rangeInput), field("chartTitle", titleInput), checks,
      field("chartLegendPos", legendSel),
      appHead, titleRow, swHead, swatches, lineRow, bgRow,
      axHead, field("chartXTitle", xTitleInput), field("chartYTitle", yTitleInput),
      preview, actions);
    modal.appendChild(card);
    deps.wrap.appendChild(modal);

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
      m.stacked = state.stacked || state.percent || undefined;
      m.percent = state.percent || undefined;
      if (isCartesian(m.kind) && m.series.length >= 2) {
        const last = m.series[m.series.length - 1];
        last.type = state.comboLine ? "line" : undefined;
        last.secondaryAxis = state.comboSecondary || undefined;
      }
      state.seriesColors.forEach((c, i) => { if (m.series[i]) m.series[i].color = c || undefined; });
      const width = state.lineWidth > 0 ? state.lineWidth : undefined;
      const dash = state.dash && state.dash !== "solid" ? state.dash : undefined;
      m.series.forEach((s) => { s.lineWidth = width; s.dash = dash; });
      const xa = state.xTitle ? { ...(m.axes?.x ?? {}), title: state.xTitle } : m.axes?.x;
      const ya = state.yTitle ? { ...(m.axes?.y ?? {}), title: state.yTitle } : m.axes?.y;
      m.axes = xa || ya ? { x: xa, y: ya } : undefined;
      m.plotFill = state.plotFill || undefined;
      m.areaFill = state.areaFill || undefined;
    }

    // One colour input per series, showing the current effective colour; blank = follow the palette.
    function renderSwatches(): void {
      const probe = baseModelRaw();
      swatches.textContent = "";
      const series = probe?.series ?? [];
      series.forEach((s, i) => {
        const wrap = document.createElement("label");
        wrap.className = "sheetedit-chart-swatch";
        const inp = document.createElement("input");
        inp.type = "color";
        inp.value = state.seriesColors[i] ?? s.color ?? probe?.palette?.[i] ?? CHART_PALETTE[i % CHART_PALETTE.length];
        inp.addEventListener("input", () => { state.seriesColors[i] = inp.value; redraw(); });
        const name = document.createElement("span");
        name.textContent = seriesName(wb, s.name) ?? `#${i + 1}`;
        wrap.append(inp, name);
        swatches.appendChild(wrap);
      });
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
