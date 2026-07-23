import { t } from "../i18n";
import { parseA1Ref, type Workbook } from "../model";
import { buildChart, defaultAnchor, type Rect } from "../chart-build";
import { chartConfig, loadChartJs, type ChartCtor } from "./chart-overlay";
import type { ChartKind, ChartModel } from "../chart-model";

// The chart create/edit UI: an Insert dialog (type, range, options, live preview) that builds a
// ChartModel, and a small floating toolbar on a selected chart (change type, delete).

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
    .sheetedit-chart-types { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:12px; }
    .sheetedit-chart-type { font:inherit; font-size:12px; padding:6px 11px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px;
      background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); cursor:pointer; }
    .sheetedit-chart-type.sel { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:#fff; }
    .sheetedit-chart-field { display:flex; flex-direction:column; gap:4px; margin-bottom:10px; }
    .sheetedit-chart-field > span { color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
    .sheetedit-chart-field input[type=text] { font:inherit; background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047);
      border-radius:5px; color:var(--sheetedit-text, #e7eaf0); padding:6px 8px; }
    .sheetedit-chart-checks { display:flex; flex-wrap:wrap; gap:14px; margin-bottom:12px; font-size:13px; }
    .sheetedit-chart-checks label { display:flex; align-items:center; gap:6px; }
    .sheetedit-chart-preview { height:220px; background:#fff; border-radius:6px; padding:6px; margin-bottom:12px; }
    .sheetedit-chart-actions { display:flex; justify-content:flex-end; gap:8px; }
    .sheetedit-chart-btn { font:inherit; font-size:13px; padding:6px 14px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px;
      background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); cursor:pointer; }
    .sheetedit-chart-btn.primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:#fff; }
    .sheetedit-chart-editbar { position:fixed; z-index:30; display:flex; align-items:center; gap:6px; padding:5px 7px;
      background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
    .sheetedit-chart-editbar[hidden] { display:none; }
    .sheetedit-chart-editbar select, .sheetedit-chart-editbar button { font:inherit; font-size:12px; background:var(--sheetedit-btn, #3a3f47);
      color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:5px; padding:4px 8px; cursor:pointer; }
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

export function setupChartUi(deps: ChartUiDeps): { openInsert(rect: Rect): void; showEdit(model: ChartModel): void; hideEdit(): void; teardown(): void } {
  injectStyles();

  // ---- Insert dialog ----
  function openInsert(sel: Rect): void {
    const wb = deps.getWorkbook();
    const sheetName = deps.activeSheetName();
    let kind: ChartKind = "column";
    const state = { firstRowHeader: true, firstColLabels: true, showLegend: true, dataLabels: false, stacked: false, percent: false, comboLine: false, comboSecondary: false, title: "" };

    const modal = document.createElement("div");
    modal.className = "sheetedit-chart-modal";
    const card = document.createElement("div");
    card.className = "sheetedit-chart-card";
    const h = document.createElement("h3");
    h.textContent = t("chartInsert");
    const types = document.createElement("div");
    types.className = "sheetedit-chart-types";
    const typeBtns = KINDS.map(({ kind: k, key }) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-chart-type" + (k === kind ? " sel" : "");
      b.textContent = t(key);
      b.addEventListener("click", () => { kind = k; typeBtns.forEach((x) => x.classList.toggle("sel", x === b)); redraw(); });
      types.appendChild(b);
      return b;
    });

    const rangeField = document.createElement("label");
    rangeField.className = "sheetedit-chart-field";
    const rangeLbl = document.createElement("span");
    rangeLbl.textContent = t("chartRange");
    const rangeInput = document.createElement("input");
    rangeInput.type = "text";
    rangeInput.value = `${sheetName}!${cellA1(sel.c1, sel.r1)}:${cellA1(sel.c2, sel.r2)}`;
    rangeInput.addEventListener("input", redraw);
    rangeField.append(rangeLbl, rangeInput);

    const titleField = document.createElement("label");
    titleField.className = "sheetedit-chart-field";
    const titleLbl = document.createElement("span");
    titleLbl.textContent = t("chartTitle");
    const titleInput = document.createElement("input");
    titleInput.type = "text";
    titleInput.addEventListener("input", () => { state.title = titleInput.value; redraw(); });
    titleField.append(titleLbl, titleInput);

    const checks = document.createElement("div");
    checks.className = "sheetedit-chart-checks";
    type BoolKey = "firstRowHeader" | "firstColLabels" | "showLegend" | "dataLabels" | "stacked" | "percent" | "comboLine" | "comboSecondary";
    const mkCheck = (label: string, key: BoolKey): HTMLLabelElement => {
      const l = document.createElement("label");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = state[key];
      cb.addEventListener("change", () => { state[key] = cb.checked; redraw(); });
      const sp = document.createElement("span");
      sp.textContent = label;
      l.append(cb, sp);
      return l;
    };
    checks.append(mkCheck(t("chartFirstRow"), "firstRowHeader"), mkCheck(t("chartFirstCol"), "firstColLabels"), mkCheck(t("chartLegend"), "showLegend"), mkCheck(t("chartDataLabels"), "dataLabels"), mkCheck(t("chartStacked"), "stacked"), mkCheck(t("chartPercent"), "percent"), mkCheck(t("chartComboLine"), "comboLine"), mkCheck(t("chartComboSecondary"), "comboSecondary"));

    const preview = document.createElement("div");
    preview.className = "sheetedit-chart-preview";
    const canvas = document.createElement("canvas");
    preview.appendChild(canvas);

    const actions = document.createElement("div");
    actions.className = "sheetedit-chart-actions";
    const cancel = document.createElement("button");
    cancel.className = "sheetedit-chart-btn";
    cancel.textContent = t("chartCancel");
    const insert = document.createElement("button");
    insert.className = "sheetedit-chart-btn primary";
    insert.textContent = t("chartInsertBtn");
    actions.append(cancel, insert);

    card.append(h, types, rangeField, titleField, checks, preview, actions);
    modal.appendChild(card);
    deps.wrap.appendChild(modal);

    let chart: InstanceType<ChartCtor> | null = null;
    function buildModel(): ChartModel | null {
      const parsed = parseRange(rangeInput.value, sheetName);
      if (!parsed) return null;
      const model = buildChart(parsed.sheet, kind, parsed.rect, { firstRowHeader: state.firstRowHeader, firstColLabels: state.firstColLabels }, "preview", defaultAnchor(parsed.rect));
      model.title = state.title || undefined;
      model.legend = { show: state.showLegend, pos: "bottom" };
      model.dataLabels = state.dataLabels;
      model.stacked = state.stacked || state.percent || undefined; model.percent = state.percent || undefined;
      if (["column","bar","line","area"].includes(model.kind) && model.series.length >= 2) { const last = model.series[model.series.length - 1]; if (state.comboLine) last.type = "line"; if (state.comboSecondary) last.secondaryAxis = true; }
      return model;
    }
    let previewSeq = 0;
    function redraw(): void {
      const model = buildModel();
      if (!model) return;
      const seq = ++previewSeq;
      void loadChartJs().then((Ctor) => {
        if (seq !== previewSeq) return; // a newer redraw superseded this one
        // Chart.js can't change its type via update(); recreate on every redraw (a preview, so cheap).
        chart?.destroy();
        chart = new Ctor(canvas.getContext("2d")!, chartConfig(model, wb));
      });
    }
    const close = (): void => { chart?.destroy(); modal.remove(); };
    cancel.addEventListener("click", close);
    modal.addEventListener("mousedown", (e) => { if (e.target === modal) close(); });
    insert.addEventListener("click", () => {
      const parsed = parseRange(rangeInput.value, sheetName);
      if (!parsed) return;
      const model = buildChart(parsed.sheet, kind, parsed.rect, { firstRowHeader: state.firstRowHeader, firstColLabels: state.firstColLabels }, uniqueId(), defaultAnchor(parsed.rect));
      model.title = state.title || undefined;
      model.legend = { show: state.showLegend, pos: "bottom" };
      model.dataLabels = state.dataLabels;
      model.stacked = state.stacked || state.percent || undefined; model.percent = state.percent || undefined;
      if (["column","bar","line","area"].includes(model.kind) && model.series.length >= 2) { const last = model.series[model.series.length - 1]; if (state.comboLine) last.type = "line"; if (state.comboSecondary) last.secondaryAxis = true; }
      close();
      deps.onCreate(model);
    });
    redraw();
  }

  // ---- Edit toolbar on a selected chart ----
  const editbar = document.createElement("div");
  editbar.className = "sheetedit-chart-editbar";
  editbar.hidden = true;
  const typeSel = document.createElement("select");
  for (const { kind, key } of KINDS) { const o = document.createElement("option"); o.value = kind; o.textContent = t(key); typeSel.appendChild(o); }
  const labelsToggle = document.createElement("label");
  labelsToggle.style.cssText = "display:flex;align-items:center;gap:4px;font-size:12px;cursor:pointer;";
  const labelsCb = document.createElement("input");
  labelsCb.type = "checkbox";
  const labelsTxt = document.createElement("span");
  labelsTxt.textContent = t("chartDataLabels");
  labelsToggle.append(labelsCb, labelsTxt);
  const delBtn = document.createElement("button");
  delBtn.textContent = t("chartDelete");
  editbar.append(typeSel, labelsToggle, delBtn);
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
    // Prefer just above the chart, but never above the grid (which would cover the main toolbar):
    // in that case sit just inside the chart's top edge.
    const above = r.top - editbar.offsetHeight - 4;
    editbar.style.top = `${(above >= gr.top ? above : r.top + 4) - wr.top}px`;
  };
  typeSel.addEventListener("change", () => { if (editing) { editing.kind = typeSel.value as ChartKind; editing.dirty = true; deps.onChange(editing); } });
  labelsCb.addEventListener("change", () => { if (editing) { editing.dataLabels = labelsCb.checked; editing.dirty = true; deps.onChange(editing); } });
  delBtn.addEventListener("click", () => { if (editing) { const m = editing; hideEdit(); deps.onDelete(m); } });
  deps.gridScroll.addEventListener("scroll", positionEditbar, { passive: true });

  function showEdit(model: ChartModel): void { editing = model; typeSel.value = model.kind; labelsCb.checked = !!model.dataLabels; positionEditbar(); }
  function hideEdit(): void { editing = null; editbar.hidden = true; }

  return { openInsert, showEdit, hideEdit, teardown() { editbar.remove(); deps.gridScroll.removeEventListener("scroll", positionEditbar); } };
}

let idSeq = 0;
const uniqueId = (): string => `chart-new-${++idSeq}`;
const cellA1 = (col: number, row: number): string => `${colLetters(col)}${row}`;
function colLetters(c: number): string { let s = ""; while (c > 0) { s = String.fromCharCode(65 + ((c - 1) % 26)) + s; c = Math.floor((c - 1) / 26); } return s; }
