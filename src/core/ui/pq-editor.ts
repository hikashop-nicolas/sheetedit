import { t } from "../i18n";
import { buildPqHost } from "./pq-host";
import { TRANSFORMS, strLit, nameList, quoteName, type TransformSpec, type TfField } from "./pq-transforms";
import { listWorkbookTables } from "../../adapters/xlsx/tables";
import { TRANSFORM_ICONS, APPEND_ICON, MERGE_ICON, LOAD_ICON, CANCEL_ICON, SAVE_ICON, NEWQUERY_ICON, svgIcon } from "./pq-icons";
import type { Workbook } from "../model";
import type { MValue } from "mlang";

/** Set a button's content to an SVG icon followed by a text label. */
function iconLabel(btn: HTMLButtonElement, inner: string, label: string): void {
  btn.textContent = "";
  btn.appendChild(svgIcon(inner));
  const span = document.createElement("span");
  span.textContent = label;
  btn.appendChild(span);
}

// Full-window Power Query editor: a queries pane, an Applied Steps pane, a live preview grid
// and a formula bar. A query is a `let` expression; its steps are the let bindings and the `in`
// clause names the returned one. Selecting a step evaluates the query up to that point through
// the workbook-backed host connectors and shows the result (row-capped). All edits accumulate
// in an in-memory draft of Section1.m; Save & Close writes it back via qdeff. mlang and the
// step API are lazy-imported, so the base editor bundle never carries them.

type MTable = Extract<MValue, { kind: "table" }>;

const PREVIEW_ROWS = 1000; // Excel-style preview cap

export interface QueryEditorDeps {
  wrap: HTMLElement; // the editor chrome (overlay parent)
  wb: Workbook;
  attachedFiles: Record<string, Uint8Array>;
  /** Persist an edited Section1.m into the workbook (rewrites the DataMashup blob) and mark dirty. */
  save(newSectionM: string): void;
  /** Load a query result into the workbook (existing destination table, or a new sheet) and
      refresh the grid. Returns where it landed. Absent hosts hide the Load button. */
  loadQuery?(name: string, result: MTable): { sheetName: string; rows: number };
  /** Called after a successful Save & Close, so the host can relist/refresh. */
  onSaved?(): void;
}

const STYLE_ID = "sheetedit-pqe-style";

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .se-pqe { position:fixed; inset:0; z-index:60; display:flex; flex-direction:column;
      background:var(--sheetedit-bg, #1f2227); color:var(--sheetedit-text, #e6e6e6);
      font:13px system-ui, sans-serif; }
    .se-pqe[hidden] { display:none; }
    .se-pqe-bar { display:flex; align-items:center; gap:10px; padding:8px 12px;
      background:var(--sheetedit-chrome, #2b2f36); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
    .se-pqe-title { font-weight:600; }
    .se-pqe-spacer { flex:1; }
    .se-pqe-btn { font:inherit; font-size:13px; display:inline-flex; align-items:center; gap:6px; background:var(--sheetedit-btn, #3a3f47);
      color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57);
      border-radius:6px; padding:5px 12px; cursor:pointer; }
    .se-pqe-btn svg { display:block; }
    .se-pqe-btn:hover:not(:disabled) { background:var(--sheetedit-btn-hover, #454b54); }
    .se-pqe-btn:disabled { opacity:.5; cursor:default; }
    .se-pqe-btn.primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:#fff; }
    .se-pqe-ribbon { display:flex; flex-wrap:wrap; align-items:center; gap:3px; padding:5px 10px;
      background:var(--sheetedit-chrome, #2b2f36); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
    .se-pqe-rbtn { font:inherit; font-size:12px; display:inline-flex; align-items:center; gap:6px; background:transparent; color:var(--sheetedit-text, #e6e6e6);
      border:1px solid transparent; border-radius:5px; padding:4px 9px; cursor:pointer; white-space:nowrap; }
    .se-pqe-rbtn svg { display:block; flex:none; color:var(--sheetedit-accent, #6e7bff); }
    .se-pqe-rbtn:hover:not(:disabled) { background:var(--sheetedit-btn, #3a3f47); border-color:var(--sheetedit-btn-border, #4a4f57); }
    .se-pqe-rbtn:disabled { opacity:.4; cursor:default; }
    .se-pqe-rsep { width:1px; align-self:stretch; background:var(--sheetedit-border, #1c1f24); margin:2px 4px; }
    .se-pqe-rgroup { color:var(--sheetedit-muted, #aab2bf); font-size:10px; text-transform:uppercase; letter-spacing:.05em; padding:0 4px 0 2px; }
    .se-pqe-modal { position:absolute; inset:0; z-index:5; display:flex; align-items:center; justify-content:center; background:rgba(0,0,0,.45); }
    .se-pqe-modal[hidden] { display:none; }
    .se-pqe-card { width:min(420px,92%); max-height:80%; overflow:auto; background:var(--sheetedit-chrome, #2b2f36);
      border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 12px 40px rgba(0,0,0,.5); padding:16px; }
    .se-pqe-card h3 { margin:0 0 12px; font-size:15px; }
    .se-pqe-field { display:flex; flex-direction:column; gap:4px; margin-bottom:11px; }
    .se-pqe-field > span { color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
    .se-pqe-field input, .se-pqe-field select { font:inherit; background:var(--sheetedit-border, #1c1f24);
      border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); padding:5px 8px; }
    .se-pqe-checks { display:flex; flex-direction:column; gap:3px; max-height:180px; overflow:auto;
      border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; padding:6px 8px; background:var(--sheetedit-border, #1c1f24); }
    .se-pqe-checks label { display:flex; align-items:center; gap:7px; font-size:13px; }
    .se-pqe-card-actions { display:flex; justify-content:flex-end; gap:8px; margin-top:6px; }
    .se-pqe-fx { display:flex; align-items:stretch; gap:6px; padding:6px 12px;
      background:var(--sheetedit-chrome2, #23262c); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
    .se-pqe-fx-lbl { align-self:center; color:var(--sheetedit-muted, #aab2bf); font:12px ui-monospace,monospace; }
    .se-pqe-fx textarea { flex:1; min-height:26px; max-height:120px; resize:vertical;
      background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047);
      border-radius:5px; color:var(--sheetedit-text, #e7eaf0); font:13px ui-monospace,monospace; padding:5px 8px; }
    .se-pqe-main { flex:1; min-height:0; display:flex; }
    .se-pqe-queries { width:200px; flex:none; overflow:auto; border-right:1px solid var(--sheetedit-border, #1c1f24); }
    .se-pqe-settings { width:240px; flex:none; overflow:auto; border-left:1px solid var(--sheetedit-border, #1c1f24); }
    .se-pqe-center { flex:1; min-width:0; display:flex; flex-direction:column; }
    .se-pqe-pane-h { padding:7px 12px; font-weight:600; color:var(--sheetedit-muted, #aab2bf);
      font-size:12px; text-transform:uppercase; letter-spacing:.04em; border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
    .se-pqe-pane-h-row { display:flex; align-items:center; justify-content:space-between; }
    .se-pqe-newq { display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; padding:0; cursor:pointer;
      background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6);
      border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:5px; }
    .se-pqe-newq svg { display:block; }
    .se-pqe-newq:hover { background:var(--sheetedit-btn-hover, #454b54); }
    .se-pqe-item { display:flex; align-items:center; gap:6px; padding:7px 12px; cursor:pointer; border-bottom:1px solid rgba(0,0,0,.12); }
    .se-pqe-item:hover { background:var(--sheetedit-btn, #3a3f47); }
    .se-pqe-item.sel { background:var(--sheetedit-accent, #6e7bff); color:#fff; }
    .se-pqe-item-name { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    .se-pqe-item-x { opacity:0; border:0; background:none; color:inherit; cursor:pointer; font-size:15px; line-height:1; padding:0 2px; border-radius:4px; }
    .se-pqe-item:hover .se-pqe-item-x { opacity:.7; }
    .se-pqe-item-x:hover { opacity:1 !important; background:rgba(255,255,255,.15); }
    .se-pqe-name-in { width:100%; box-sizing:border-box; font:inherit; background:var(--sheetedit-border, #1c1f24);
      border:1px solid var(--sheetedit-accent, #6e7bff); border-radius:4px; color:inherit; padding:3px 6px; }
    .se-pqe-preview { flex:1; min-height:0; overflow:auto; background:#e9e9ec; }
    .se-pqe-ptable { border-collapse:collapse; font:13px/1.3 ui-sans-serif, system-ui, sans-serif; color:#1a1a1a; }
    .se-pqe-ptable th, .se-pqe-ptable td { border:1px solid #d4d4d8; padding:3px 9px; text-align:left; white-space:nowrap; }
    .se-pqe-ptable th { position:sticky; top:0; background:#f1f1f4; color:#333; font-weight:600; z-index:1; }
    .se-pqe-ptable th .nm { display:block; }
    .se-pqe-ptable th .ty { display:block; font-weight:400; font-size:10px; color:#888; }
    .se-pqe-ptable th .qbar { display:flex; height:3px; margin-top:3px; border-radius:2px; overflow:hidden; background:#d4d4d8; }
    .se-pqe-ptable th .qv { background:#3fb950; }
    .se-pqe-ptable th .qe { background:#c0c4cc; }
    .se-pqe-ptable th .qx { background:#d33d3d; }
    .se-pqe-ptable td.num { text-align:right; font-variant-numeric:tabular-nums; }
    .se-pqe-ptable td.null, .se-pqe-ptable td.obj { color:#8a8f98; }
    .se-pqe-ptable td.err { color:#c0392b; }
    .se-pqe-ptable tr:nth-child(even) td { background:#f6f6f8; }
    .se-pqe-scalar { padding:16px; font:14px ui-monospace,monospace; color:var(--sheetedit-text, #e6e6e6); }
    .se-pqe-foot { display:flex; align-items:center; gap:14px; padding:5px 12px; color:var(--sheetedit-muted, #aab2bf);
      font-size:12px; background:var(--sheetedit-chrome2, #23262c); border-top:1px solid var(--sheetedit-border, #1c1f24); }
    .se-pqe-foot .err { color:#ff8a8a; }
    .se-pqe-empty { padding:24px; color:var(--sheetedit-muted, #aab2bf); }
  `;
  document.head.appendChild(s);
}

const pad = (n: number, w = 2): string => String(Math.abs(n)).padStart(w, "0");
function hms(secs: number): string {
  const s = Math.floor(secs); return `${pad(Math.floor(s / 3600))}:${pad(Math.floor((s % 3600) / 60))}:${pad(s % 60)}`;
}
/** Compact display of a preview cell (Excel shows nested table/list/record as objects). */
function cellText(v: MValue): { text: string; cls: string } {
  switch (v.kind) {
    case "null": return { text: "null", cls: "null" };
    case "logical": return { text: v.value ? "TRUE" : "FALSE", cls: "" };
    case "number": return { text: v.big !== undefined ? v.big.toString() : String(v.value), cls: "num" };
    case "text": return { text: v.value, cls: "" };
    case "date": return { text: `${pad(v.y, 4)}-${pad(v.m)}-${pad(v.d)}`, cls: "" };
    case "time": return { text: hms(v.secs), cls: "" };
    case "datetime": return { text: `${pad(v.y, 4)}-${pad(v.m)}-${pad(v.d)} ${hms(v.secs)}`, cls: "" };
    case "datetimezone": { const o = v.offset; const sg = o < 0 ? "-" : "+"; return { text: `${pad(v.y, 4)}-${pad(v.m)}-${pad(v.d)} ${hms(v.secs)} ${sg}${pad(Math.floor(Math.abs(o) / 60))}:${pad(Math.abs(o) % 60)}`, cls: "" }; }
    case "duration": return { text: `${(v.secs / 86400).toFixed(6)}d`, cls: "num" };
    case "binary": return { text: "[Binary]", cls: "obj" };
    case "list": return { text: "[List]", cls: "obj" };
    case "record": return { text: "[Record]", cls: "obj" };
    case "table": return { text: "[Table]", cls: "obj" };
    case "error": return { text: "Error", cls: "err" };
    default: return { text: `[${v.kind}]`, cls: "obj" };
  }
}

export function setupQueryEditor(deps: QueryEditorDeps): { open(sectionM: string): void } {
  const { wrap, wb, attachedFiles } = deps;
  injectStyles();

  const overlay = document.createElement("div");
  overlay.className = "se-pqe";
  overlay.hidden = true;

  // Title bar
  const bar = document.createElement("div");
  bar.className = "se-pqe-bar";
  const title = document.createElement("span");
  title.className = "se-pqe-title";
  title.textContent = t("pqEditorTitle");
  const spacer = document.createElement("span");
  spacer.className = "se-pqe-spacer";
  const loadBtn = document.createElement("button");
  loadBtn.className = "se-pqe-btn";
  iconLabel(loadBtn, LOAD_ICON, t("pqLoad"));
  loadBtn.title = t("pqLoadTitle");
  loadBtn.hidden = !deps.loadQuery;
  const saveBtn = document.createElement("button");
  saveBtn.className = "se-pqe-btn primary";
  iconLabel(saveBtn, SAVE_ICON, t("pqSaveClose"));
  const cancelBtn = document.createElement("button");
  cancelBtn.className = "se-pqe-btn";
  iconLabel(cancelBtn, CANCEL_ICON, t("pqCancel"));
  bar.append(title, spacer, loadBtn, cancelBtn, saveBtn);

  // Transform ribbon: each button appends a step to the query's final result.
  const ribbon = document.createElement("div");
  ribbon.className = "se-pqe-ribbon";
  const ribbonButtons: HTMLButtonElement[] = [];
  {
    let lastGroup = "";
    for (const spec of TRANSFORMS) {
      if (spec.group !== lastGroup) {
        if (lastGroup) { const sep = document.createElement("span"); sep.className = "se-pqe-rsep"; ribbon.appendChild(sep); }
        const g = document.createElement("span"); g.className = "se-pqe-rgroup"; g.textContent = spec.group; ribbon.appendChild(g);
        lastGroup = spec.group;
      }
      const b = document.createElement("button");
      b.className = "se-pqe-rbtn";
      b.type = "button";
      iconLabel(b, TRANSFORM_ICONS[spec.id] ?? "", spec.label);
      b.addEventListener("click", () => void applyTransform(spec));
      ribbonButtons.push(b);
      ribbon.appendChild(b);
    }
    // Combine group: cross-query operations (handled specially, not plain Table.* on one input).
    const sep = document.createElement("span"); sep.className = "se-pqe-rsep"; ribbon.appendChild(sep);
    const cg = document.createElement("span"); cg.className = "se-pqe-rgroup"; cg.textContent = t("pqCombine"); ribbon.appendChild(cg);
    for (const [icon, label, fn] of [[APPEND_ICON, t("pqAppend"), () => void appendQueries()], [MERGE_ICON, t("pqMerge"), () => void mergeQueries()]] as const) {
      const b = document.createElement("button");
      b.className = "se-pqe-rbtn";
      b.type = "button";
      iconLabel(b, icon, label);
      b.addEventListener("click", fn);
      ribbonButtons.push(b);
      ribbon.appendChild(b);
    }
  }

  // Formula bar
  const fx = document.createElement("div");
  fx.className = "se-pqe-fx";
  const fxLbl = document.createElement("span");
  fxLbl.className = "se-pqe-fx-lbl";
  fxLbl.textContent = "=";
  const fxArea = document.createElement("textarea");
  fxArea.rows = 1;
  fxArea.spellcheck = false;
  fx.append(fxLbl, fxArea);

  // Panes
  const main = document.createElement("div");
  main.className = "se-pqe-main";
  const queriesPane = document.createElement("div");
  queriesPane.className = "se-pqe-queries";
  const queriesHead = document.createElement("div");
  queriesHead.className = "se-pqe-pane-h se-pqe-pane-h-row";
  const queriesHeadLbl = document.createElement("span");
  queriesHeadLbl.textContent = t("pqQueries");
  const newQueryBtn = document.createElement("button");
  newQueryBtn.className = "se-pqe-newq";
  newQueryBtn.type = "button";
  newQueryBtn.appendChild(svgIcon(NEWQUERY_ICON, 14));
  newQueryBtn.title = t("pqNewQuery");
  newQueryBtn.addEventListener("click", () => void getData());
  queriesHead.append(queriesHeadLbl, newQueryBtn);
  const queriesList = document.createElement("div");
  queriesPane.append(queriesHead, queriesList);

  const center = document.createElement("div");
  center.className = "se-pqe-center";
  const preview = document.createElement("div");
  preview.className = "se-pqe-preview";
  const foot = document.createElement("div");
  foot.className = "se-pqe-foot";
  center.append(preview, foot);

  const settingsPane = document.createElement("div");
  settingsPane.className = "se-pqe-settings";
  const settingsHead = document.createElement("div");
  settingsHead.className = "se-pqe-pane-h";
  settingsHead.textContent = t("pqAppliedSteps");
  const stepsList = document.createElement("div");
  settingsPane.append(settingsHead, stepsList);

  main.append(queriesPane, center, settingsPane);

  // Transform dialog (a modal within the overlay).
  const modal = document.createElement("div");
  modal.className = "se-pqe-modal";
  modal.hidden = true;

  overlay.append(bar, ribbon, fx, main, modal);
  wrap.appendChild(overlay);

  // ---- state ----
  let draft = "";
  let queryNames: string[] = [];
  let curQuery: string | null = null;
  let curStep: string | null = null;
  let curInTarget: string | null = null; // the query's returned (final) step
  let steps: { name: string; rawName: string; expression: string }[] = [];
  let previewColumns: string[] = []; // columns of the last table preview (for transform pickers)
  let previewToken = 0;

  const setRibbonEnabled = (on: boolean): void => { for (const b of ribbonButtons) b.disabled = !on; };

  function close(): void {
    overlay.hidden = true;
  }

  async function renderQueries(): Promise<void> {
    const { evaluateSection } = await import("mlang");
    try {
      queryNames = (await evaluateSection(draft, {})).names;
    } catch {
      queryNames = [];
    }
    queriesList.textContent = "";
    if (queryNames.length === 0) {
      const e = document.createElement("div");
      e.className = "se-pqe-empty";
      e.textContent = t("pqNoQueries");
      queriesList.appendChild(e);
      return;
    }
    for (const name of queryNames) {
      const item = document.createElement("div");
      item.className = "se-pqe-item" + (name === curQuery ? " sel" : "");
      const label = document.createElement("span");
      label.className = "se-pqe-item-name";
      label.textContent = name;
      const del = document.createElement("button");
      del.className = "se-pqe-item-x";
      del.textContent = "×";
      del.title = t("pqDeleteQuery");
      del.addEventListener("click", (ev) => { ev.stopPropagation(); void deleteQuery(name); });
      item.append(label, del);
      item.addEventListener("click", () => void selectQuery(name));
      label.addEventListener("dblclick", (ev) => { ev.stopPropagation(); beginRenameQuery(item, label, name); });
      queriesList.appendChild(item);
    }
  }

  async function selectQuery(name: string): Promise<void> {
    curQuery = name;
    const { parseMemberSteps } = await import("mlang/steps");
    try {
      const parsed = await parseMemberSteps(draft, name);
      steps = parsed.steps;
      curInTarget = parsed.inTarget;
      // Select the returned step (the `in` target) by default, like Excel.
      curStep = parsed.steps.find((s) => s.name === parsed.inTarget)?.name ?? parsed.steps[parsed.steps.length - 1]?.name ?? null;
    } catch (e) {
      steps = [];
      curStep = null;
      curInTarget = null;
      foot.innerHTML = `<span class="err">${escapeHtml((e as Error).message)}</span>`;
    }
    setRibbonEnabled(steps.length > 0);
    await renderQueries();
    renderSteps();
    if (curStep) await selectStep(curStep);
  }

  function renderSteps(): void {
    stepsList.textContent = "";
    for (const step of steps) {
      const item = document.createElement("div");
      item.className = "se-pqe-item" + (step.name === curStep ? " sel" : "");
      const label = document.createElement("span");
      label.className = "se-pqe-item-name";
      label.textContent = step.name;
      label.title = step.name;
      const del = document.createElement("button");
      del.className = "se-pqe-item-x";
      del.textContent = "×";
      del.title = t("pqStepDelete");
      del.addEventListener("click", (ev) => { ev.stopPropagation(); void deleteStep(step.name); });
      item.append(label, del);
      item.addEventListener("click", () => void selectStep(step.name));
      // Double-click the name to rename the step.
      label.addEventListener("dblclick", (ev) => { ev.stopPropagation(); beginRename(item, label, step.name); });
      stepsList.appendChild(item);
    }
  }

  async function selectStep(name: string): Promise<void> {
    curStep = name;
    const step = steps.find((s) => s.name === name);
    fxArea.value = step?.expression ?? "";
    renderSteps();
    await runPreview(name);
  }

  async function runPreview(uptoStep: string): Promise<void> {
    if (!curQuery) return;
    const token = ++previewToken;
    foot.textContent = t("pqPreviewing");
    preview.textContent = "";
    try {
      const { previewSection } = await import("mlang/steps");
      const { evaluateSection, isMissingConnector, missingConnectorName } = await import("mlang");
      const section = await previewSection(draft, curQuery, uptoStep);
      const host = await buildPqHost({ wb, attachedFiles });
      let result: MValue;
      try {
        result = await evaluateSection(section, host).then((s) => s.run(curQuery!));
      } catch (e) {
        if (isMissingConnector(e)) { if (token === previewToken) foot.innerHTML = `<span class="err">${escapeHtml(t("pqPreviewExternal", { connector: missingConnectorName(e) }))}</span>`; return; }
        throw e;
      }
      if (token !== previewToken) return; // a newer preview superseded this one
      renderPreview(result);
    } catch (e) {
      if (token === previewToken) foot.innerHTML = `<span class="err">${escapeHtml(t("pqPreviewError", { msg: (e as Error).message }))}</span>`;
    }
  }

  function renderPreview(result: MValue): void {
    preview.textContent = "";
    previewColumns = result.kind === "table" ? result.columns : [];
    if (result.kind !== "table") {
      const box = document.createElement("div");
      box.className = "se-pqe-scalar";
      box.textContent = cellText(result).text;
      preview.appendChild(box);
      foot.textContent = t("pqPreviewValue", { kind: result.kind });
      return;
    }
    const table = result as MTable;
    const total = table.rows.length;
    const shown = Math.min(total, PREVIEW_ROWS);
    // Column quality (Excel-style): tally valid / empty / error over the shown rows.
    const quality = table.columns.map(() => ({ valid: 0, empty: 0, error: 0 }));
    for (let r = 0; r < shown; r++) {
      const row = table.rows[r];
      for (let c = 0; c < table.columns.length; c++) {
        const k = (row[c] ?? { kind: "null" }).kind;
        quality[c][k === "null" ? "empty" : k === "error" ? "error" : "valid"]++;
      }
    }
    const tbl = document.createElement("table");
    tbl.className = "se-pqe-ptable";
    const thead = document.createElement("thead");
    const htr = document.createElement("tr");
    table.columns.forEach((col, c) => {
      const th = document.createElement("th");
      const nm = document.createElement("span");
      nm.className = "nm";
      nm.textContent = col;
      th.appendChild(nm);
      const ty = table.types?.get(col);
      if (ty) { const s = document.createElement("span"); s.className = "ty"; s.textContent = ty.ascription ?? ty.name; th.appendChild(s); }
      // Quality bar: green valid, grey empty, red error (proportional over the preview).
      const q = quality[c];
      const n = Math.max(1, q.valid + q.empty + q.error);
      const bar = document.createElement("span");
      bar.className = "qbar";
      bar.title = `${q.valid} valid · ${q.empty} empty · ${q.error} error`;
      for (const [cls, v] of [["v", q.valid], ["e", q.empty], ["x", q.error]] as const) {
        if (v === 0) continue;
        const seg = document.createElement("span");
        seg.className = `q${cls}`;
        seg.style.width = `${(v / n) * 100}%`;
        bar.appendChild(seg);
      }
      th.appendChild(bar);
      htr.appendChild(th);
    });
    thead.appendChild(htr);
    tbl.appendChild(thead);
    const tbody = document.createElement("tbody");
    for (let r = 0; r < shown; r++) {
      const tr = document.createElement("tr");
      const row = table.rows[r];
      for (let c = 0; c < table.columns.length; c++) {
        const td = document.createElement("td");
        const { text, cls } = cellText(row[c] ?? { kind: "null" });
        if (cls) td.className = cls;
        td.textContent = text;
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    tbl.appendChild(tbody);
    preview.appendChild(tbl);
    foot.textContent = t("pqPreviewCols", { cols: table.columns.length }) + "  ·  " +
      (total > shown ? t("pqPreviewCapped", { shown, total: PREVIEW_ROWS }) : t("pqPreviewRows", { rows: total }));
  }

  // ---- edits (mutate the draft, then re-parse/re-preview) ----
  async function commitFormula(): Promise<void> {
    if (!curQuery || !curStep) return;
    const newExpr = fxArea.value.trim();
    const step = steps.find((s) => s.name === curStep);
    if (!step || newExpr === step.expression) return;
    try {
      const { replaceStepExpression, parseMemberSteps } = await import("mlang/steps");
      draft = await replaceStepExpression(draft, curQuery, curStep, newExpr);
      steps = (await parseMemberSteps(draft, curQuery)).steps;
      await runPreview(curStep);
    } catch (e) {
      foot.innerHTML = `<span class="err">${escapeHtml((e as Error).message)}</span>`;
    }
  }

  async function deleteStep(name: string): Promise<void> {
    if (!curQuery) return;
    try {
      const { removeStep, parseMemberSteps } = await import("mlang/steps");
      draft = await removeStep(draft, curQuery, name);
      const parsed = await parseMemberSteps(draft, curQuery);
      steps = parsed.steps;
      curInTarget = parsed.inTarget;
      curStep = parsed.steps.find((s) => s.name === parsed.inTarget)?.name ?? parsed.steps[parsed.steps.length - 1]?.name ?? null;
      renderSteps();
      if (curStep) await selectStep(curStep);
    } catch (e) {
      foot.innerHTML = `<span class="err">${escapeHtml((e as Error).message)}</span>`;
    }
  }

  function beginRename(item: HTMLElement, label: HTMLElement, oldName: string): void {
    const input = document.createElement("input");
    input.className = "se-pqe-name-in";
    input.value = oldName;
    item.replaceChild(input, label);
    input.focus();
    input.select();
    let done = false;
    const finish = async (commit: boolean): Promise<void> => {
      if (done) return; done = true;
      const newName = input.value.trim();
      if (commit && newName && newName !== oldName && curQuery) {
        try {
          const { renameStep, parseMemberSteps } = await import("mlang/steps");
          draft = await renameStep(draft, curQuery, oldName, newName);
          steps = (await parseMemberSteps(draft, curQuery)).steps;
          if (curStep === oldName) curStep = newName;
        } catch (e) {
          foot.innerHTML = `<span class="err">${escapeHtml((e as Error).message)}</span>`;
        }
      }
      renderSteps();
    };
    input.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") { ev.preventDefault(); void finish(true); }
      else if (ev.key === "Escape") { ev.preventDefault(); void finish(false); }
    });
    input.addEventListener("blur", () => void finish(true));
  }

  function uniqueStepName(base: string): string {
    if (!steps.some((s) => s.name === base)) return base;
    for (let i = 1; ; i++) if (!steps.some((s) => s.name === `${base} ${i}`)) return `${base} ${i}`;
  }

  /** Render one dialog field and return a getter for its collected value. */
  function renderField(f: TfField, cols: string[]): { el: HTMLElement; get(): string | string[] } {
    const wrap = document.createElement("label");
    wrap.className = "se-pqe-field";
    const span = document.createElement("span");
    span.textContent = f.label;
    wrap.appendChild(span);
    if (f.type === "columns" && f.multi) {
      const box = document.createElement("div");
      box.className = "se-pqe-checks";
      const boxes: HTMLInputElement[] = [];
      for (const c of cols) {
        const l = document.createElement("label");
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = c;
        boxes.push(cb);
        const txt = document.createElement("span");
        txt.textContent = c;
        l.append(cb, txt);
        box.appendChild(l);
      }
      wrap.appendChild(box);
      return { el: wrap, get: () => boxes.filter((b) => b.checked).map((b) => b.value) };
    }
    if (f.type === "columns" || f.type === "select") {
      const sel = document.createElement("select");
      const opts = f.type === "columns" ? cols.map((c) => ({ value: c, label: c })) : (f.options ?? []);
      for (const o of opts) { const op = document.createElement("option"); op.value = o.value; op.textContent = o.label; sel.appendChild(op); }
      if (f.default) sel.value = f.default;
      wrap.appendChild(sel);
      return { el: wrap, get: () => sel.value };
    }
    const input = document.createElement("input");
    input.type = f.type === "number" ? "number" : "text";
    if (f.placeholder) input.placeholder = f.placeholder;
    if (f.default) input.value = f.default;
    wrap.appendChild(input);
    return { el: wrap, get: () => input.value };
  }

  function showFieldsDialog(dlgTitle: string, fields: TfField[], cols: string[]): Promise<Record<string, string | string[]> | null> {
    return new Promise((resolve) => {
      modal.textContent = "";
      const card = document.createElement("div");
      card.className = "se-pqe-card";
      const h = document.createElement("h3");
      h.textContent = dlgTitle;
      card.appendChild(h);
      const getters = fields.map((f) => { const r = renderField(f, cols); card.appendChild(r.el); return { key: f.key, get: r.get }; });
      const actions = document.createElement("div");
      actions.className = "se-pqe-card-actions";
      const cancel = document.createElement("button");
      cancel.className = "se-pqe-btn";
      cancel.textContent = t("pqCancel");
      const ok = document.createElement("button");
      ok.className = "se-pqe-btn primary";
      ok.textContent = t("pqApply");
      actions.append(cancel, ok);
      card.appendChild(actions);
      modal.appendChild(card);
      modal.hidden = false;
      const done = (v: Record<string, string | string[]> | null): void => { modal.hidden = true; modal.textContent = ""; resolve(v); };
      cancel.addEventListener("click", () => done(null));
      modal.addEventListener("click", (e) => { if (e.target === modal) done(null); });
      ok.addEventListener("click", () => { const v: Record<string, string | string[]> = {}; for (const g of getters) v[g.key] = g.get(); done(v); });
      (card.querySelector("input, select") as HTMLElement | null)?.focus();
    });
  }

  const prevRawName = (): string => steps.find((s) => s.name === curInTarget)?.rawName ?? curInTarget ?? "Source";

  /** Append an M expression as a new step of the current query and select it. */
  async function appendStepAndSelect(baseName: string, expr: string): Promise<void> {
    if (!curQuery) return;
    try {
      const { appendStep, parseMemberSteps } = await import("mlang/steps");
      draft = await appendStep(draft, curQuery, uniqueStepName(baseName), expr);
      const parsed = await parseMemberSteps(draft, curQuery);
      steps = parsed.steps;
      curInTarget = parsed.inTarget;
      renderSteps();
      await selectStep(parsed.steps[parsed.steps.length - 1].name);
    } catch (e) {
      foot.innerHTML = `<span class="err">${escapeHtml((e as Error).message)}</span>`;
    }
  }

  async function applyTransform(spec: TransformSpec): Promise<void> {
    if (!curQuery || !curInTarget) return;
    // Transforms append to the query's final result; make sure it's the previewed step so the
    // column pickers reflect the insertion point.
    if (curStep !== curInTarget) await selectStep(curInTarget);
    const values = await showFieldsDialog(spec.label, spec.fields(previewColumns), previewColumns);
    if (!values) return;
    await appendStepAndSelect(spec.stepName, spec.buildM(prevRawName(), values));
  }

  /** The output columns of another query (for merge key/expand pickers). */
  async function queryColumns(name: string): Promise<string[]> {
    try {
      const { previewSection } = await import("mlang/steps");
      const { evaluateSection } = await import("mlang");
      const section = await previewSection(draft, name);
      const r = await evaluateSection(section, await buildPqHost({ wb, attachedFiles })).then((s) => s.run(name));
      return r.kind === "table" ? r.columns : [];
    } catch { return []; }
  }

  async function appendQueries(): Promise<void> {
    if (!curQuery || !curInTarget) return;
    const others = queryNames.filter((n) => n !== curQuery);
    if (others.length === 0) { foot.innerHTML = `<span class="err">${escapeHtml(t("pqNeedTwoQueries"))}</span>`; return; }
    if (curStep !== curInTarget) await selectStep(curInTarget);
    const v = await showFieldsDialog(t("pqAppend"), [{ key: "other", label: t("pqOtherQuery"), type: "select", options: others.map((n) => ({ value: n, label: n })) }], []);
    if (!v) return;
    await appendStepAndSelect("Appended Query", `Table.Combine({${prevRawName()}, ${quoteName(first(v.other))}})`);
  }

  const JOIN_KINDS = [
    { value: "JoinKind.LeftOuter", label: "Left outer (all from first)" },
    { value: "JoinKind.RightOuter", label: "Right outer (all from second)" },
    { value: "JoinKind.FullOuter", label: "Full outer (all from both)" },
    { value: "JoinKind.Inner", label: "Inner (matching only)" },
    { value: "JoinKind.LeftAnti", label: "Left anti (non-matching from first)" },
    { value: "JoinKind.RightAnti", label: "Right anti (non-matching from second)" },
  ];

  async function mergeQueries(): Promise<void> {
    if (!curQuery || !curInTarget) return;
    const others = queryNames.filter((n) => n !== curQuery);
    if (others.length === 0) { foot.innerHTML = `<span class="err">${escapeHtml(t("pqNeedTwoQueries"))}</span>`; return; }
    if (curStep !== curInTarget) await selectStep(curInTarget);
    const pick = await showFieldsDialog(t("pqMerge"), [{ key: "other", label: t("pqOtherQuery"), type: "select", options: others.map((n) => ({ value: n, label: n })) }], []);
    if (!pick) return;
    const other = first(pick.other);
    const otherCols = await queryColumns(other);
    const v = await showFieldsDialog(t("pqMerge"), [
      { key: "thisKey", label: t("pqMergeThisKey"), type: "columns", multi: false },
      { key: "otherKey", label: t("pqMergeOtherKey"), type: "select", options: otherCols.map((c) => ({ value: c, label: c })) },
      { key: "join", label: t("pqMergeJoin"), type: "select", options: JOIN_KINDS },
    ], previewColumns);
    if (!v) return;
    const col = "Merged";
    const nested = `Table.NestedJoin(${prevRawName()}, ${nameList([first(v.thisKey)])}, ${quoteName(other)}, ${nameList([first(v.otherKey)])}, ${strLit(col)}, ${first(v.join)})`;
    await appendStepAndSelect("Merged Queries", nested);
    if (otherCols.length) {
      const prefixed = otherCols.map((c) => `${other}.${c}`);
      await appendStepAndSelect(`Expanded ${other}`, `Table.ExpandTableColumn(${quoteName("Merged Queries")}, ${strLit(col)}, ${nameList(otherCols)}, ${nameList(prefixed)})`);
    }
  }

  // ---- Get Data / query management ----
  const first = (v: string | string[] | undefined): string => (Array.isArray(v) ? (v[0] ?? "") : (v ?? ""));
  function uniqueQueryName(): string {
    for (let i = 1; ; i++) if (!queryNames.includes(`Query${i}`)) return `Query${i}`;
  }
  function buildSource(src: string, v: Record<string, string | string[]>): string {
    if (src === "table") return `let\n    Source = Excel.CurrentWorkbook(){[Name=${strLit(first(v.table))}]}[Content]\nin\n    Source`;
    if (src === "web") {
      const url = strLit(first(v.url));
      if (first(v.format) === "json") return `let\n    Source = Json.Document(Web.Contents(${url}))\nin\n    Source`;
      return `let\n    Source = Csv.Document(Web.Contents(${url}), [Delimiter=",", QuoteStyle=QuoteStyle.Csv]),\n    #"Promoted Headers" = Table.PromoteHeaders(Source)\nin\n    #"Promoted Headers"`;
    }
    return `let\n    Source = #table({"Column1"}, {})\nin\n    Source`;
  }
  async function getData(): Promise<void> {
    const tables = listWorkbookTables(wb).map((tb) => tb.displayName);
    const opts = [{ value: "table", label: t("pqFromTable") }, { value: "web", label: t("pqFromWeb") }, { value: "blank", label: t("pqBlank") }];
    const pick = await showFieldsDialog(t("pqGetData"), [{ key: "source", label: t("pqSource"), type: "select", options: opts }], []);
    if (!pick) return;
    const src = first(pick.source);
    const suggested = uniqueQueryName();
    const nameField: TfField = { key: "name", label: t("pqQueryName"), type: "text", default: suggested };
    let fields: TfField[];
    if (src === "table") fields = [nameField, { key: "table", label: t("pqSourceTable"), type: "select", options: tables.map((n) => ({ value: n, label: n })) }];
    else if (src === "web") fields = [nameField, { key: "url", label: t("pqSourceUrl"), type: "text", placeholder: "https://…" }, { key: "format", label: t("pqWebFormat"), type: "select", options: [{ value: "csv", label: "CSV" }, { value: "json", label: "JSON" }] }];
    else fields = [nameField];
    const v = await showFieldsDialog(t("pqGetData"), fields, []);
    if (!v) return;
    const qname = first(v.name).trim() || suggested;
    try {
      const { addMember } = await import("mlang/steps");
      draft = await addMember(draft, qname, buildSource(src, v), { shared: true });
      await renderQueries();
      await selectQuery(qname);
    } catch (e) {
      foot.innerHTML = `<span class="err">${escapeHtml((e as Error).message)}</span>`;
    }
  }
  async function deleteQuery(name: string): Promise<void> {
    try {
      const { removeMember } = await import("mlang/steps");
      draft = await removeMember(draft, name);
      if (curQuery === name) { curQuery = null; curStep = null; steps = []; preview.textContent = ""; foot.textContent = ""; setRibbonEnabled(false); }
      await renderQueries();
      if (curQuery === null && queryNames[0]) await selectQuery(queryNames[0]);
    } catch (e) {
      foot.innerHTML = `<span class="err">${escapeHtml((e as Error).message)}</span>`;
    }
  }
  function beginRenameQuery(item: HTMLElement, label: HTMLElement, oldName: string): void {
    const input = document.createElement("input");
    input.className = "se-pqe-name-in";
    input.value = oldName;
    item.replaceChild(input, label);
    input.focus();
    input.select();
    let done = false;
    const finish = async (commit: boolean): Promise<void> => {
      if (done) return; done = true;
      const newName = input.value.trim();
      if (commit && newName && newName !== oldName) {
        try {
          const { renameMember } = await import("mlang/steps");
          draft = await renameMember(draft, oldName, newName);
          if (curQuery === oldName) curQuery = newName;
        } catch (e) {
          foot.innerHTML = `<span class="err">${escapeHtml((e as Error).message)}</span>`;
        }
      }
      await renderQueries();
    };
    input.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") { ev.preventDefault(); void finish(true); }
      else if (ev.key === "Escape") { ev.preventDefault(); void finish(false); }
    });
    input.addEventListener("blur", () => void finish(true));
  }

  fxArea.addEventListener("keydown", (ev) => {
    // Enter commits, Shift+Enter inserts a newline (M can be multi-line).
    if (ev.key === "Enter" && !ev.shiftKey) { ev.preventDefault(); void commitFormula(); }
  });
  fxArea.addEventListener("blur", () => void commitFormula());

  // Load the current query's FULL result (uncapped) into the workbook, and persist its M so a
  // later refresh reproduces the same data.
  async function loadCurrent(): Promise<void> {
    if (!curQuery || !deps.loadQuery) return;
    loadBtn.disabled = true;
    foot.textContent = t("pqPreviewing");
    try {
      const { evaluateSection, isMissingConnector, missingConnectorName } = await import("mlang");
      let result: MValue;
      try {
        result = await evaluateSection(draft, await buildPqHost({ wb, attachedFiles })).then((s) => s.run(curQuery!));
      } catch (e) {
        foot.innerHTML = `<span class="err">${escapeHtml(isMissingConnector(e) ? t("pqPreviewExternal", { connector: missingConnectorName(e) }) : (e as Error).message)}</span>`;
        return;
      }
      if (result.kind !== "table") { foot.innerHTML = `<span class="err">${escapeHtml(t("pqLoadNotTable", { kind: result.kind }))}</span>`; return; }
      deps.save(draft);
      const { sheetName, rows } = deps.loadQuery(curQuery, result as MTable);
      foot.textContent = t("pqLoaded", { rows, sheet: sheetName });
    } finally {
      loadBtn.disabled = false;
    }
  }
  loadBtn.addEventListener("click", () => void loadCurrent());

  cancelBtn.addEventListener("click", close);
  saveBtn.addEventListener("click", () => {
    deps.save(draft);
    close();
    deps.onSaved?.();
  });
  document.addEventListener("keydown", (e) => {
    if (!overlay.hidden && e.key === "Escape" && document.activeElement === document.body) close();
  });

  return {
    open(sectionM: string): void {
      draft = sectionM;
      curQuery = null;
      curStep = null;
      steps = [];
      overlay.hidden = false;
      preview.textContent = "";
      foot.textContent = "";
      setRibbonEnabled(false);
      void renderQueries().then(() => { if (queryNames[0]) return selectQuery(queryNames[0]); });
    },
  };
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
}
