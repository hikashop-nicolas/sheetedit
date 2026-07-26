import { t } from "../i18n";
import { buildPqHost } from "./pq-host";
import { TRANSFORMS, strLit, nameList, quoteName, type TransformSpec, type TfField } from "./pq-transforms";
import { listWorkbookTables } from "../../adapters/xlsx/tables";
import { TRANSFORM_ICONS, APPEND_ICON, MERGE_ICON, LOAD_ICON, CANCEL_ICON, SAVE_ICON, NEWQUERY_ICON, QUERIES_ICON, STEPS_ICON, svgIcon } from "./pq-icons";
import type { Workbook } from "../model";
import type { MValue } from "mlang";

/** Set a button's content to an SVG icon followed by a text label. The label doubles as the
    tooltip so the button stays meaningful when the label is hidden (icon-only on narrow screens). */
function iconLabel(btn: HTMLButtonElement, inner: string, label: string): void {
  btn.textContent = "";
  btn.appendChild(svgIcon(inner));
  const span = document.createElement("span");
  span.textContent = label;
  btn.appendChild(span);
  if (!btn.title) btn.title = label;
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
  // Pane toggles: shown only on narrow screens, where the side panels become drawers.
  const queriesToggle = document.createElement("button");
  queriesToggle.className = "se-pqe-btn se-pqe-panetoggle";
  queriesToggle.title = t("pqQueries");
  queriesToggle.appendChild(svgIcon(QUERIES_ICON));
  const stepsToggle = document.createElement("button");
  stepsToggle.className = "se-pqe-btn se-pqe-panetoggle";
  stepsToggle.title = t("pqAppliedSteps");
  stepsToggle.appendChild(svgIcon(STEPS_ICON));

  const loadBtn = document.createElement("button");
  loadBtn.className = "se-pqe-btn";
  loadBtn.title = t("pqLoadTitle");
  iconLabel(loadBtn, LOAD_ICON, t("pqLoad"));
  loadBtn.hidden = !deps.loadQuery;
  const saveBtn = document.createElement("button");
  saveBtn.className = "se-pqe-btn primary";
  iconLabel(saveBtn, SAVE_ICON, t("pqSaveClose"));
  const cancelBtn = document.createElement("button");
  cancelBtn.className = "se-pqe-btn";
  iconLabel(cancelBtn, CANCEL_ICON, t("pqCancel"));
  bar.append(title, queriesToggle, stepsToggle, spacer, loadBtn, cancelBtn, saveBtn);

  // Drawer toggling (narrow screens): one side panel open at a time; tapping the preview closes.
  const closeDrawers = (): void => overlay.classList.remove("show-queries", "show-steps");
  const toggleDrawer = (cls: "show-queries" | "show-steps"): void => {
    const on = overlay.classList.contains(cls);
    closeDrawers();
    if (!on) overlay.classList.add(cls);
  };
  queriesToggle.addEventListener("click", () => toggleDrawer("show-queries"));
  stepsToggle.addEventListener("click", () => toggleDrawer("show-steps"));

  // Transform ribbon: each button appends a step to the query's final result. Buttons are laid
  // out in Excel-style groups (a 3-row column of buttons with the group name underneath).
  const ribbon = document.createElement("div");
  ribbon.className = "se-pqe-ribbon";
  const ribbonButtons: HTMLButtonElement[] = [];
  const makeRibbonBtn = (icon: string, label: string, fn: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.className = "se-pqe-rbtn";
    b.type = "button";
    iconLabel(b, icon, label);
    b.addEventListener("click", fn);
    ribbonButtons.push(b);
    return b;
  };
  const addRibbonGroup = (title: string, buttons: HTMLButtonElement[]): void => {
    const grp = document.createElement("div");
    grp.className = "se-pqe-grp";
    const body = document.createElement("div");
    body.className = "se-pqe-grp-body";
    buttons.forEach((b) => body.appendChild(b));
    const lbl = document.createElement("div");
    lbl.className = "se-pqe-grp-label";
    lbl.textContent = title;
    grp.append(body, lbl);
    ribbon.appendChild(grp);
  };
  {
    // Preserve the group order declared in TRANSFORMS.
    const groups: { title: string; specs: TransformSpec[] }[] = [];
    for (const spec of TRANSFORMS) {
      let g = groups.find((x) => x.title === spec.group);
      if (!g) { g = { title: spec.group, specs: [] }; groups.push(g); }
      g.specs.push(spec);
    }
    for (const g of groups) addRibbonGroup(g.title, g.specs.map((s) => makeRibbonBtn(TRANSFORM_ICONS[s.id] ?? "", s.label, () => void applyTransform(s))));
    // Combine group: cross-query operations (handled specially, not plain Table.* on one input).
    addRibbonGroup(t("pqCombine"), [
      makeRibbonBtn(APPEND_ICON, t("pqAppend"), () => void appendQueries()),
      makeRibbonBtn(MERGE_ICON, t("pqMerge"), () => void mergeQueries()),
    ]);
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
  // Tapping the preview dismisses an open drawer (narrow screens).
  center.addEventListener("click", () => { if (overlay.classList.contains("show-queries") || overlay.classList.contains("show-steps")) overlay.classList.remove("show-queries", "show-steps"); });

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
      item.addEventListener("click", () => { closeDrawers(); void selectQuery(name); });
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
      item.addEventListener("click", () => { closeDrawers(); void selectStep(step.name); });
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
