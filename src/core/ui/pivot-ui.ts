import { colToLetters, ensureCell, getCell, type PivotTableInfo, type Sheet, type Workbook } from "../model";
import { computePivot, pivotColumnItems, pivotValueName, type PivotFunc, type PivotShowAs, type PivotSpec, type PivotValue } from "../pivot";
import { setCellInput } from "../workbook";
import { addSheet, newPivotId } from "../sheet-ops";
import { t } from "../i18n";
import { deleteXlsxPivotParts, setXlsxCellNumFmt, writeXlsxPivotParts } from "../../adapters/xlsx";
import { createXlsxSlicer, createXlsxTableSlicer } from "../../adapters/xlsx/slicer-create";
import { listWorkbookTables } from "../../adapters/xlsx/tables";
import { parseXmlOpt } from "../model";
import { formDialog } from "./form-dialog";
import { deleteOdsPivotDef, setOdsCellNumFmt, writeOdsPivotDef } from "../../adapters/ods";

// The pivot-table authoring UI: the insert/edit dialog (two panes, live preview), the create /
// refresh / edit-in-place / delete operations, and the overlay's action menu. Extracted from the
// editor's god-closure; it talks to the host through a small context object.

type Rect = { r1: number; c1: number; r2: number; c2: number };

export interface PivotUiCtx {
  wb: Workbook;
  /** The active sheet index (read fresh; it changes as sheets switch). */
  active: () => number;
  /** The container the modal / menu mount into. */
  wrap: HTMLElement;
  getSelRect: () => Rect;
  mark: () => void;
  renderGrid: () => void;
  switchSheet: (idx: number) => void;
  refreshPivotLayer: () => void;
  /** Re-render the slicer overlay after one is created. */
  refreshSlicers: () => void;
  /** Hide/show one row (a table slicer filters rows, not a pivot). */
  setRowHidden: (sheet: Sheet, row: number, hidden: boolean) => void;
  currentRegion: (sheet: Sheet, r: number, c: number) => Rect;
  chartsOn: boolean;
  chartInsert: (rect: Rect) => void;
}

/** ISO datetime of a pivot item's value: an Excel serial, or an already-date-like string. */
function itemIso(value: string | number, num: boolean): string | null {
  if (num || typeof value === "number") {
    const n = Number(value);
    if (!Number.isFinite(n)) return null;
    const d = new Date(Math.round((n - 25569) * 86400000));
    return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 19);
  }
  const s = String(value).trim();
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}T00:00:00`;
  const t = Date.parse(s);
  return Number.isNaN(t) ? null : new Date(t).toISOString().slice(0, 19);
}

export function setupPivotUi(ctx: PivotUiCtx): { openPivotMenu: (p: PivotTableInfo, x: number, y: number) => void; openPivotDialog: (opts?: PivotDialogOpts) => void; closeMenu: () => void; applySlicer: (s: import("../model").SheetSlicer) => void; applyTimeline: (t: import("../model").SheetTimeline) => void; addTableSlicer: () => void } {
  const { wb, wrap } = ctx;
  const activeSheet = (): Sheet => wb.sheets[ctx.active()]!;

  const clearRegion = (sheet: Sheet, rect?: Rect): void => {
    if (!rect) return;
    for (let r = rect.r1; r <= rect.r2; r++) for (let c = rect.c1; c <= rect.c2; c++) if (getCell(sheet, r, c)) setCellInput(sheet, r, c, "");
  };
  const placeMatrix = (sheet: Sheet, matrix: import("../pivot").PivotOutCell[][], anchor: { r: number; c: number }): void => {
    for (let r = 0; r < matrix.length; r++) for (let c = 0; c < matrix[r]!.length; c++) {
      const cell = matrix[r]![c]!;
      if (cell.value === "") continue;
      const rr = anchor.r + r, cc = anchor.c + c;
      setCellInput(sheet, rr, cc, String(cell.value));
      if (cell.numFmt) { const m = ensureCell(sheet, rr, cc); if (wb.kind === "ods") setOdsCellNumFmt(wb, sheet, m, cell.numFmt); else setXlsxCellNumFmt(wb, sheet, m, cell.numFmt); }
    }
  };
  // Emit a pivot (compute + place + write the format definition) onto dest at anchor, returning the
  // model to store on the sheet. Shared by create, edit and refresh.
  const buildPivotOn = (spec: PivotSpec, sourceSheetName: string, dest: Sheet, anchor: { r: number; c: number }): PivotTableInfo => {
    const srcSheet = wb.sheets.find((s) => s.name === sourceSheetName) ?? activeSheet();
    const computed = computePivot(srcSheet, spec);
    placeMatrix(dest, computed.matrix, anchor);
    let part: string | undefined, cachePart: string | undefined;
    if (wb.kind === "xlsx") ({ part, cachePart } = writeXlsxPivotParts(wb, dest, { row: anchor.r, col: anchor.c }, sourceSheetName, spec, computed));
    else { dest.odsDirty = true; writeOdsPivotDef(wb, dest.name, sourceSheetName, spec, computed); }
    return {
      cid: newPivotId(),
      name: "PivotTable",
      sourceSheet: sourceSheetName,
      sourceRange: { ...spec.source },
      targetRange: { r1: anchor.r, c1: anchor.c, r2: anchor.r + computed.height - 1, c2: anchor.c + computed.width - 1 },
      rowFields: spec.rows.map((c) => computed.fields[c]!.name),
      colFields: spec.cols.map((c) => computed.fields[c]!.name),
      pageFields: (spec.pages ?? []).map((p) => computed.fields[p.field]!.name),
      dataFields: spec.values.map((v) => ({ name: pivotValueName(v, (i) => computed.fields[i]!.name), func: v.func })),
      authorSpec: spec, part, cachePart, hostSheet: dest.name,
    };
  };
  // Create a pivot on a fresh sheet and switch to it.
  const createPivot = (spec: PivotSpec, sourceSheetName: string): void => {
    if (wb.kind !== "xlsx" && wb.kind !== "ods") return;
    const destIdx = addSheet(wb, "Pivot");
    const dest = wb.sheets[destIdx]!;
    dest.pivotTables = [buildPivotOn(spec, sourceSheetName, dest, { r: 1, c: 1 })];
    ctx.switchSheet(destIdx);
    ctx.mark();
  };
  // Recompute an authored pivot's output from its source (definition unchanged; the file's cache
  // carries refreshOnLoad so the desktop apps rebuild too).
  const refreshPivot = (host: Sheet, info: PivotTableInfo): void => {
    if (!info.authorSpec || !info.sourceSheet) return;
    const srcSheet = wb.sheets.find((s) => s.name === info.sourceSheet);
    if (!srcSheet) return;
    const computed = computePivot(srcSheet, info.authorSpec);
    const anchor = { r: info.targetRange?.r1 ?? 1, c: info.targetRange?.c1 ?? 1 };
    clearRegion(host, info.targetRange);
    placeMatrix(host, computed.matrix, anchor);
    if (wb.kind === "ods") host.odsDirty = true;
    info.targetRange = { r1: anchor.r, c1: anchor.c, r2: anchor.r + computed.height - 1, c2: anchor.c + computed.width - 1 };
    ctx.mark(); ctx.renderGrid(); ctx.refreshPivotLayer();
  };
  // Rewrite an authored pivot in place with a new spec: clear the old output, drop the old
  // definition parts, and emit the new pivot at the same anchor on the same sheet.
  const applyPivotEdit = (host: Sheet, info: PivotTableInfo, spec: PivotSpec): void => {
    const anchor = { r: info.targetRange?.r1 ?? 1, c: info.targetRange?.c1 ?? 1 };
    clearRegion(host, info.targetRange);
    if (wb.kind === "xlsx" && info.part && info.cachePart) deleteXlsxPivotParts(wb, host, info.part, info.cachePart);
    else if (wb.kind === "ods") deleteOdsPivotDef(wb, host.name);
    const fresh = buildPivotOn(spec, info.sourceSheet ?? host.name, host, anchor);
    const idx = host.pivotTables?.indexOf(info) ?? -1;
    if (host.pivotTables && idx >= 0) host.pivotTables[idx] = fresh; else (host.pivotTables ??= []).push(fresh);
    ctx.mark(); ctx.renderGrid(); ctx.refreshPivotLayer();
  };
  // Reopen the insert dialog for an authored pivot, prefilled from its spec; apply rewrites in place.
  const editPivot = (host: Sheet, info: PivotTableInfo): void => {
    const spec = info.authorSpec;
    if (!spec || !info.sourceSheet) return;
    const srcSheet = wb.sheets.find((s) => s.name === info.sourceSheet);
    if (!srcSheet) return;
    const width = spec.source.c2 - spec.source.c1 + 1;
    const roles = Array.from({ length: width }, () => "unused");
    const funcs = Array.from({ length: width }, () => "sum");
    const showAs = Array.from({ length: width }, () => "normal");
    const pageItems: (number | null)[] = Array.from({ length: width }, () => null);
    for (const i of spec.rows) roles[i] = "rows";
    for (const i of spec.cols) roles[i] = "columns";
    for (const v of spec.values) if (v.field != null && v.calc == null) { roles[v.field] = "values"; funcs[v.field] = v.func ?? "sum"; showAs[v.field] = v.showAs ?? "normal"; }
    for (const p of spec.pages ?? []) { roles[p.field] = "page"; pageItems[p.field] = p.item; }
    const calcFields = spec.values.filter((v) => v.calc != null).map((v) => ({ name: v.name ?? "Calc", formula: v.calc! }));
    const calcItems = (spec.calcItems ?? []).map((c) => ({ ...c }));
    openPivotDialog({ sheet: srcSheet, range: { ...spec.source }, initial: { roles, funcs, showAs, pageItems, subtotals: !!spec.subtotals, calcFields, calcItems }, onApply: (ns) => applyPivotEdit(host, info, ns) });
  };
  // A small action menu shown when a pivot's overlay tag is clicked: refresh or edit an authored
  // pivot (pivots read from a file are read-only in place; open in Excel/LibreOffice to change them).
  let pivotMenu: HTMLElement | null = null;
  const closePivotMenu = (): void => { pivotMenu?.remove(); pivotMenu = null; };
  const openPivotMenu = (pivot: PivotTableInfo, x: number, y: number): void => {
    closePivotMenu();
    const host = activeSheet();
    const menu = document.createElement("div");
    menu.className = "sheetedit-pivot-menu";
    menu.className = "sheetedit-pivot-menu";
    // Position is the only thing that cannot be a class: it follows the click.
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    const item = (label: string, onClick: (() => void) | null): HTMLElement => {
      const b = document.createElement("button");
      b.textContent = label;
      if (!onClick) b.classList.add("is-inert");
      if (onClick) b.addEventListener("click", () => { closePivotMenu(); onClick(); }); // hover is a CSS rule
      return b;
    };
    const items: HTMLElement[] = [];
    if (pivot.authorSpec) items.push(item(t("pivotRefresh"), () => refreshPivot(host, pivot)), item(t("pivotEditAction"), () => editPivot(host, pivot)));
    // A slicer can be added for any field the pivot groups by (those carry sharedItems in the cache).
    if (wb.kind === "xlsx" && pivot.authorSpec && pivot.sourceSheet) items.push(item(t("slicerInsert"), () => addSlicer(host, pivot)));
    // A chart over the pivot's output (updates as the pivot recomputes). Exclude the grand total row
    // and, when there are column fields, the grand total column, so the totals are not charted.
    if (ctx.chartsOn && pivot.targetRange) items.push(item(t("pivotChart"), () => {
      const tr = pivot.targetRange!;
      const gc = pivot.colFields.length > 0 ? 1 : 0;
      ctx.chartInsert({ r1: tr.r1, c1: tr.c1, r2: Math.max(tr.r1, tr.r2 - 1), c2: Math.max(tr.c1, tr.c2 - gc) });
    }));
    if (!items.length) items.push(item(t("pivotReadOnly"), null));
    menu.append(...items);
    document.body.appendChild(menu);
    pivotMenu = menu;
    setTimeout(() => document.addEventListener("pointerdown", function h(e) { if (!menu.contains(e.target as Node)) { closePivotMenu(); document.removeEventListener("pointerdown", h, true); } }, true), 0);
  };

  // Read the dialog state (per-column role, function, "show as", page selection, subtotals, and any
  // calculated fields) into a spec.
  const pivotSpecFrom = (range: Rect, roles: string[], funcs: string[], showAs: string[], pageItems: (number | null)[], subtotals: boolean, calcFields: { name: string; formula: string }[], calcItems: { field: number; name: string; formula: string }[]): PivotSpec => {
    const rows: number[] = [], cols: number[] = [], values: PivotValue[] = [], pages: { field: number; item: number | null }[] = [];
    for (let i = 0; i < roles.length; i++) {
      if (roles[i] === "rows") rows.push(i);
      else if (roles[i] === "columns") cols.push(i);
      else if (roles[i] === "values") values.push({ field: i, func: funcs[i] as PivotFunc, showAs: showAs[i] === "normal" ? undefined : (showAs[i] as PivotShowAs) });
      else if (roles[i] === "page") pages.push({ field: i, item: pageItems[i] ?? null });
    }
    for (const cf of calcFields) if (cf.formula.trim() && cf.name.trim()) values.push({ calc: cf.formula.trim(), name: cf.name.trim() });
    const ci = calcItems.filter((c) => c.formula.trim() && c.name.trim() && (rows.includes(c.field) || cols.includes(c.field)));
    return { source: range, rows, cols, values, pages: pages.length ? pages : undefined, subtotals: subtotals || undefined, calcItems: ci.length ? ci.map((c) => ({ field: c.field, name: c.name.trim(), formula: c.formula.trim() })) : undefined };
  };

  // Insert-pivot dialog: a two-pane modal. Left: assign each source column (from the selection's
  // header row) to Rows / Columns / Values (with a function). Right: a live preview of the resulting
  // pivot that updates as you change roles. Needs at least one Rows field and one Values field.
  const openPivotDialog = (opts?: PivotDialogOpts): void => {
    if (wb.kind !== "xlsx" && wb.kind !== "ods") return;
    const sheet = opts?.sheet ?? activeSheet();
    const s = ctx.getSelRect();
    const range = opts?.range ?? (s.r2 > s.r1 || s.c2 > s.c1 ? s : ctx.currentRegion(sheet, s.r1, s.c1));
    const width = range.c2 - range.c1 + 1;
    const headers: string[] = [];
    for (let c = 0; c < width; c++) { const cell = getCell(sheet, range.r1, range.c1 + c); headers.push((cell?.value ?? "").trim() || `Column ${c + 1}`); }
    const headerBlank: boolean[] = [];
    for (let c = 0; c < width; c++) headerBlank.push((getCell(sheet, range.r1, range.c1 + c)?.value ?? "").trim() === "");
    // A usable data region needs a header row plus at least one data row and a real header somewhere.
    const hasData = !!opts || (range.r2 > range.r1 && headerBlank.some((b) => !b));
    const roles: string[] = opts?.initial.roles.slice() ?? headers.map((_, i) => (hasData ? (i === 0 ? "rows" : i === width - 1 ? "values" : "unused") : "unused"));
    const funcs: string[] = opts?.initial.funcs.slice() ?? headers.map(() => "sum");
    const showAs: string[] = opts?.initial.showAs.slice() ?? headers.map(() => "normal");
    const pageItems: (number | null)[] = opts?.initial.pageItems.slice() ?? headers.map(() => null);
    let subtotals = opts?.initial.subtotals ?? false;
    const calcFields: { name: string; formula: string }[] = opts?.initial.calcFields.map((c) => ({ ...c })) ?? [];
    const calcItems: { field: number; name: string; formula: string }[] = opts?.initial.calcItems.map((c) => ({ ...c })) ?? [];
    const onApply = opts?.onApply ?? ((spec: PivotSpec) => createPivot(spec, sheet.name));
    // Distinct values per column, for the page-filter pickers (aligned to the engine's item order).
    const colItems = headers.map((_, i) => (hasData ? pivotColumnItems(sheet, range, i) : []));

    const modal = document.createElement("div");
    modal.className = "sheetedit-form-modal sheetedit-pivot-modal";
    modal.classList.add("sheetedit-modal");
    const card = document.createElement("div");
    card.classList.add("sheetedit-card", "is-wide");
    const h = document.createElement("h3"); h.textContent = t("pivotInsert"); card.appendChild(h);
    const srcLine = document.createElement("div");
    srcLine.textContent = `${t("pivotSource")}: ${sheet.name}!${colToLetters(range.c1)}${range.r1}:${colToLetters(range.c2)}${range.r2}`;
    srcLine.className = "sheetedit-note sheetedit-subtle"; card.appendChild(srcLine);
    const body = document.createElement("div"); body.className = "sheetedit-pivot-body"; card.appendChild(body);
    const left = document.createElement("div"); left.className = "sheetedit-pivot-left"; body.appendChild(left);
    const right = document.createElement("div"); right.className = "sheetedit-pivot-right"; body.appendChild(right);
    const previewLbl = document.createElement("div"); previewLbl.textContent = t("pivotPreview"); previewLbl.className = "sheetedit-subtle"; right.appendChild(previewLbl);
    const preview = document.createElement("div"); preview.className = "sheetedit-pivot-preview"; right.appendChild(preview);

    const roleOpts: [string, string][] = [["unused", t("pivotUnused")], ["rows", t("pivotRows")], ["columns", t("pivotColumns")], ["values", t("pivotValues")], ["page", t("pivotPageF")]];
    const funcOpts: [string, string][] = (["sum", "count", "average", "min", "max"] as const).map((f) => [f, t(`pivotFn_${f}`)]);
    const showAsOpts: [string, string][] = (["normal", "percentOfTotal", "percentOfCol", "percentOfRow", "runningTotal"] as const).map((f) => [f, t(`pivotShow_${f}`)]);
    const mkSelect = (opts: [string, string][], value: string, onChange: (v: string) => void): HTMLSelectElement => {
      const sel = document.createElement("select"); sel.className = "sheetedit-input";
      for (const [v, l] of opts) { const o = document.createElement("option"); o.value = v; o.textContent = l; sel.appendChild(o); }
      sel.value = value; sel.addEventListener("change", () => onChange(sel.value));
      return sel;
    };

    let insertBtn: HTMLButtonElement;
    const renderPreview = (): void => {
      const spec = pivotSpecFrom(range, roles, funcs, showAs, pageItems, subtotals, calcFields, calcItems);
      const valid = hasData && spec.rows.length > 0 && spec.values.length > 0;
      if (insertBtn) insertBtn.disabled = !valid;
      preview.textContent = "";
      // A message in place of the preview grid reads as muted text, not as data.
      if (!hasData) { preview.textContent = t("pivotNoData"); preview.classList.add("is-message"); return; }
      if (!valid) { preview.textContent = t("pivotHint"); preview.classList.add("is-message"); return; }
      preview.classList.remove("is-message");
      const computed = computePivot(sheet, spec);
      const table = document.createElement("table");
      const maxR = Math.min(computed.matrix.length, 12), maxC = Math.min(computed.width, 8);
      for (let r = 0; r < maxR; r++) {
        const tr = document.createElement("tr");
        for (let c = 0; c < maxC; c++) {
          const cell = computed.matrix[r]![c]!;
          const td = document.createElement("td");
          td.textContent = cell.value === "" ? "" : cell.numFmt === "0.00%" && cell.kind === "n" ? `${(Number(cell.value) * 100).toFixed(2)}%` : String(cell.value);
          if (cell.kind === "n") td.classList.add("num");
          if (cell.bold) td.classList.add("bold");
          tr.appendChild(td);
        }
        if (computed.width > maxC) { const td = document.createElement("td"); td.textContent = "…"; td.className = "ellipsis"; tr.appendChild(td); }
        table.appendChild(tr);
      }
      if (computed.matrix.length > maxR) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.textContent = "…"; td.colSpan = maxC; td.className = "ellipsis"; tr.appendChild(td); table.appendChild(tr); }
      preview.appendChild(table);
    };

    if (!hasData) { const p = document.createElement("div"); p.textContent = t("pivotNoData"); p.className = "sheetedit-subtle"; left.appendChild(p); }
    for (let i = 0; i < width; i++) {
      const row = document.createElement("div"); row.className = "sheetedit-pivot-row";
      const name = document.createElement("span"); name.textContent = headers[i]!; name.title = headers[i]!;
      name.className = "sheetedit-pivot-name"; row.appendChild(name);
      const funcSel = mkSelect(funcOpts, funcs[i]!, (v) => { funcs[i] = v; renderPreview(); });
      funcSel.dataset.field = `func_${i}`;
      funcSel.classList.toggle("sheetedit-hidden", roles[i] !== "values");
      // "Show values as" picker for a value field (% of total / running total).
      const showSel = mkSelect(showAsOpts, showAs[i]!, (v) => { showAs[i] = v; renderPreview(); });
      showSel.dataset.field = `show_${i}`;
      showSel.classList.toggle("sheetedit-hidden", roles[i] !== "values");
      // Page-filter value picker (All + each distinct value of this column).
      const pageOpts: [string, string][] = [["", t("pivotAll")], ...colItems[i]!.map((it, k): [string, string] => [String(k), it.label])];
      const pageSel = mkSelect(pageOpts, "", (v) => { pageItems[i] = v === "" ? null : Number(v); renderPreview(); });
      pageSel.dataset.field = `page_${i}`;
      pageSel.classList.toggle("sheetedit-hidden", roles[i] !== "page");
      const roleSel = mkSelect(roleOpts, roles[i]!, (v) => { roles[i] = v; const val = v === "values"; funcSel.classList.toggle("sheetedit-hidden", !val); showSel.classList.toggle("sheetedit-hidden", !val); pageSel.classList.toggle("sheetedit-hidden", v !== "page"); renderPreview(); });
      roleSel.dataset.field = `role_${i}`;
      row.append(roleSel, funcSel, showSel, pageSel);
      if (hasData) left.appendChild(row);
    }
    // Calculated fields: a name + a formula over field names (e.g. "Revenue - Cost"); each becomes
    // an extra value field. Rebuilt live in the preview; rows can be removed.
    const calcWrap = document.createElement("div"); calcWrap.className = "sheetedit-pivot-group";
    const renderCalc = (): void => {
      calcWrap.textContent = "";
      calcFields.forEach((cf, k) => {
        const row = document.createElement("div"); row.className = "sheetedit-pivot-calc";
        const nameI = document.createElement("input"); nameI.type = "text"; nameI.placeholder = t("pivotCalcName"); nameI.value = cf.name; nameI.className = "sheetedit-input name"; nameI.dataset.field = `calcname_${k}`;
        nameI.addEventListener("input", () => { cf.name = nameI.value; renderPreview(); });
        const fI = document.createElement("input"); fI.type = "text"; fI.placeholder = t("pivotCalcFormula"); fI.value = cf.formula; fI.className = "sheetedit-input formula"; fI.dataset.field = `calcformula_${k}`;
        fI.addEventListener("input", () => { cf.formula = fI.value; renderPreview(); });
        const rm = document.createElement("button"); rm.type = "button"; rm.textContent = "✕"; rm.title = t("chartDelete"); rm.className = "sheetedit-iconbtn";
        rm.addEventListener("click", () => { calcFields.splice(k, 1); renderCalc(); renderPreview(); });
        row.append(nameI, fI, rm); calcWrap.appendChild(row);
      });
      const add = document.createElement("button"); add.type = "button"; add.dataset.role = "add-calc"; add.textContent = t("pivotCalcAdd"); add.className = "sheetedit-addbtn";
      add.addEventListener("click", () => { calcFields.push({ name: `Calc${calcFields.length + 1}`, formula: "" }); renderCalc(); renderPreview(); });
      calcWrap.appendChild(add);
    };
    if (hasData) { renderCalc(); left.appendChild(calcWrap); }
    // Calculated items: a synthetic member of a row/column field, from a formula over that field's
    // item names (e.g. on Region: "North + South"). Shown as an extra row/column.
    const ciWrap = document.createElement("div"); ciWrap.className = "sheetedit-pivot-group";
    const renderCalcItems = (): void => {
      ciWrap.textContent = "";
      const fieldOpts: [string, string][] = headers.map((h, i): [string, string] => [String(i), h]);
      calcItems.forEach((ci, k) => {
        const row = document.createElement("div"); row.className = "sheetedit-pivot-calcitem";
        const fieldSel = mkSelect(fieldOpts, String(ci.field), (v) => { ci.field = Number(v); renderPreview(); }); fieldSel.dataset.field = `citemfield_${k}`;
        const nameI = document.createElement("input"); nameI.type = "text"; nameI.placeholder = t("pivotCalcName"); nameI.value = ci.name; nameI.className = "sheetedit-input name"; nameI.dataset.field = `citemname_${k}`;
        nameI.addEventListener("input", () => { ci.name = nameI.value; renderPreview(); });
        const fI = document.createElement("input"); fI.type = "text"; fI.placeholder = t("pivotItemFormula"); fI.value = ci.formula; fI.className = "sheetedit-input formula"; fI.dataset.field = `citemformula_${k}`;
        fI.addEventListener("input", () => { ci.formula = fI.value; renderPreview(); });
        const rm = document.createElement("button"); rm.type = "button"; rm.textContent = "✕"; rm.className = "sheetedit-iconbtn";
        rm.addEventListener("click", () => { calcItems.splice(k, 1); renderCalcItems(); renderPreview(); });
        row.append(fieldSel, nameI, fI, rm); ciWrap.appendChild(row);
      });
      const add = document.createElement("button"); add.type = "button"; add.dataset.role = "add-calcitem"; add.textContent = t("pivotItemAdd"); add.className = "sheetedit-addbtn";
      add.addEventListener("click", () => { const f = [...roles.keys()].find((i) => roles[i] === "rows" || roles[i] === "columns") ?? 0; calcItems.push({ field: f, name: `Item${calcItems.length + 1}`, formula: "" }); renderCalcItems(); renderPreview(); });
      ciWrap.appendChild(add);
    };
    if (hasData) { renderCalcItems(); left.appendChild(ciWrap); }
    // Subtotals toggle (meaningful once there are nested row/column fields).
    if (hasData) {
      const stRow = document.createElement("label"); stRow.className = "sheetedit-checkline";
      const cb = document.createElement("input"); cb.type = "checkbox"; cb.dataset.field = "subtotals"; cb.checked = subtotals;
      cb.addEventListener("change", () => { subtotals = cb.checked; renderPreview(); });
      const sp = document.createElement("span"); sp.textContent = t("pivotSubtotals");
      stRow.append(cb, sp); left.appendChild(stRow);
    }

    const actions = document.createElement("div"); actions.className = "sheetedit-actions is-plain";
    const close = (): void => modal.remove();
    const cancel = document.createElement("button"); cancel.textContent = t("chartCancel"); cancel.dataset.role = "cancel";
    cancel.className = "sheetedit-dlg-btn";
    insertBtn = document.createElement("button"); insertBtn.textContent = t("pivotCreate"); insertBtn.dataset.role = "ok";
    insertBtn.className = "sheetedit-dlg-btn is-primary";
    cancel.addEventListener("click", close);
    insertBtn.addEventListener("click", () => { const spec = pivotSpecFrom(range, roles, funcs, showAs, pageItems, subtotals, calcFields, calcItems); if (!spec.rows.length || !spec.values.length) return; close(); onApply(spec); });
    actions.append(cancel, insertBtn); card.appendChild(actions);
    modal.appendChild(card); wrap.appendChild(modal);
    modal.addEventListener("mousedown", (e) => { if (e.target === modal) close(); });
    renderPreview();
  };

  /** Add a slicer for one of the pivot's grouping fields, placed to the right of its output. */
  const addSlicer = (host: Sheet, info: PivotTableInfo): void => {
    const spec = info.authorSpec;
    const src = wb.sheets.find((s) => s.name === info.sourceSheet);
    if (!spec || !src) return;
    // Only fields the pivot groups by are in the cache with sharedItems, which a slicer indexes.
    const fields = [...spec.rows, ...spec.cols, ...(spec.pages ?? []).map((p) => p.field)];
    const opts = fields.map((f) => ({ value: String(f), label: (getCell(src, spec.source.r1, spec.source.c1 + f)?.value ?? `Column ${f + 1}`).trim() || `Column ${f + 1}` }));
    if (!opts.length) return;
    formDialog(ctx.wrap, t("slicerInsert"), [{ key: "field", label: t("slicerField"), type: "select", value: opts[0]!.value, options: opts }], (v) => {
      const field = Number(v.field);
      const label = opts.find((o) => o.value === String(field))?.label ?? "";
      const items = pivotColumnItems(src, spec.source, field).map((i) => i.label);
      // Park it just right of the pivot output, roughly 3 columns by 8 rows.
      const tr = info.targetRange;
      const c1 = (tr?.c2 ?? 1) + 2, r1 = tr?.r1 ?? 1;
      const anchor = { fromCol: c1, fromRow: r1, fromColOff: 0, fromRowOff: 0, toCol: c1 + 3, toRow: r1 + 8, toColOff: 0, toRowOff: 0 };
      const sl = createXlsxSlicer(wb, host, info, label, items, anchor);
      if (sl) { ctx.mark(); ctx.renderGrid(); ctx.refreshSlicers(); }
    });
  };

  /**
   * Apply a slicer's selection to every pivot it drives and recompute them.
   * The slicer's item indices are cache-order, which need not match the engine's item order, so the
   * selection is mapped by LABEL onto the engine's items.
   */
  const applySlicer = (sl: import("../model").SheetSlicer): void => {
    const wanted = new Set(sl.items.filter((i) => i.selected).map((i) => i.label));
    const allOn = sl.items.every((i) => i.selected);
    // A table slicer filters the table's ROWS rather than a pivot: hide the ones it excludes.
    if (sl.kind === "table" && sl.table) {
      const t = sl.table;
      const sheet = wb.sheets[t.sheetIndex];
      if (sheet) {
        const hidden = new Set<number>(sheet.filterHidden ?? []);
        for (let r = t.r1 + t.headerRows; r <= t.r2; r++) {
          const cell = getCell(sheet, r, t.c1 + t.col);
          const label = cell ? (cell.display ?? cell.value) : "";
          const show = allOn || wanted.has(label);
          if (show) hidden.delete(r); else hidden.add(r);
          ctx.setRowHidden(sheet, r, !show);
        }
        sheet.filterHidden = hidden.size ? hidden : undefined;
        ctx.mark(); ctx.renderGrid();
      }
      return;
    }
    if (sl.kind === "olap") return; // no modelled OLAP source to filter
    for (const host of wb.sheets) {
      for (const info of host.pivotTables ?? []) {
        // An empty pivotTables list on the cache means "every pivot on this cache".
        if (sl.pivotTables.length && !sl.pivotTables.includes(info.name)) continue;
        const spec = info.authorSpec;
        if (!spec || !info.sourceSheet) continue;
        const src = wb.sheets.find((s) => s.name === info.sourceSheet);
        if (!src) continue;
        // Which source column is this slicer's field?
        const width = spec.source.c2 - spec.source.c1 + 1;
        let field = -1;
        for (let c = 0; c < width; c++) {
          const head = (getCell(src, spec.source.r1, spec.source.c1 + c)?.value ?? "").trim();
          if (head === sl.sourceName) { field = c; break; }
        }
        if (field < 0) continue;
        const others = (spec.itemFilters ?? []).filter((f) => f.field !== field);
        if (allOn) spec.itemFilters = others.length ? others : undefined;
        else {
          // Compute once to learn the engine's item order for this field, then map labels to indices.
          const probe = computePivot(src, { ...spec, itemFilters: others });
          const idx: number[] = [];
          probe.fields[field]!.items.forEach((it, i) => { if (wanted.has(it.label)) idx.push(i); });
          spec.itemFilters = [...others, { field, items: idx }];
        }
        refreshPivot(host, info);
      }
    }
  };

  /**
   * Apply a timeline's date range to every pivot it drives. The range is expressed as item filters
   * on the timeline's date field: items whose date falls outside [startDate, endDate) are dropped.
   * An empty range means "no filter" and removes the timeline's filter from the pivot.
   */
  const applyTimeline = (tl: import("../model").SheetTimeline): void => {
    for (const host of wb.sheets) {
      for (const info of host.pivotTables ?? []) {
        if (tl.pivotTables.length && !tl.pivotTables.includes(info.name)) continue;
        const spec = info.authorSpec;
        if (!spec || !info.sourceSheet) continue;
        const src = wb.sheets.find((s) => s.name === info.sourceSheet);
        if (!src) continue;
        const width = spec.source.c2 - spec.source.c1 + 1;
        let field = -1;
        for (let c = 0; c < width; c++) {
          const head = (getCell(src, spec.source.r1, spec.source.c1 + c)?.value ?? "").trim();
          if (head === tl.sourceName) { field = c; break; }
        }
        if (field < 0) continue;
        const others = (spec.itemFilters ?? []).filter((f) => f.field !== field);
        if (!tl.startDate || !tl.endDate) spec.itemFilters = others.length ? others : undefined;
        else {
          const probe = computePivot(src, { ...spec, itemFilters: others });
          const idx: number[] = [];
          probe.fields[field]!.items.forEach((it, i) => {
            const iso = itemIso(it.value, it.num);
            if (iso && iso >= tl.startDate! && iso < tl.endDate!) idx.push(i);
          });
          spec.itemFilters = [...others, { field, items: idx }];
        }
        refreshPivot(host, info);
      }
    }
  };

  /**
   * Add a slicer for one column of the Excel table under the cursor.
   * A table slicer binds by the tableColumn's @id, so the column offset the UI filters on is mapped
   * back to that id through the table part.
   */
  const addTableSlicer = (): void => {
    if (wb.kind !== "xlsx") return;
    const sheet = activeSheet();
    const si = wb.sheets.indexOf(sheet);
    const sel = ctx.getSelRect();
    const table = listWorkbookTables(wb).find((tb) => tb.sheetIndex === si && sel.r1 >= tb.r1 && sel.r1 <= tb.r2 && sel.c1 >= tb.c1 && sel.c1 <= tb.c2);
    if (!table) { formDialog(ctx.wrap, t("slicerInsert"), [{ key: "none", label: t("slicerNoTable"), type: "note" }], () => undefined); return; }
    const doc = wb.files[table.path] ? parseXmlOpt(wb.files[table.path]!) : undefined;
    const tableId = Number(doc?.documentElement.getAttribute("id") ?? "0");
    const colEls = doc ? Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "tableColumn") : [];
    if (!tableId || !colEls.length) return;
    const opts = colEls.map((e, i) => ({ value: String(i), label: e.getAttribute("name") ?? `Column ${i + 1}` }));
    formDialog(ctx.wrap, t("slicerInsert"), [{ key: "field", label: t("slicerField"), type: "select", value: String(Math.max(0, sel.c1 - table.c1)), options: opts }], (v) => {
      const col = Number(v.field);
      const columnId = Number(colEls[col]?.getAttribute("id") ?? String(col + 1));
      const label = opts[col]?.label ?? "";
      // The slicer lists the column's distinct values, in first-seen order.
      const seen: string[] = [];
      for (let r = table.r1 + table.headerRows; r <= table.r2; r++) {
        const val = (getCell(sheet, r, table.c1 + col)?.display ?? getCell(sheet, r, table.c1 + col)?.value ?? "").toString();
        if (val !== "" && !seen.includes(val)) seen.push(val);
      }
      if (!seen.length) return;
      // Park it right of the table, stepping right for each slicer already on the sheet.
      const c1 = table.c2 + 2 + (sheet.slicers?.length ?? 0) * 4, r1 = table.r1;
      const anchor = { fromCol: c1, fromRow: r1, fromColOff: 0, fromRowOff: 0, toCol: c1 + 3, toRow: r1 + 8, toColOff: 0, toRowOff: 0 };
      const sl = createXlsxTableSlicer(wb, sheet, table, tableId, columnId, col, label, seen, anchor);
      if (sl) { ctx.mark(); ctx.renderGrid(); ctx.refreshSlicers(); }
    });
  };

  return { openPivotMenu, openPivotDialog, closeMenu: closePivotMenu, applySlicer, applyTimeline, addTableSlicer };
}

export interface PivotDialogOpts {
  sheet: Sheet;
  range: Rect;
  initial: { roles: string[]; funcs: string[]; showAs: string[]; pageItems: (number | null)[]; subtotals: boolean; calcFields: { name: string; formula: string }[]; calcItems: { field: number; name: string; formula: string }[] };
  onApply: (spec: PivotSpec) => void;
}
