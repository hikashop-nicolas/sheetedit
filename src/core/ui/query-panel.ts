import { t } from "../i18n";
import { highlightM } from "./m-highlight";
import { buildPqHost } from "./pq-host";
import { listWorkbookTables, refreshOnLoadQueries, tableForQuery, type WorkbookTable } from "../../adapters/xlsx/tables";
import type { Workbook } from "../model";
import type { MValue } from "mlang";

// Power Query panel: lists the workbook's queries (from the MS-QDEFF DataMashup payload),
// shows the M source read-only, and refreshes a query on demand via mlang - entirely
// on-device. mlang is imported LAZILY on first open, so the editor's base bundle never
// carries the M engine. Queries whose source is external (Web/Sql/...) fail with the
// engine's precise error, which the row surfaces; the workbook is never modified by a
// failed refresh.

type MTable = Extract<MValue, { kind: "table" }>;

export interface QueryPanelDeps {
  wrap: HTMLElement; // positioned editor chrome (popover parent; toolbar clips)
  wb: Workbook;
  /** Files attached for File.Contents/Folder.Files, keyed by filename. Shared with the editor
      so files attached here are usable in query previews. */
  attachedFiles: Record<string, Uint8Array>;
  /** Apply a refreshed result to its destination table (undo/recalc/redraw are the host's).
      `silent` (on-open auto-refresh) means don't mark the workbook edited or add an undo step. */
  apply(target: WorkbookTable, result: MTable, opts?: { silent?: boolean }): { rows: number };
  /** Mark the workbook edited (so a save includes the rewritten query definitions). */
  markEdited?(): void;
}

export function setupQueryPanel(deps: QueryPanelDeps): { open(anchor: HTMLElement): void; runOnLoad(): Promise<void> } {
  const { wrap, wb, attachedFiles } = deps;
  const pop = document.createElement("div");
  pop.className = "sheetedit-qp-pop";
  pop.hidden = true;
  const title = document.createElement("div");
  title.className = "sheetedit-qp-title";
  title.textContent = t("queriesTitle");
  const body = document.createElement("div");
  body.className = "sheetedit-qp-body";
  const note = document.createElement("div");
  note.className = "sheetedit-qp-note";
  note.textContent = t("queriesNote");
  // Files the user attaches for File.Contents("name") to read (shared with the editor).
  const attach = document.createElement("label");
  attach.className = "sheetedit-qp-attach";
  attach.textContent = t("queryAttach");
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.multiple = true;
  fileInput.hidden = true;
  fileInput.addEventListener("change", async () => {
    for (const f of Array.from(fileInput.files ?? [])) attachedFiles[f.name] = new Uint8Array(await f.arrayBuffer());
    const names = Object.keys(attachedFiles);
    attach.textContent = names.length ? t("queryAttached", { names: names.join(", ") }) : t("queryAttach");
    fileInput.value = "";
  });
  attach.appendChild(fileInput);
  pop.append(title, body, attach, note);
  wrap.appendChild(pop);

  let anchor: HTMLElement | null = null;

  function position(): void {
    if (!anchor) return;
    const a = anchor.getBoundingClientRect();
    const w = wrap.getBoundingClientRect();
    pop.style.top = `${a.bottom - w.top + 4}px`;
    pop.style.left = `${Math.max(4, Math.min(a.right - w.left - pop.offsetWidth, w.width - pop.offsetWidth - 4))}px`;
  }

  async function load(): Promise<void> {
    body.textContent = t("queriesReading");
    try {
      const { readWorkbookQueries } = await import("mlang/qdeff");
      const q = readWorkbookQueries(wb.files);
      if (!q) {
        body.textContent = t("queriesNone");
        return;
      }
      const { evaluateSection } = await import("mlang");
      // Names come from the section without evaluating anything (members are lazy).
      const section = await evaluateSection(q.mashup.sectionM, {});
      body.textContent = "";
      for (const name of section.names) body.appendChild(row(name, q.mashup.sectionM));
      if (section.names.length === 0) body.textContent = t("queriesNone");
    } catch (e) {
      console.error("[sheetedit] query panel failed", e);
      body.textContent = t("queriesError", { msg: (e as Error).message });
    }
  }

  function row(name: string, sectionM: string): HTMLElement {
    const el = document.createElement("div");
    el.className = "sheetedit-qp-row";
    const head = document.createElement("div");
    head.className = "sheetedit-qp-rowhead";
    const label = document.createElement("span");
    label.className = "sheetedit-qp-name";
    label.textContent = name;
    const refreshBtn = btn(t("queryRefresh"), () => void refresh());
    const viewBtn = btn("M", () => {
      editorBox.hidden = !editorBox.hidden;
    });
    viewBtn.title = t("queryViewM");
    head.append(label, refreshBtn, viewBtn);
    const status = document.createElement("div");
    status.className = "sheetedit-qp-status";

    // Editable M: the whole Section1.m (all queries). Saving rewrites the DataMashup blob.
    // A transparent textarea sits over a syntax-highlighted layer that mirrors its content.
    const editorBox = document.createElement("div");
    editorBox.className = "sheetedit-qp-medit";
    editorBox.hidden = true;
    const mwrap = document.createElement("div");
    mwrap.className = "sheetedit-qp-mwrap";
    const highlight = document.createElement("pre");
    highlight.className = "sheetedit-qp-mhl";
    highlight.setAttribute("aria-hidden", "true");
    const code = document.createElement("code");
    highlight.appendChild(code);
    const area = document.createElement("textarea");
    area.className = "sheetedit-qp-m";
    area.value = sectionM;
    area.spellcheck = false;
    const paint = (): void => {
      code.innerHTML = highlightM(area.value);
    };
    area.addEventListener("input", paint);
    area.addEventListener("scroll", () => {
      highlight.scrollTop = area.scrollTop;
      highlight.scrollLeft = area.scrollLeft;
    });
    paint();
    mwrap.append(highlight, area);
    const saveBtn = btn(t("querySaveM"), () => void saveM());
    editorBox.append(mwrap, saveBtn);
    el.append(head, status, editorBox);

    async function saveM(): Promise<void> {
      saveBtn.disabled = true;
      status.textContent = t("querySaving");
      try {
        const newM = area.value;
        const { evaluateSection } = await import("mlang");
        await evaluateSection(newM, {}); // parse-check: reject invalid M before writing
        const { writeWorkbookSectionM } = await import("mlang/qdeff");
        wb.files = writeWorkbookSectionM(wb.files, newM);
        deps.markEdited?.();
        status.textContent = t("querySaved");
        await load(); // relist queries from the new section
      } catch (e) {
        status.textContent = t("querySaveError", { msg: (e as Error).message });
      } finally {
        saveBtn.disabled = false;
      }
    }

    async function refresh(): Promise<void> {
      refreshBtn.disabled = true;
      status.textContent = t("queryRunning");
      try {
        const st = await executeQuery(sectionM, name, false);
        if (st.kind === "ok") status.textContent = t("queryRefreshed", { rows: st.rows });
        else if (st.kind === "notTable") status.textContent = t("queryNotTable", { kind: st.detail });
        else status.textContent = t("queryNoTarget", { name });
      } catch (e) {
        const { isMissingConnector, missingConnectorName } = await import("mlang");
        if (isMissingConnector(e)) status.textContent = t("queryExternal", { connector: missingConnectorName(e) });
        else status.textContent = t("queryError", { msg: (e as Error).message });
      } finally {
        refreshBtn.disabled = false;
      }
    }
    return el;
  }

  /** Build the host bindings: the workbook tables plus the browser connectors (fetch a URL,
      read attached files). Identical for manual and on-open refresh. */
  const buildHost = (): Promise<Record<string, MValue>> => buildPqHost({ wb, attachedFiles });

  type ExecStatus = { kind: "ok"; rows: number } | { kind: "notTable"; detail: string } | { kind: "noTarget" };

  /** Evaluate one query and write its result to the destination table. `silent` (on-open
      auto-refresh) skips the dirty mark and undo entry. Throws on connector/eval errors. */
  async function executeQuery(sectionM: string, name: string, silent: boolean): Promise<ExecStatus> {
    const { evaluateSection } = await import("mlang");
    const host = await buildHost();
    const section = await evaluateSection(sectionM, host);
    const result = await section.run(name);
    if (result.kind !== "table") return { kind: "notTable", detail: result.kind };
    const target = tableForQuery(listWorkbookTables(wb), name);
    if (!target) return { kind: "noTarget" };
    const { rows } = deps.apply(target, result, { silent });
    return { kind: "ok", rows };
  }

  /** Honor "Refresh data when opening the file": run each flagged query (full connectors,
      exactly as a manual refresh) and write results in. A query that can't complete (e.g.
      File.Contents with nothing attached yet) is left as-is. */
  async function runOnLoad(): Promise<void> {
    const names = refreshOnLoadQueries(wb.files);
    if (names.length === 0) return;
    const { readWorkbookQueries } = await import("mlang/qdeff");
    const found = readWorkbookQueries(wb.files);
    if (!found) return;
    for (const name of names) {
      try {
        await executeQuery(found.mashup.sectionM, name, true);
      } catch {
        /* leave this query's table as saved; the user can refresh it manually */
      }
    }
  }

  document.addEventListener("mousedown", (e) => {
    if (!pop.isConnected) return;
    if (!pop.hidden && !pop.contains(e.target as Node) && e.target !== anchor) pop.hidden = true;
  });
  document.addEventListener("keydown", (e) => {
    if (pop.isConnected && e.key === "Escape") pop.hidden = true;
  });
  window.addEventListener("resize", () => {
    if (!pop.hidden) position();
  });

  return {
    open(at: HTMLElement): void {
      if (!pop.hidden) {
        pop.hidden = true;
        return;
      }
      anchor = at;
      pop.hidden = false;
      position();
      void load().then(position);
    },
    runOnLoad,
  };
}

function btn(label: string, fn: () => void): HTMLButtonElement {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "sheetedit-qp-btn";
  b.textContent = label;
  b.addEventListener("mousedown", (e) => e.preventDefault());
  b.addEventListener("click", (e) => {
    e.preventDefault();
    fn();
  });
  return b;
}
