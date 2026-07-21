import { t } from "../i18n";
import { listWorkbookTables, tableForQuery, tableValue, type WorkbookTable } from "../../adapters/xlsx/tables";
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
  /** Apply a refreshed result to its destination table (undo/recalc/redraw are the host's). */
  apply(target: WorkbookTable, result: MTable): { rows: number };
}

export function setupQueryPanel(deps: QueryPanelDeps): { open(anchor: HTMLElement): void } {
  const { wrap, wb } = deps;
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
  pop.append(title, body, note);
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
      pre.hidden = !pre.hidden;
    });
    viewBtn.title = t("queryViewM");
    head.append(label, refreshBtn, viewBtn);
    const status = document.createElement("div");
    status.className = "sheetedit-qp-status";
    const pre = document.createElement("pre");
    pre.className = "sheetedit-qp-m";
    pre.textContent = sectionM;
    pre.hidden = true;
    el.append(head, status, pre);

    async function refresh(): Promise<void> {
      refreshBtn.disabled = true;
      status.textContent = t("queryRunning");
      try {
        const { evaluateSection } = await import("mlang");
        // Host: every workbook table, exposed the way Excel.CurrentWorkbook does.
        const tables = listWorkbookTables(wb);
        const host = {
          "Excel.CurrentWorkbook": {
            kind: "function" as const,
            name: "Excel.CurrentWorkbook",
            params: [],
            call: (): MValue => ({
              kind: "table",
              columns: ["Name", "Content"],
              rows: tables.map((tb) => [{ kind: "text", value: tb.displayName } as MValue, tableValue(wb, tb) as MValue]),
            }),
          } as MValue,
        };
        const section = await evaluateSection(sectionM, host);
        const result = section.run(name);
        if (result.kind !== "table") {
          status.textContent = t("queryNotTable", { kind: result.kind });
          return;
        }
        const target = tableForQuery(tables, name);
        if (!target) {
          status.textContent = t("queryNoTarget", { name });
          return;
        }
        const { rows } = deps.apply(target, result);
        status.textContent = t("queryRefreshed", { rows });
      } catch (e) {
        status.textContent = t("queryError", { msg: (e as Error).message });
      } finally {
        refreshBtn.disabled = false;
      }
    }
    return el;
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
