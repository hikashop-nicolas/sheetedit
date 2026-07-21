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
  /** Mark the workbook edited (so a save includes the rewritten query definitions). */
  markEdited?(): void;
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
  // Files the user attaches for File.Contents("name") to read, keyed by filename.
  const attachedFiles: Record<string, Uint8Array> = {};
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
    const editorBox = document.createElement("div");
    editorBox.className = "sheetedit-qp-medit";
    editorBox.hidden = true;
    const area = document.createElement("textarea");
    area.className = "sheetedit-qp-m";
    area.value = sectionM;
    area.spellcheck = false;
    const saveBtn = btn(t("querySaveM"), () => void saveM());
    editorBox.append(area, saveBtn);
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
        const mlang = await import("mlang");
        const { evaluateSection, asyncConnector, tableFromJson, MError } = mlang;
        const urlOf = (args: MValue[]): string => (args[0] as { value: string }).value;
        // Host: every workbook table, exposed the way Excel.CurrentWorkbook does.
        const tables = listWorkbookTables(wb);
        const host: Record<string, MValue> = {
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
          // Browser connectors: fetch a URL (CORS-permitting) and read attached files. Both
          // return an mlang binary that composes with Csv/Json/Xml/Excel.Workbook.
          "Web.Contents": asyncConnector("Web.Contents", async (args) => {
            const url = urlOf(args);
            const resp = await fetch(url);
            if (!resp.ok) throw new MError("DataSource.Error", `Web.Contents: ${resp.status} ${resp.statusText} for ${url}`);
            return { kind: "binary", bytes: new Uint8Array(await resp.arrayBuffer()) } as MValue;
          }) as MValue,
          // Web.Page: fetch a page as HTML text, for Html.Table to consume.
          "Web.Page": asyncConnector("Web.Page", async (args) => {
            const url = urlOf(args);
            const resp = await fetch(url);
            if (!resp.ok) throw new MError("DataSource.Error", `Web.Page: ${resp.status} ${resp.statusText} for ${url}`);
            return { kind: "text", value: await resp.text() } as MValue;
          }) as MValue,
          // OData.Feed: fetch the JSON feed and expand its `value` array to a table, following
          // @odata.nextLink for server-side paging.
          "OData.Feed": asyncConnector("OData.Feed", async (args) => {
            const records: unknown[] = [];
            let next: string | null = urlOf(args);
            for (let page = 0; next && page < 100; page++) {
              const resp: Response = await fetch(next, { headers: { Accept: "application/json" } });
              if (!resp.ok) throw new MError("DataSource.Error", `OData.Feed: ${resp.status} ${resp.statusText} for ${next}`);
              const body = (await resp.json()) as { value?: unknown[]; "@odata.nextLink"?: string };
              records.push(...(Array.isArray(body.value) ? body.value : [body]));
              next = body["@odata.nextLink"] ?? null;
            }
            return tableFromJson(records) as MValue;
          }) as MValue,
          "File.Contents": asyncConnector("File.Contents", async (args) => {
            const path = urlOf(args);
            const key = path.split(/[\\/]/).pop() ?? path;
            const file = attachedFiles[key] ?? attachedFiles[path];
            if (!file) throw new MError("DataSource.Error", `File.Contents: no attached file named '${key}' (attach it in the panel).`);
            return { kind: "binary", bytes: file } as MValue;
          }) as MValue,
          // Folder.Files over the attached-file set: [Content, Name, Extension].
          "Folder.Files": asyncConnector("Folder.Files", () =>
            Promise.resolve({
              kind: "table",
              columns: ["Content", "Name", "Extension"],
              rows: Object.entries(attachedFiles).map(([nm, bytes]) => [
                { kind: "binary", bytes } as MValue,
                { kind: "text", value: nm } as MValue,
                { kind: "text", value: nm.includes(".") ? `.${nm.split(".").pop()}` : "" } as MValue,
              ]),
            } as MValue)) as MValue,
        };
        const section = await evaluateSection(sectionM, host);
        const result = await section.run(name);
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
        const { isMissingConnector, missingConnectorName } = await import("mlang");
        if (isMissingConnector(e)) status.textContent = t("queryExternal", { connector: missingConnectorName(e) });
        else status.textContent = t("queryError", { msg: (e as Error).message });
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
