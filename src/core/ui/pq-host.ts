import { localeCode } from "../i18n";
import { listWorkbookTables, tableValue } from "../../adapters/xlsx/tables";
import type { Workbook } from "../model";
import type { MValue } from "mlang";

// The host bindings that back a Power Query refresh or preview: the workbook's own tables
// (Excel.CurrentWorkbook) plus the browser-reachable connectors (fetch a URL, read an attached
// file). Shared by the quick-refresh panel and the full editor so both resolve queries the
// same way. mlang stays deterministic; these bindings decide what a connector actually does.

export interface PqHostDeps {
  wb: Workbook;
  /** Files the user attached for File.Contents("name") / Folder.Files, keyed by filename. */
  attachedFiles: Record<string, Uint8Array>;
}

export async function buildPqHost(deps: PqHostDeps): Promise<Record<string, MValue>> {
  const { wb, attachedFiles } = deps;
  const { asyncConnector, tableFromJson, MError } = await import("mlang");
  const urlOf = (args: MValue[]): string => (args[0] as { value: string }).value;
  const tables = listWorkbookTables(wb);
  return {
    // Culture.Current reflects the editor's active language (overrides mlang's default).
    "Culture.Current": { kind: "text", value: localeCode() } as MValue,
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
    "Web.Contents": asyncConnector("Web.Contents", async (args) => {
      const url = urlOf(args);
      const resp = await fetch(url);
      if (!resp.ok) throw new MError("DataSource.Error", `Web.Contents: ${resp.status} ${resp.statusText} for ${url}`);
      return { kind: "binary", bytes: new Uint8Array(await resp.arrayBuffer()) } as MValue;
    }) as MValue,
    "Web.Page": asyncConnector("Web.Page", async (args) => {
      const url = urlOf(args);
      const resp = await fetch(url);
      if (!resp.ok) throw new MError("DataSource.Error", `Web.Page: ${resp.status} ${resp.statusText} for ${url}`);
      return { kind: "text", value: await resp.text() } as MValue;
    }) as MValue,
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
}
