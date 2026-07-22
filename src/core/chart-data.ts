import { cellDisplay, getCell, parseA1Ref, type Sheet, type Workbook } from "./model";
import type { ChartRef } from "./chart-model";

// Resolve a chart data ref against the live workbook (shared by the Chart.js overlay and the
// writers, which embed the resolved values as caches). Falls back to the ref's cached points.

export function resolveCells(wb: Workbook, ref: string): { row: number; col: number; sheet: Sheet }[] | null {
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(ref);
  const sheetName = m ? (m[1] ?? m[2]) : undefined;
  const body = (m ? m[3] : ref).replace(/\$/g, "");
  const sheet = sheetName ? wb.sheets.find((s) => s.name === sheetName) : wb.sheets[0];
  if (!sheet) return null;
  const [a, b] = body.split(":");
  const p1 = parseA1Ref(a ?? "");
  const p2 = b ? parseA1Ref(b) : p1;
  if (!p1 || !p2) return null;
  const out: { row: number; col: number; sheet: Sheet }[] = [];
  for (let r = Math.min(p1.row, p2.row); r <= Math.max(p1.row, p2.row); r++)
    for (let c = Math.min(p1.col, p2.col); c <= Math.max(p1.col, p2.col); c++) out.push({ row: r, col: c, sheet });
  return out;
}

export function resolveNumbers(wb: Workbook, ref: ChartRef | undefined): (number | null)[] {
  if (ref?.ref) {
    const cells = resolveCells(wb, ref.ref);
    if (cells) return cells.map(({ sheet, row, col }) => { const v = getCell(sheet, row, col)?.value ?? ""; const n = Number(v); return v !== "" && Number.isFinite(n) ? n : null; });
  }
  return (ref?.cache ?? []).map((v) => (typeof v === "number" ? v : v != null && v !== "" && Number.isFinite(Number(v)) ? Number(v) : null));
}

export function resolveLabels(wb: Workbook, ref: ChartRef | undefined): string[] {
  if (ref?.ref) {
    const cells = resolveCells(wb, ref.ref);
    if (cells) return cells.map(({ sheet, row, col }) => { const c = getCell(sheet, row, col); return c ? cellDisplay(c) : ""; });
  }
  return (ref?.cache ?? []).map((v) => (v == null ? "" : String(v)));
}

/** A series' display name (a literal string, or the first label of its name ref). */
export function seriesName(wb: Workbook, name: string | ChartRef | undefined): string | undefined {
  if (typeof name === "string") return name;
  if (!name) return undefined;
  return resolveLabels(wb, name)[0] || undefined;
}
