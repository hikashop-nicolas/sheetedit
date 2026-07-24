import { parseXmlOpt, serializeXml, type Sheet, type Workbook } from "../../core/model";
import { colToLetters } from "../../core/model";
import { ODS } from "./shared";

// Author LibreOffice sparklines on an ods sheet. The model (sheet.sparklines) is the source of
// truth and drives the live render; writeOdsSparklines regenerates the table's
// <calcext:sparkline-groups> (grouped by identical style, appended after the rows) from it on save.
// Only sheets whose sparklines were edited are rebuilt, so untouched groups keep their original
// (richer) attributes verbatim.

type SparkStyle = { type: "line" | "column" | "stacked"; color: string; negColor?: string };
type SparkSpec = SparkStyle & { dataRef: string };
type SparkItem = { host: { r: number; c: number }; dataRef: string };
const CALC = ODS.calcext;

/** Add or remove the sparkline hosted at `host`. */
export function setOdsSparkline(sheet: Sheet, host: { r: number; c: number }, spec: SparkSpec | null): void {
  sheet.sparklines = (sheet.sparklines ?? []).filter((s) => !(s.host.r === host.r && s.host.c === host.c));
  if (spec) sheet.sparklines.push({ type: spec.type, color: spec.color, negColor: spec.negColor, host: { ...host }, dataRef: spec.dataRef });
  if (!sheet.sparklines.length) sheet.sparklines = undefined;
  sheet.sparklinesDirty = true;
  sheet.odsDirty = true;
}

/** Add a group of sparklines sharing one style (replacing any at the same host cells). */
export function setOdsSparklineGroup(sheet: Sheet, style: SparkStyle, items: SparkItem[]): void {
  if (!items.length) return;
  const hosts = new Set(items.map((it) => `${it.host.r}:${it.host.c}`));
  sheet.sparklines = (sheet.sparklines ?? []).filter((s) => !hosts.has(`${s.host.r}:${s.host.c}`));
  for (const it of items) sheet.sparklines.push({ type: style.type, color: style.color, negColor: style.negColor, host: { ...it.host }, dataRef: it.dataRef });
  sheet.sparklinesDirty = true;
  sheet.odsDirty = true;
}

/** A1 (optionally sheet-qualified) -> an ODF sparkline address ("Sheet.A1" / "Sheet.A1:Sheet.E1"),
    quoting the sheet name when it is not a bare identifier. */
function a1ToOdfAddr(a1: string, defaultSheet: string): string {
  const bang = a1.indexOf("!");
  let sheet = bang >= 0 ? a1.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'") : defaultSheet;
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet)) sheet = `'${sheet.replace(/'/g, "''")}'`;
  const body = (bang >= 0 ? a1.slice(bang + 1) : a1).replace(/\$/g, "");
  const [c1, c2] = body.split(":");
  return c2 ? `${sheet}.${c1}:${sheet}.${c2}` : `${sheet}.${c1}`;
}

/** Build a fresh <calcext:sparkline-groups> for a sheet's sparklines (one group per distinct style). */
function buildGroups(doc: Document, sheet: Sheet): Element {
  const groups = doc.createElementNS(CALC, "calcext:sparkline-groups");
  const byStyle = new Map<string, { style: SparkStyle; items: { host: { r: number; c: number }; dataRef: string }[] }>();
  for (const sp of sheet.sparklines ?? []) {
    const k = `${sp.type}|${sp.color}|${sp.negColor ?? ""}`;
    const g = byStyle.get(k) ?? { style: { type: sp.type, color: sp.color, negColor: sp.negColor }, items: [] };
    g.items.push({ host: sp.host, dataRef: sp.dataRef });
    byStyle.set(k, g);
  }
  for (const { style, items } of byStyle.values()) {
    const group = doc.createElementNS(CALC, "calcext:sparkline-group");
    group.setAttributeNS(CALC, "calcext:type", style.type);
    group.setAttributeNS(CALC, "calcext:color-series", style.color);
    if (style.negColor) group.setAttributeNS(CALC, "calcext:color-negative", style.negColor);
    // Sensible defaults so LibreOffice renders without the (optional) marker/axis attributes.
    group.setAttributeNS(CALC, "calcext:color-axis", "#000000");
    group.setAttributeNS(CALC, "calcext:line-width", "1");
    group.setAttributeNS(CALC, "calcext:display-empty-cells-as", "gap");
    for (const a of ["markers", "high", "low", "first", "last", "negative", "display-x-axis", "display-hidden", "right-to-left"]) group.setAttributeNS(CALC, `calcext:${a}`, "false");
    for (const a of ["min-axis-type", "max-axis-type"]) group.setAttributeNS(CALC, `calcext:${a}`, "individual");
    const list = doc.createElementNS(CALC, "calcext:sparklines");
    for (const it of items) {
      const el = doc.createElementNS(CALC, "calcext:sparkline");
      el.setAttributeNS(CALC, "calcext:cell-address", a1ToOdfAddr(`${sheet.name}!${colToLetters(it.host.c)}${it.host.r}`, sheet.name));
      el.setAttributeNS(CALC, "calcext:data-range", a1ToOdfAddr(it.dataRef, sheet.name));
      list.appendChild(el);
    }
    group.appendChild(list);
    groups.appendChild(group);
  }
  return groups;
}

/** Persist authored/edited sparklines into their tables' calcext:sparkline-groups. Call AFTER
    writeOds has serialized content.xml (so the rows are in their final order). */
export function writeOdsSparklines(wb: Workbook): void {
  if (!wb.sheets.some((s) => s.sparklinesDirty)) return;
  const doc = parseXmlOpt(wb.files["content.xml"]);
  if (!doc) return;
  const tables = Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "table" && e.namespaceURI === ODS.table);
  for (const sheet of wb.sheets) {
    if (!sheet.sparklinesDirty) continue;
    const table = tables.find((t) => (t.getAttributeNS(ODS.table, "name") || t.getAttribute("table:name")) === sheet.name);
    if (!table) continue;
    for (const g of Array.from(table.children)) if (g.localName === "sparkline-groups") table.removeChild(g);
    if (sheet.sparklines?.length) table.appendChild(buildGroups(doc, sheet)); // after the rows
    sheet.sparklinesDirty = false;
  }
  wb.files["content.xml"] = serializeXml(doc);
}
