import { parseXmlOpt, serializeXml, type Workbook } from "../../core/model";
import type { ChartModel, ChartRef } from "../../core/chart-model";
import { resolveLabels, resolveNumbers, seriesName } from "../../core/chart-data";

// Write created / edited charts to ODS as embedded chart objects. Each dirty chart becomes an
// "Object N/" sub-document (content.xml with chart:chart + an internal data table), listed in the
// manifest and referenced by a draw:frame in the sheet's content.xml. Only dirty charts are
// written; untouched originals are preserved. Runs AFTER writeOds has regenerated content.xml.
//
// Positioning note: the frame is placed absolutely (svg:x/y/width/height) from the grid origin
// using sheetedit's default cell metrics, so the on-screen position round-trips but is only exact
// for default-width sheets (a documented approximation).

const OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK = "http://www.w3.org/1999/xlink";
const MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

const COL_W = 96;
const ROW_H = 24;
const PX_CM = 37.795;
const cm = (px: number): string => `${(px / PX_CM).toFixed(3)}cm`;
const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

const KIND_CLASS: Record<string, string> = {
  column: "bar", bar: "bar", line: "line", area: "area", pie: "circle", doughnut: "ring", scatter: "scatter", bubble: "bubble", radar: "radar",
};

/** "Sheet1!$B$2:$B$3" -> "Sheet1.$B$2:Sheet1.$B$3". */
function toOdsRef(ref: string | undefined): string | undefined {
  if (!ref) return undefined;
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(ref);
  const sheet = m ? (m[1] ?? m[2]) : "";
  const body = m ? m[3] : ref;
  const [a, b] = body.split(":");
  const q = (c: string): string => `${sheet}.${c}`;
  return b ? `${q(a)}:${q(b)}` : q(a);
}
const nameRefOf = (name: string | ChartRef | undefined): ChartRef | undefined => (typeof name === "object" ? name : undefined);

/** The embedded chart object's content.xml. */
function chartContent(model: ChartModel, wb: Workbook): string {
  const cls = KIND_CLASS[model.kind] ?? "bar";
  const title = model.title ? `<chart:title><text:p>${esc(model.title)}</text:p></chart:title>` : "";
  const legend = model.legend?.show ? `<chart:legend chart:legend-position="${({ top: "top", bottom: "bottom", left: "start", right: "end" } as Record<string, string>)[model.legend.pos] ?? "end"}"/>` : "";
  const catRef = toOdsRef(model.categories?.ref);
  const series = model.series.map((s) => {
    const val = toOdsRef(s.values.ref);
    const lbl = toOdsRef(nameRefOf(s.name)?.ref);
    const scls = KIND_CLASS[s.type ?? model.kind] ?? cls; // combo: per-series class
    return `<chart:series chart:style-name="ch-auto"${val ? ` chart:values-cell-range-address="${val}"` : ""}${lbl ? ` chart:label-cell-address="${lbl}"` : ""} chart:class="chart:${scls}"><chart:data-point chart:repeated="${resolveNumbers(wb, s.values).length}"/></chart:series>`;
  }).join("");
  const cats = catRef ? `<chart:categories table:cell-range-address="${catRef}"/>` : "";
  // A minimal internal data table (fallback data). Categories in the first column, one column per series.
  const labels = resolveLabels(wb, model.categories);
  const cols = model.series.map((s) => resolveNumbers(wb, s.values));
  const headerRow = `<table:table-row><table:table-cell/>${model.series.map((s) => `<table:table-cell office:value-type="string"><text:p>${esc(seriesName(wb, s.name) ?? "")}</text:p></table:table-cell>`).join("")}</table:table-row>`;
  const bodyRows = labels.map((lab, r) => `<table:table-row><table:table-cell office:value-type="string"><text:p>${esc(lab)}</text:p></table:table-cell>${cols.map((col) => { const v = col[r]; return v == null ? "<table:table-cell/>" : `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`; }).join("")}</table:table-row>`).join("");
  const localTable = `<table:table table:name="local-table"><table:table-header-columns/><table:table-columns/><table:table-header-rows>${headerRow}</table:table-header-rows><table:table-rows>${bodyRows}</table:table-rows></table:table>`;
  return `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<office:document-content xmlns:office="${OFFICE}" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="${TABLE}" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">` +
    `<office:body><office:chart><chart:chart chart:class="chart:${cls}">${title}${legend}<chart:plot-area>${series}${cats}</chart:plot-area>${localTable}</chart:chart></office:chart></office:body></office:document-content>`;
}

const OBJ_STYLES = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles xmlns:office="${OFFICE}"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>`;

function addManifest(wb: Workbook, dir: string): void {
  const path = "META-INF/manifest.xml";
  let doc = wb.files[path] ? parseXmlOpt(wb.files[path]) : undefined;
  if (!doc) doc = parseXmlOpt(new TextEncoder().encode(`<manifest:manifest xmlns:manifest="${MANIFEST}"></manifest:manifest>`))!;
  const add = (full: string, media: string): void => {
    if (Array.from(doc!.getElementsByTagName("*")).some((e) => e.localName === "file-entry" && (e.getAttributeNS(MANIFEST, "full-path") || e.getAttribute("manifest:full-path")) === full)) return;
    const fe = doc!.createElementNS(MANIFEST, "manifest:file-entry");
    fe.setAttributeNS(MANIFEST, "manifest:full-path", full);
    fe.setAttributeNS(MANIFEST, "manifest:media-type", media);
    doc!.documentElement.appendChild(fe);
  };
  add(`${dir}/`, "application/vnd.oasis.opendocument.chart");
  add(`${dir}/content.xml`, "text/xml");
  add(`${dir}/styles.xml`, "text/xml");
  wb.files[path] = serializeXml(doc);
}

function frameXml(doc: Document, model: ChartModel, dir: string): Element {
  const a = model.anchor;
  const x = (a.fromCol - 1) * COL_W + a.fromColOff;
  const y = (a.fromRow - 1) * ROW_H + a.fromRowOff;
  const w = (a.toCol - a.fromCol) * COL_W + a.toColOff - a.fromColOff;
  const h = (a.toRow - a.fromRow) * ROW_H + a.toRowOff - a.fromRowOff;
  const frame = doc.createElementNS(DRAW, "draw:frame");
  frame.setAttributeNS(SVG, "svg:x", cm(x));
  frame.setAttributeNS(SVG, "svg:y", cm(y));
  frame.setAttributeNS(SVG, "svg:width", cm(Math.max(60, w)));
  frame.setAttributeNS(SVG, "svg:height", cm(Math.max(40, h)));
  const obj = doc.createElementNS(DRAW, "draw:object");
  obj.setAttributeNS(XLINK, "xlink:href", `./${dir}`);
  obj.setAttributeNS(XLINK, "xlink:type", "simple");
  obj.setAttributeNS(XLINK, "xlink:show", "embed");
  obj.setAttributeNS(XLINK, "xlink:actuate", "onLoad");
  frame.appendChild(obj);
  return frame;
}

/** Persist all dirty charts on an ods workbook. Call after writeOds. */
export function writeOdsCharts(wb: Workbook): void {
  const hasDirty = wb.sheets.some((s) => (s.charts ?? []).some((c) => c.dirty));
  if (!hasDirty) return;
  const doc = parseXmlOpt(wb.files["content.xml"]);
  if (!doc) return;
  const tables = Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "table" && e.namespaceURI === TABLE);
  let objN = 1;
  const usedDir = (): string => { while (Object.keys(wb.files).some((f) => f.startsWith(`Object ${objN}/`))) objN++; return `Object ${objN}`; };

  wb.sheets.forEach((sheet) => {
    const table = tables.find((t) => (t.getAttributeNS(TABLE, "name") || t.getAttribute("table:name")) === sheet.name);
    if (!table) return;
    for (const chart of sheet.charts ?? []) {
      if (!chart.dirty) continue;
      const dir = chart.original?.objectDir ?? usedDir();
      wb.files[`${dir}/content.xml`] = new TextEncoder().encode(chartContent(chart, wb));
      if (!wb.files[`${dir}/styles.xml`]) wb.files[`${dir}/styles.xml`] = new TextEncoder().encode(OBJ_STYLES);
      addManifest(wb, dir);
      if (!chart.original?.objectDir) {
        // New chart: insert a frame into the first cell of the table's first row (creating one if needed).
        let row = Array.from(table.getElementsByTagName("*")).find((e) => e.localName === "table-row" && e.namespaceURI === TABLE);
        if (!row) { row = doc.createElementNS(TABLE, "table:table-row"); table.appendChild(row); }
        let cellEl = Array.from(row.children).find((c) => c.localName === "table-cell");
        if (!cellEl) { cellEl = doc.createElementNS(TABLE, "table:table-cell"); row.insertBefore(cellEl, row.firstChild); }
        cellEl.appendChild(frameXml(doc, chart, dir));
        chart.original = { objectDir: dir };
      }
      chart.dirty = false;
      objN++;
    }
  });
  wb.files["content.xml"] = serializeXml(doc);
}
