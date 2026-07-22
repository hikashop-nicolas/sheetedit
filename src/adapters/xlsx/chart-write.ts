import { parseXmlOpt, serializeXml, type Sheet, type Workbook } from "../../core/model";
import { pxToEmu, type ChartModel, type ChartSeries } from "../../core/chart-model";
import { resolveLabels, resolveNumbers, seriesName } from "../../core/chart-data";

// Write created / edited charts to xlsx DrawingML. Only dirty charts are emitted; a chart read
// from the file and left untouched keeps its original parts verbatim. A created chart gets a new
// chart part + a two-cell anchor in the sheet's drawing (creating the drawing + registrations if
// needed); an edited chart's part is rewritten and its anchor updated in place.

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const strCache = (vals: string[]): string => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("")}</c:strCache>`;
const numCache = (vals: (number | null)[]): string => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (v == null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:numCache>`;
const strRef = (ref: string | undefined, cache: string[]): string => `<c:strRef><c:f>${esc(ref ?? "")}</c:f>${strCache(cache)}</c:strRef>`;
const numRef = (ref: string | undefined, cache: (number | null)[]): string => `<c:numRef><c:f>${esc(ref ?? "")}</c:f>${numCache(cache)}</c:numRef>`;

function serCategory(wb: Workbook, s: ChartSeries, i: number, catRef: string | undefined, catLabels: string[]): string {
  const name = seriesName(wb, s.name) ?? `Series ${i + 1}`;
  const nameRef = typeof s.name === "object" ? s.name : undefined;
  const tx = nameRef?.ref ? `<c:tx>${strRef(nameRef.ref, [name])}</c:tx>` : `<c:tx><c:v>${esc(name)}</c:v></c:tx>`;
  const spPr = s.color ? `<c:spPr><a:solidFill><a:srgbClr val="${s.color.replace("#", "")}"/></a:solidFill></c:spPr>` : "";
  const cat = catRef ? `<c:cat>${strRef(catRef, catLabels)}</c:cat>` : "";
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${cat}<c:val>${numRef(s.values.ref, resolveNumbers(wb, s.values))}</c:val></c:ser>`;
}
function serXY(wb: Workbook, s: ChartSeries, i: number): string {
  const name = seriesName(wb, s.name) ?? `Series ${i + 1}`;
  const nameRef = typeof s.name === "object" ? s.name : undefined;
  const tx = nameRef?.ref ? `<c:tx>${strRef(nameRef.ref, [name])}</c:tx>` : `<c:tx><c:v>${esc(name)}</c:v></c:tx>`;
  const x = `<c:xVal>${numRef(s.xValues?.ref, resolveNumbers(wb, s.xValues))}</c:xVal>`;
  const y = `<c:yVal>${numRef(s.values.ref, resolveNumbers(wb, s.values))}</c:yVal>`;
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${x}${y}<c:smooth val="0"/></c:ser>`;
}

const catAx = (id: number, cross: number, pos: string): string => `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:crossAx val="${cross}"/></c:catAx>`;
const valAx = (id: number, cross: number, pos: string): string => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:crossAx val="${cross}"/></c:valAx>`;

/** Generate the DrawingML chart part for a model, embedding resolved values as caches. */
export function chartXml(model: ChartModel, wb: Workbook): string {
  const catRef = model.categories?.ref;
  const catLabels = resolveLabels(wb, model.categories);
  const AX1 = 111111111;
  const AX2 = 222222222;
  let body: string;
  if (model.kind === "scatter" || model.kind === "bubble") {
    const sers = model.series.map((s, i) => serXY(wb, s, i)).join("");
    body = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:scatterChart>${valAx(AX1, AX2, "b")}${valAx(AX2, AX1, "l")}`;
  } else {
    const sers = model.series.map((s, i) => serCategory(wb, s, i, catRef, catLabels)).join("");
    const group = model.stacked ? "stacked" : model.kind === "line" || model.kind === "area" ? "standard" : "clustered";
    if (model.kind === "pie" || model.kind === "doughnut") {
      body = `<c:${model.kind}Chart><c:varyColors val="1"/>${sers}${model.kind === "doughnut" ? '<c:holeSize val="50"/>' : ""}</c:${model.kind}Chart>`;
    } else if (model.kind === "radar") {
      body = `<c:radarChart><c:radarStyle val="marker"/>${sers}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:radarChart>${catAx(AX1, AX2, "b")}${valAx(AX2, AX1, "l")}`;
    } else if (model.kind === "line") {
      body = `<c:lineChart><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}<c:marker val="1"/><c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:lineChart>${catAx(AX1, AX2, "b")}${valAx(AX2, AX1, "l")}`;
    } else if (model.kind === "area") {
      body = `<c:areaChart><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:areaChart>${catAx(AX1, AX2, "b")}${valAx(AX2, AX1, "l")}`;
    } else {
      const dir = model.kind === "bar" ? "bar" : "col";
      const cAxPos = model.kind === "bar" ? "l" : "b";
      const vAxPos = model.kind === "bar" ? "b" : "l";
      body = `<c:barChart><c:barDir val="${dir}"/><c:grouping val="${group}"/><c:varyColors val="0"/>${sers}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:barChart>${catAx(AX1, AX2, cAxPos)}${valAx(AX2, AX1, vAxPos)}`;
    }
  }
  const title = model.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${esc(model.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : `<c:autoTitleDeleted val="1"/>`;
  const legend = model.legend?.show ? `<c:legend><c:legendPos val="${(model.legend.pos ?? "b")[0]}"/><c:overlay val="0"/></c:legend>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart>${title}<c:plotArea><c:layout/>${body}</c:plotArea>${legend}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function anchorXml(model: ChartModel, chartRid: string, frameId: number): string {
  const a = model.anchor;
  const pt = (tag: string, col: number, colOff: number, row: number, rowOff: number): string =>
    `<xdr:${tag}><xdr:col>${col - 1}</xdr:col><xdr:colOff>${pxToEmu(colOff)}</xdr:colOff><xdr:row>${row - 1}</xdr:row><xdr:rowOff>${pxToEmu(rowOff)}</xdr:rowOff></xdr:${tag}>`;
  return `<xdr:twoCellAnchor>${pt("from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff)}${pt("to", a.toCol, a.toColOff, a.toRow, a.toRowOff)}` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${frameId}" name="Chart ${frameId}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="${chartRid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";

function addContentType(wb: Workbook, partPath: string, ct: string): void {
  const doc = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (!doc || doc.documentElement.localName !== "Types") return;
  if (Array.from(doc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === `/${partPath}`)) return;
  const ov = doc.createElement("Override");
  ov.setAttribute("PartName", `/${partPath}`);
  ov.setAttribute("ContentType", ct);
  doc.documentElement.appendChild(ov);
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

/** Add a relationship to a .rels part (creating it), returning the new id. */
function addRel(wb: Workbook, relsPath: string, type: string, target: string): string {
  let doc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (!doc) doc = parseXmlOpt(new TextEncoder().encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`))!;
  const ids = new Set(Array.from(doc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let n = 1;
  while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const rel = doc.createElement("Relationship");
  rel.setAttribute("Id", id);
  rel.setAttribute("Type", type);
  rel.setAttribute("Target", target);
  doc.documentElement.appendChild(rel);
  wb.files[relsPath] = serializeXml(doc);
  return id;
}

const uniquePath = (wb: Workbook, dir: string, base: string, ext: string): { path: string; n: number } => {
  let n = 1;
  while (wb.files[`${dir}/${base}${n}.${ext}`]) n++;
  return { path: `${dir}/${base}${n}.${ext}`, n };
};

/** The drawing part for a sheet (its rels reference it), creating and wiring one if absent. */
function ensureSheetDrawing(wb: Workbook, sheet: Sheet): string {
  const sheetRels = sheet.path!.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const relsDoc = wb.files[sheetRels] ? parseXmlOpt(wb.files[sheetRels]) : undefined;
  const existing = relsDoc && Array.from(relsDoc.getElementsByTagName("Relationship")).find((r) => /drawing/i.test(r.getAttribute("Type") ?? "") && /drawings\//i.test(r.getAttribute("Target") ?? ""));
  if (existing) {
    const parts: string[] = [];
    for (const seg of `xl/worksheets/${existing.getAttribute("Target")}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    return parts.join("/");
  }
  const { path, n } = uniquePath(wb, "xl/drawings", "drawing", "xml");
  wb.files[path] = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}"></xdr:wsDr>`);
  addContentType(wb, path, CT_DRAWING);
  const rid = addRel(wb, sheetRels, `${R}/drawing`, `../drawings/drawing${n}.xml`);
  // Add <drawing r:id> to the worksheet XML (namespaced, after sheetData per the schema).
  const wsDoc = parseXmlOpt(wb.files[sheet.path!]);
  const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  if (wsDoc && !Array.from(wsDoc.getElementsByTagName("*")).some((e) => e.localName === "drawing")) {
    const d = wsDoc.createElementNS(SS, "drawing");
    d.setAttributeNS(R, "r:id", rid);
    wsDoc.documentElement.appendChild(d);
    wb.files[sheet.path!] = serializeXml(wsDoc);
    sheet.doc = wsDoc;
    sheet.sheetData = wsDoc.getElementsByTagName("sheetData")[0] ?? sheet.sheetData;
  }
  return path;
}

let frameSeq = 1000;

function createNew(wb: Workbook, sheet: Sheet, model: ChartModel): void {
  const { path: chartPath, n } = uniquePath(wb, "xl/charts", "chart", "xml");
  wb.files[chartPath] = new TextEncoder().encode(chartXml(model, wb));
  addContentType(wb, chartPath, CT_CHART);
  const drawingPath = ensureSheetDrawing(wb, sheet);
  const drawRels = drawingPath.replace(/drawings\/(drawing[^/]+\.xml)$/i, "drawings/_rels/$1.rels");
  const rid = addRel(wb, drawRels, `${R}/chart`, `../charts/chart${n}.xml`);
  // Append the anchor to the drawing XML (before </xdr:wsDr>).
  const xml = new TextDecoder().decode(wb.files[drawingPath]);
  const anchor = anchorXml(model, rid, ++frameSeq);
  wb.files[drawingPath] = new TextEncoder().encode(xml.replace(/<\/xdr:wsDr>\s*$/, `${anchor}</xdr:wsDr>`));
  model.original = { partPath: chartPath, drawingPath };
}

function rewriteExisting(wb: Workbook, model: ChartModel): void {
  const { partPath, drawingPath } = model.original!;
  if (partPath && wb.files[partPath]) wb.files[partPath] = new TextEncoder().encode(chartXml(model, wb));
  if (!drawingPath || !wb.files[drawingPath]) return;
  const doc = parseXmlOpt(wb.files[drawingPath]);
  if (!doc) return;
  const drawRels = drawingPath.replace(/drawings\/(drawing[^/]+\.xml)$/i, "drawings/_rels/$1.rels");
  const relsDoc = wb.files[drawRels] ? parseXmlOpt(wb.files[drawRels]) : undefined;
  const ridOfPart = new Map<string, string>(); // rId -> resolved part path
  if (relsDoc) for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
    const parts: string[] = [];
    for (const seg of `xl/drawings/${r.getAttribute("Target")}`.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    ridOfPart.set(r.getAttribute("Id") ?? "", parts.join("/"));
  }
  // Find the anchor whose graphicFrame chart r:id resolves to this chart's part; rewrite its from/to.
  const anchors = Array.from(doc.documentElement.children).filter((e) => /Anchor$/.test(e.localName));
  const a = model.anchor;
  for (const anchorEl of anchors) {
    const chartEl = Array.from(anchorEl.getElementsByTagName("*")).find((e) => e.localName === "chart");
    const rid = chartEl?.getAttributeNS(R, "id") ?? chartEl?.getAttribute("r:id");
    if (!rid || ridOfPart.get(rid) !== partPath) continue;
    const setPt = (tag: string, col: number, colOff: number, row: number, rowOff: number): void => {
      const p = Array.from(anchorEl.children).find((c) => c.localName === tag);
      if (!p) return;
      const set = (local: string, v: number): void => { const e = Array.from(p.children).find((x) => x.localName === local); if (e) e.textContent = String(v); };
      set("col", col - 1); set("colOff", pxToEmu(colOff)); set("row", row - 1); set("rowOff", pxToEmu(rowOff));
    };
    setPt("from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff);
    setPt("to", a.toCol, a.toColOff, a.toRow, a.toRowOff);
    break;
  }
  wb.files[drawingPath] = serializeXml(doc);
}

/** Persist all dirty charts to the workbook's DrawingML parts. */
export function writeXlsxCharts(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    if (!sheet.path) continue;
    for (const chart of sheet.charts ?? []) {
      if (!chart.dirty) continue;
      if (chart.original?.partPath) rewriteExisting(wb, chart);
      else createNew(wb, sheet, chart);
      chart.dirty = false;
    }
  }
}
