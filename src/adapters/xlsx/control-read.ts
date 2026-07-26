import { parseXmlOpt, type SheetControl, type Workbook } from "../../core/model";
import { anchorOf, relMap, resolvePart } from "./chart-read";

// A form control is spread across three parts:
//   the worksheet   <controls><control shapeId r:id name>  (often inside mc:AlternateContent)
//   xl/ctrlProps/*  <formControlPr objectType checked fmlaLink fmlaRange sel val min max .../>
//   the VML drawing <v:shape><x:ClientData ObjectType>  - the label, the position, and on older
//                   files the state as well, since they predate ctrlProps entirely.
//
// So ctrlProps is read when present and the VML fills in the rest (and stands in for it when the
// file has no ctrlProps at all). ActiveX controls are a different thing entirely - Windows COM in a
// .bin part - and are left untouched rather than pretended at.

/** objectType / VML ObjectType -> the model's kind. Both spellings appear in the wild. */
const KINDS: Record<string, SheetControl["kind"]> = {
  checkbox: "checkbox",
  radio: "radio",
  drop: "dropdown",
  list: "list",
  spin: "spin",
  scroll: "scroll",
  button: "button",
  label: "label",
  gbox: "groupBox",
  groupbox: "groupBox",
};
const kindOf = (raw: string | null | undefined): SheetControl["kind"] => KINDS[(raw ?? "").toLowerCase()] ?? "label";

const num = (v: string | null | undefined): number | undefined => {
  if (v == null || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
};

/** The text a VML control shape shows, which is where a control's label actually lives. */
function vmlLabel(shape: Element): string | undefined {
  const box = Array.from(shape.getElementsByTagName("*")).find((e) => e.localName === "textbox");
  const text = box?.textContent?.replace(/\s+/g, " ").trim();
  return text || undefined;
}

/** One `<x:ClientData>` child's text, e.g. FmlaLink or Checked. */
const clientVal = (client: Element, local: string): string | undefined =>
  Array.from(client.children).find((e) => e.localName === local)?.textContent?.trim() || undefined;

/** Everything the VML knows about a control shape, keyed by its shape id. */
function readVml(files: Record<string, Uint8Array>, path: string): Map<string, { el: Element; client: Element; kind: SheetControl["kind"]; label?: string }> {
  const out = new Map<string, { el: Element; client: Element; kind: SheetControl["kind"]; label?: string }>();
  const doc = files[path] ? parseXmlOpt(files[path]) : undefined;
  if (!doc) return out;
  for (const shape of Array.from(doc.getElementsByTagName("*"))) {
    if (shape.localName !== "shape") continue;
    const client = Array.from(shape.children).find((e) => e.localName === "ClientData");
    if (!client) continue;
    // "_x0000_s1025" -> "1025": the worksheet's @shapeId is the trailing number, and the "_s"
    // prefix has to be matched from the END (a lazy match from the start stops at the first "_").
    const raw = shape.getAttribute("id") ?? "";
    const id = (/_s(\d+)$/.exec(raw) ?? /(\d+)$/.exec(raw))?.[1];
    if (!id) continue;
    out.set(id, { el: shape, client, kind: kindOf(client.getAttribute("ObjectType")), label: vmlLabel(shape) });
  }
  return out;
}

/** Apply a `<formControlPr>` onto a control. */
function applyProps(ctl: SheetControl, el: Element): void {
  ctl.kind = kindOf(el.getAttribute("objectType"));
  const checked = el.getAttribute("checked");
  if (checked != null) ctl.checked = checked.toLowerCase() === "checked" || checked === "1";
  const link = el.getAttribute("fmlaLink");
  if (link) ctl.linkedCell = link;
  const range = el.getAttribute("fmlaRange");
  if (range) ctl.sourceRange = range;
  const sel = num(el.getAttribute("sel"));
  if (sel != null) ctl.selected = sel;
  for (const [attr, key] of [["val", "value"], ["min", "min"], ["max", "max"], ["inc", "inc"]] as const) {
    const v = num(el.getAttribute(attr));
    if (v != null) ctl[key] = v;
  }
}

/** Apply what the VML says, for a file whose control has no ctrlProps part. */
function applyClientData(ctl: SheetControl, client: Element): void {
  const checked = clientVal(client, "Checked");
  if (checked != null && ctl.checked === undefined) ctl.checked = checked !== "0";
  const link = clientVal(client, "FmlaLink");
  if (link && !ctl.linkedCell) ctl.linkedCell = link;
  const range = clientVal(client, "FmlaRange");
  if (range && !ctl.sourceRange) ctl.sourceRange = range;
  const sel = num(clientVal(client, "Sel"));
  if (sel != null && ctl.selected === undefined) ctl.selected = sel;
  for (const [tag, key] of [["Val", "value"], ["Min", "min"], ["Max", "max"], ["Inc", "inc"]] as const) {
    const v = num(clientVal(client, tag));
    if (v != null && ctl[key] === undefined) ctl[key] = v;
  }
}

/** `<x:Anchor>fromCol,fromColOff,fromRow,fromRowOff,toCol,toColOff,toRow,toRowOff</x:Anchor>`. */
function vmlAnchor(client: Element): SheetControl["anchor"] | undefined {
  const parts = (clientVal(client, "Anchor") ?? "").split(",").map((p) => Number(p.trim()));
  if (parts.length < 8 || parts.some((n) => !Number.isFinite(n))) return undefined;
  // The offsets are in VML units rather than pixels; the cell span is what matters for placement.
  return { fromCol: parts[0]! + 1, fromRow: parts[2]! + 1, fromColOff: 0, fromRowOff: 0, toCol: parts[4]! + 1, toRow: parts[6]! + 1, toColOff: 0, toRowOff: 0 };
}

/** Populate sheet.controls for every sheet that has form controls. */
export function readXlsxControls(wb: Workbook, files: Record<string, Uint8Array>): void {
  for (const sheet of wb.sheets) {
    if (!sheet.path || !sheet.doc) continue;
    const relsPath = sheet.path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
    const rels = relMap(files, relsPath);
    // The legacy VML drawing carries the shapes; a sheet can only have one.
    const vmlRel = rels.byType.find((r) => /vmlDrawing/i.test(r.type));
    const vmlPath = vmlRel ? resolvePart("xl/worksheets", vmlRel.target) : undefined;
    const vml = vmlPath ? readVml(files, vmlPath) : new Map();

    const controls: SheetControl[] = [];
    for (const el of Array.from(sheet.doc.getElementsByTagName("*"))) {
      if (el.localName !== "control") continue;
      const shapeId = el.getAttribute("shapeId") ?? undefined;
      const rid = Array.from(el.attributes).find((a) => a.localName === "id" && a.name !== "id")?.value ?? el.getAttribute("r:id") ?? undefined;
      const ctl: SheetControl = { kind: "label", name: el.getAttribute("name") ?? `Control ${controls.length + 1}`, shapeId, vmlPath };

      const target = rid ? rels.byId.get(rid) : undefined;
      const propsPath = target ? resolvePart("xl/worksheets", target) : undefined;
      const propsDoc = propsPath && files[propsPath] ? parseXmlOpt(files[propsPath]) : undefined;
      if (propsDoc?.documentElement) {
        ctl.propsPath = propsPath;
        applyProps(ctl, propsDoc.documentElement);
      }
      // The anchor is on the worksheet's controlPr when the file is modern enough to have one.
      const anchorEl = Array.from(el.getElementsByTagName("*")).find((e) => e.localName === "anchor");
      if (anchorEl) ctl.anchor = anchorOf(anchorEl) ?? undefined;

      const shape = shapeId ? vml.get(shapeId) : undefined;
      if (shape) {
        if (!propsDoc) ctl.kind = shape.kind;
        ctl.label ??= shape.label;
        applyClientData(ctl, shape.client);
        ctl.anchor ??= vmlAnchor(shape.client);
      }
      controls.push(ctl);
    }
    if (controls.length) sheet.controls = controls;
  }
}

/** Whether the workbook carries ActiveX controls, which are Windows COM and left untouched. */
export const hasActiveX = (files: Record<string, Uint8Array>): boolean =>
  Object.keys(files).some((p) => /^xl\/activeX\//i.test(p));
