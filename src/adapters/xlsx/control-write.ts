import { parseXmlOpt, serializeXml, type SheetControl, type Workbook } from "../../core/model";

// A control's state lives in two places: the ctrlProps part and, mirrored, the VML shape's
// <x:ClientData>. Excel reads ctrlProps, older readers only the VML, so both are updated. Every
// other attribute of either part is left exactly as it was.

/** Set (or drop) one attribute on the formControlPr element. */
const setAttr = (el: Element, name: string, v: string | undefined): void => {
  if (v == null) el.removeAttribute(name);
  else el.setAttribute(name, v);
};

/** Set the text of an `<x:...>` child of ClientData, creating it when absent. */
function setClient(doc: Document, client: Element, local: string, value: string | undefined): void {
  const existing = Array.from(client.children).find((e) => e.localName === local);
  if (value == null) {
    if (existing) client.removeChild(existing);
    return;
  }
  if (existing) {
    existing.textContent = value;
    return;
  }
  // Match the prefix the file already uses, rather than assuming "x".
  const prefix = client.prefix ? `${client.prefix}:` : "";
  const el = doc.createElementNS(client.namespaceURI, `${prefix}${local}`);
  el.textContent = value;
  client.appendChild(el);
}

/** Write one control's state into its ctrlProps part. */
function writeProps(wb: Workbook, ctl: SheetControl): void {
  if (!ctl.propsPath) return;
  const file = wb.files[ctl.propsPath];
  const doc = file ? parseXmlOpt(file) : undefined;
  const el = doc?.documentElement;
  if (!doc || !el) return;
  if (ctl.kind === "checkbox" || ctl.kind === "radio") setAttr(el, "checked", ctl.checked ? "Checked" : "Unchecked");
  if (ctl.kind === "dropdown" || ctl.kind === "list") setAttr(el, "sel", String(ctl.selected ?? 0));
  if (ctl.kind === "spin" || ctl.kind === "scroll") setAttr(el, "val", String(ctl.value ?? 0));
  wb.files[ctl.propsPath] = serializeXml(doc);
}

/** Mirror the state into the VML shape, which is all an older reader looks at. */
function writeVml(wb: Workbook, controls: SheetControl[], path: string): void {
  const file = wb.files[path];
  const doc = file ? parseXmlOpt(file) : undefined;
  if (!doc) return;
  const byId = new Map<string, Element>();
  for (const shape of Array.from(doc.getElementsByTagName("*"))) {
    if (shape.localName !== "shape") continue;
    const raw = shape.getAttribute("id") ?? "";
    const id = (/_s(\d+)$/.exec(raw) ?? /(\d+)$/.exec(raw))?.[1];
    const client = Array.from(shape.children).find((e) => e.localName === "ClientData");
    if (id && client) byId.set(id, client);
  }
  let touched = false;
  for (const ctl of controls) {
    const client = ctl.shapeId ? byId.get(ctl.shapeId) : undefined;
    if (!client) continue;
    if (ctl.kind === "checkbox" || ctl.kind === "radio") setClient(doc, client, "Checked", ctl.checked ? "1" : "0");
    if (ctl.kind === "dropdown" || ctl.kind === "list") setClient(doc, client, "Sel", String(ctl.selected ?? 0));
    if (ctl.kind === "spin" || ctl.kind === "scroll") setClient(doc, client, "Val", String(ctl.value ?? 0));
    touched = true;
  }
  if (touched) wb.files[path] = serializeXml(doc);
}

/** Persist every control whose state the user changed. */
export function writeXlsxControls(wb: Workbook): void {
  const byVml = new Map<string, SheetControl[]>();
  for (const sheet of wb.sheets) {
    for (const ctl of sheet.controls ?? []) {
      if (!ctl.dirty) continue;
      ctl.dirty = false;
      writeProps(wb, ctl);
      if (ctl.vmlPath) byVml.set(ctl.vmlPath, [...(byVml.get(ctl.vmlPath) ?? []), ctl]);
    }
  }
  // One pass per VML part: several controls usually share a sheet's single drawing.
  for (const [path, controls] of byVml) writeVml(wb, controls, path);
}
