import { colToLetters, parseXmlOpt, serializeXml, type Sheet, type SheetControl, type Workbook } from "../../core/model";
import { SS_MAIN } from "./shared";
import { insertWsChild } from "./write";

// Creating a form control means building the three parts Excel expects and wiring them together:
//   xl/ctrlProps/ctrlPropN.xml   the state (objectType, checked/sel/val, fmlaLink, fmlaRange)
//   the sheet's VML drawing      the shape that positions it, plus <x:ClientData> mirroring state
//   the worksheet                a <control> entry pointing at both, and <legacyDrawing> at the VML
// plus a content-type override for the props part, a Default for the .vml extension, and the
// relationships. Miss any one of them and Excel silently drops the control.

const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OREL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const VML_NS = "urn:schemas-microsoft-com:vml";
const XL_NS = "urn:schemas-microsoft-com:office:excel";
const CT_CTRLPROP = "application/vnd.ms-excel.controlproperties+xml";
const CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const REL_CTRLPROP = `${OREL}/ctrlProp`;
const REL_VML = `${OREL}/vmlDrawing`;

/**
 * The two parts spell some of these differently: ctrlProps says "CheckBox", the VML ClientData says
 * "Checkbox". Our reader is case-insensitive, but Excel's is not, so each gets its own spelling.
 */
const VML_OBJECT_TYPES: Partial<Record<SheetControl["kind"], string>> = { checkbox: "Checkbox", groupBox: "GBox" };

/** The objectType a kind is stored as in ctrlProps. */
const OBJECT_TYPES: Partial<Record<SheetControl["kind"], string>> = {
  checkbox: "CheckBox",
  radio: "Radio",
  dropdown: "Drop",
  list: "List",
  spin: "Spin",
  scroll: "Scroll",
  button: "Button",
  label: "Label",
  groupBox: "GBox",
};

function addContentTypeOverride(wb: Workbook, partPath: string, ct: string): void {
  const doc = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (!doc || doc.documentElement.localName !== "Types") return;
  if (Array.from(doc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === `/${partPath}`)) return;
  const ov = doc.createElementNS(CT_NS, "Override");
  ov.setAttribute("PartName", `/${partPath}`);
  ov.setAttribute("ContentType", ct);
  doc.documentElement.appendChild(ov);
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

/** A Default for an extension, which is how .vml parts are declared. */
function addContentTypeDefault(wb: Workbook, ext: string, ct: string): void {
  const doc = parseXmlOpt(wb.files["[Content_Types].xml"]);
  if (!doc || doc.documentElement.localName !== "Types") return;
  if (Array.from(doc.getElementsByTagName("Default")).some((d) => (d.getAttribute("Extension") ?? "").toLowerCase() === ext)) return;
  const def = doc.createElementNS(CT_NS, "Default");
  def.setAttribute("Extension", ext);
  def.setAttribute("ContentType", ct);
  doc.documentElement.insertBefore(def, doc.documentElement.firstChild);
  wb.files["[Content_Types].xml"] = serializeXml(doc);
}

function addRel(wb: Workbook, relsPath: string, type: string, target: string): string {
  const doc = wb.files[relsPath]
    ? parseXmlOpt(wb.files[relsPath])!
    : parseXmlOpt(new TextEncoder().encode(`<Relationships xmlns="${REL_NS}"></Relationships>`))!;
  const ids = new Set(Array.from(doc.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let n = 1;
  while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const rel = doc.createElementNS(REL_NS, "Relationship");
  rel.setAttribute("Id", id);
  rel.setAttribute("Type", type);
  rel.setAttribute("Target", target);
  doc.documentElement.appendChild(rel);
  wb.files[relsPath] = serializeXml(doc);
  return id;
}

/** An existing relationship of this type, so a second control reuses the sheet's one VML drawing. */
function findRel(wb: Workbook, relsPath: string, type: string): { id: string; target: string } | undefined {
  const doc = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]) : undefined;
  if (!doc) return undefined;
  for (const r of Array.from(doc.getElementsByTagName("Relationship"))) {
    if (r.getAttribute("Type") === type) return { id: r.getAttribute("Id") ?? "", target: r.getAttribute("Target") ?? "" };
  }
  return undefined;
}

const relsPathFor = (sheetPath: string): string => sheetPath.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
const nextFreePath = (wb: Workbook, dir: string, base: string, ext: string): string => {
  let n = 1;
  while (wb.files[`${dir}/${base}${n}.${ext}`]) n++;
  return `${dir}/${base}${n}.${ext}`;
};

/** The next shape id not already used by this sheet's VML. */
function nextShapeId(wb: Workbook, vmlPath: string): number {
  let max = 1024;
  const doc = wb.files[vmlPath] ? parseXmlOpt(wb.files[vmlPath]) : undefined;
  if (doc) {
    for (const shape of Array.from(doc.getElementsByTagName("*"))) {
      if (shape.localName !== "shape") continue;
      const id = Number((/(\d+)$/.exec(shape.getAttribute("id") ?? "") ?? [])[1] ?? 0);
      if (id > max) max = id;
    }
  }
  return max + 1;
}

export interface NewControl {
  kind: SheetControl["kind"];
  label?: string;
  linkedCell?: string;
  sourceRange?: string;
  /** Where it sits, as a cell span (1-based, inclusive). */
  at: { r1: number; c1: number; r2: number; c2: number };
}

/**
 * Add a form control to a sheet, creating every part it needs. Returns the model entry, already
 * pushed onto the sheet, or null when the workbook is not one we can add parts to.
 */
export function createXlsxControl(wb: Workbook, sheet: Sheet, spec: NewControl): SheetControl | null {
  const doc = sheet.doc;
  const ws = doc?.documentElement;
  if (!doc || !ws || !sheet.path) return null;
  const relsPath = relsPathFor(sheet.path);

  // The sheet has one VML drawing shared by all its controls; make it only if there is none.
  const vmlRel = findRel(wb, relsPath, REL_VML);
  let vmlPath: string;
  let vmlRelId: string;
  if (vmlRel) {
    vmlPath = vmlRel.target.replace(/^\.\.\//, "xl/");
    vmlRelId = vmlRel.id;
  } else {
    vmlPath = nextFreePath(wb, "xl/drawings", "vmlDrawing", "vml");
    wb.files[vmlPath] = new TextEncoder().encode(
      `<xml xmlns:v="${VML_NS}" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="${XL_NS}">` +
      `<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">` +
      `<v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/></v:shapetype></xml>`,
    );
    addContentTypeDefault(wb, "vml", CT_VML);
    vmlRelId = addRel(wb, relsPath, REL_VML, `../${vmlPath.replace(/^xl\//, "")}`);
  }

  const objectType = OBJECT_TYPES[spec.kind] ?? "Label";
  const shapeId = nextShapeId(wb, vmlPath);
  const linked = spec.linkedCell ? spec.linkedCell.replace(/\$/g, "") : undefined;
  const absLink = linked ? `$${linked.replace(/^([A-Za-z]+)(\d+)$/, "$1$$$2")}` : undefined;

  // 1. the props part
  const propsPath = nextFreePath(wb, "xl/ctrlProps", "ctrlProp", "xml");
  const attrs = [`objectType="${objectType}"`];
  if (spec.kind === "checkbox" || spec.kind === "radio") attrs.push(`checked="Unchecked"`);
  if (spec.kind === "dropdown" || spec.kind === "list") attrs.push(`dropLines="8"`, `sel="0"`);
  if (spec.kind === "spin" || spec.kind === "scroll") attrs.push(`min="0"`, `max="100"`, `inc="1"`, `val="0"`);
  if (absLink) attrs.push(`fmlaLink="${absLink}"`);
  if (spec.sourceRange) attrs.push(`fmlaRange="${spec.sourceRange}"`);
  attrs.push(`lockText="1"`, `noThreeD="1"`);
  wb.files[propsPath] = new TextEncoder().encode(`<formControlPr xmlns="${X14}" ${attrs.join(" ")}/>`);
  addContentTypeOverride(wb, propsPath, CT_CTRLPROP);
  const propsRelId = addRel(wb, relsPath, REL_CTRLPROP, `../${propsPath.replace(/^xl\//, "")}`);

  // 2. the VML shape
  const vmlDoc = parseXmlOpt(wb.files[vmlPath]!);
  if (vmlDoc) {
    const { r1, c1, r2, c2 } = spec.at;
    const shape = vmlDoc.createElementNS(VML_NS, "v:shape");
    shape.setAttribute("id", `_x0000_s${shapeId}`);
    shape.setAttribute("type", "#_x0000_t201");
    shape.setAttribute("style", "position:absolute");
    if (spec.label) {
      // Excel nests the caption in <div><font>; LibreOffice reads the label only in that shape.
      const box = vmlDoc.createElementNS(VML_NS, "v:textbox");
      const div = vmlDoc.createElement("div");
      const font = vmlDoc.createElement("font");
      font.textContent = spec.label;
      div.appendChild(font);
      box.appendChild(div);
      shape.appendChild(box);
    }
    const client = vmlDoc.createElementNS(XL_NS, "x:ClientData");
    client.setAttribute("ObjectType", VML_OBJECT_TYPES[spec.kind] ?? objectType);
    const add = (local: string, text: string): void => {
      const el = vmlDoc.createElementNS(XL_NS, `x:${local}`);
      el.textContent = text;
      client.appendChild(el);
    };
    // The VML anchor is 0-based and states fromCol,fromColOff,fromRow,fromRowOff, then the to pair.
    add("Anchor", `${c1 - 1},0,${r1 - 1},0,${c2 - 1},0,${r2 - 1},0`);
    add("AutoFill", "False");
    if (spec.kind === "checkbox" || spec.kind === "radio") add("Checked", "0");
    if (spec.kind === "dropdown" || spec.kind === "list") { add("Sel", "0"); add("DropLines", "8"); }
    if (spec.kind === "spin" || spec.kind === "scroll") { add("Val", "0"); add("Min", "0"); add("Max", "100"); add("Inc", "1"); }
    if (absLink) add("FmlaLink", absLink);
    if (spec.sourceRange) add("FmlaRange", spec.sourceRange);
    shape.appendChild(client);
    vmlDoc.documentElement.appendChild(shape);
    wb.files[vmlPath] = serializeXml(vmlDoc);
  }

  // 3. the worksheet entry, plus the legacyDrawing that points at the VML
  if (!Array.from(ws.children).some((e) => e.localName === "legacyDrawing")) {
    const ld = doc.createElementNS(ws.namespaceURI || SS_MAIN, "legacyDrawing");
    ld.setAttributeNS(OREL, "r:id", vmlRelId);
    insertWsChild(ws, ld);
  }
  let controls = Array.from(ws.children).find((e) => e.localName === "controls");
  if (!controls) {
    controls = doc.createElementNS(ws.namespaceURI || SS_MAIN, "controls");
    insertWsChild(ws, controls);
  }
  const name = `${objectType} ${(sheet.controls?.length ?? 0) + 1}`;
  // Excel wraps each control in AlternateContent so an older reader falls back to the VML alone.
  const alt = doc.createElementNS(MC, "mc:AlternateContent");
  const choice = doc.createElementNS(MC, "mc:Choice");
  choice.setAttribute("Requires", "x14");
  const control = doc.createElementNS(ws.namespaceURI || SS_MAIN, "control");
  control.setAttribute("shapeId", String(shapeId));
  control.setAttributeNS(OREL, "r:id", propsRelId);
  control.setAttribute("name", name);
  const pr = doc.createElementNS(ws.namespaceURI || SS_MAIN, "controlPr");
  pr.setAttribute("defaultSize", "0");
  pr.setAttribute("autoFill", "0");
  pr.setAttribute("autoLine", "0");
  const anchor = doc.createElementNS(ws.namespaceURI || SS_MAIN, "anchor");
  anchor.setAttribute("moveWithCells", "1");
  const corner = (local: string, col: number, row: number): Element => {
    const el = doc.createElementNS(ws.namespaceURI || SS_MAIN, local);
    for (const [tag, v] of [["col", col - 1], ["colOff", 0], ["row", row - 1], ["rowOff", 0]] as const) {
      const c = doc.createElementNS(XDR, `xdr:${tag}`);
      c.textContent = String(v);
      el.appendChild(c);
    }
    return el;
  };
  anchor.appendChild(corner("from", spec.at.c1, spec.at.r1));
  anchor.appendChild(corner("to", spec.at.c2, spec.at.r2));
  pr.appendChild(anchor);
  control.appendChild(pr);
  choice.appendChild(control);
  alt.appendChild(choice);
  controls.appendChild(alt);
  sheet.layoutDirty = true;

  const model: SheetControl = {
    kind: spec.kind,
    name,
    label: spec.label,
    linkedCell: absLink,
    sourceRange: spec.sourceRange,
    checked: spec.kind === "checkbox" || spec.kind === "radio" ? false : undefined,
    selected: spec.kind === "dropdown" || spec.kind === "list" ? 0 : undefined,
    value: spec.kind === "spin" || spec.kind === "scroll" ? 0 : undefined,
    min: spec.kind === "spin" || spec.kind === "scroll" ? 0 : undefined,
    max: spec.kind === "spin" || spec.kind === "scroll" ? 100 : undefined,
    inc: spec.kind === "spin" || spec.kind === "scroll" ? 1 : undefined,
    anchor: { fromCol: spec.at.c1, fromRow: spec.at.r1, fromColOff: 0, fromRowOff: 0, toCol: spec.at.c2, toRow: spec.at.r2, toColOff: 0, toRowOff: 0 },
    propsPath,
    vmlPath,
    shapeId: String(shapeId),
  };
  (sheet.controls ??= []).push(model);
  return model;
}

/** Remove a control: its worksheet entry, its props part, and its VML shape. */
export function deleteXlsxControl(wb: Workbook, sheet: Sheet, control: SheetControl): void {
  const ws = sheet.doc?.documentElement;
  if (ws) {
    for (const el of Array.from(ws.getElementsByTagName("*"))) {
      if (el.localName !== "control" || el.getAttribute("shapeId") !== control.shapeId) continue;
      // Take the mc:AlternateContent wrapper with it, or an empty shell is left behind. closest()
      // cannot match a namespaced element by local name, so walk the ancestors.
      let drop: Element = el;
      for (let p = el.parentElement; p; p = p.parentElement) {
        if (p.localName === "controls") break; // the wrapper, if any, sits below <controls>
        if (p.localName === "AlternateContent" || p.localName === "Choice") drop = p;
      }
      drop.parentNode?.removeChild(drop);
      break;
    }
    sheet.layoutDirty = true;
  }
  if (control.propsPath) delete wb.files[control.propsPath];
  if (control.vmlPath && wb.files[control.vmlPath]) {
    const doc = parseXmlOpt(wb.files[control.vmlPath]);
    if (doc) {
      for (const shape of Array.from(doc.getElementsByTagName("*"))) {
        if (shape.localName !== "shape") continue;
        const id = (/(\d+)$/.exec(shape.getAttribute("id") ?? "") ?? [])[1];
        if (id === control.shapeId) { shape.parentNode?.removeChild(shape); break; }
      }
      wb.files[control.vmlPath] = serializeXml(doc);
    }
  }
  sheet.controls = (sheet.controls ?? []).filter((c) => c !== control);
}

/** The A1 reference a linked-cell input should store, normalised to absolute form. */
export const absoluteRef = (ref: string): string | undefined => {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref.trim());
  return m ? `$${m[1]!.toUpperCase()}$${m[2]}` : undefined;
};

/** The A1 range a source-range input should store ("D1:D3" -> "$D$1:$D$3"). */
export function absoluteRange(ref: string): string | undefined {
  const parts = ref.trim().split(":");
  if (parts.length > 2) return undefined;
  const from = parts[0] ? absoluteRef(parts[0]) : undefined;
  if (!from) return undefined;
  // A trailing colon ("D1:") is a half-typed range, not a single cell.
  if (parts.length === 1) return from;
  const to = parts[1] ? absoluteRef(parts[1]) : undefined;
  return to ? `${from}:${to}` : undefined;
}

/**
 * Rewrite the `<anchor>` inside a worksheet `<control>`, which is where a modern file keeps a
 * control's placement. A file that has none (older Excel wrote the VML alone) simply has nothing
 * here to update, and the VML anchor carries it.
 */
function updateWorksheetAnchor(sheet: Sheet, control: SheetControl): void {
  const doc = sheet.doc;
  const a = control.anchor;
  if (!doc || !a || !control.shapeId) return;
  for (const el of Array.from(doc.getElementsByTagName("*"))) {
    if (el.localName !== "control" || el.getAttribute("shapeId") !== control.shapeId) continue;
    const anchorEl = Array.from(el.getElementsByTagName("*")).find((e) => e.localName === "anchor");
    if (!anchorEl) return;
    const corner = (tag: "from" | "to", col: number, colOff: number, row: number, rowOff: number): void => {
      const c = Array.from(anchorEl.children).find((e) => e.localName === tag);
      if (!c) return;
      const put = (local: string, v: number): void => {
        const t = Array.from(c.children).find((e) => e.localName === local);
        if (t) t.textContent = String(v);
      };
      put("col", col - 1);
      put("colOff", colOff);
      put("row", row - 1);
      put("rowOff", rowOff);
    };
    corner("from", a.fromCol, a.fromColOff, a.fromRow, a.fromRowOff);
    corner("to", a.toCol, a.toColOff, a.toRow, a.toRowOff);
    sheet.layoutDirty = true;
    return;
  }
}

/** Rewrite a control's linked cell, source range, macro and placement in its parts. */
export function updateXlsxControlLinks(wb: Workbook, control: SheetControl, sheet?: Sheet): void {
  if (sheet) updateWorksheetAnchor(sheet, control);
  if (control.propsPath && wb.files[control.propsPath]) {
    const doc = parseXmlOpt(wb.files[control.propsPath]);
    const el = doc?.documentElement;
    if (doc && el) {
      if (control.linkedCell) el.setAttribute("fmlaLink", control.linkedCell);
      else el.removeAttribute("fmlaLink");
      if (control.sourceRange) el.setAttribute("fmlaRange", control.sourceRange);
      else el.removeAttribute("fmlaRange");
      wb.files[control.propsPath] = serializeXml(doc);
    }
  }
  if (!control.vmlPath || !wb.files[control.vmlPath]) return;
  const doc = parseXmlOpt(wb.files[control.vmlPath]);
  if (!doc) return;
  for (const shape of Array.from(doc.getElementsByTagName("*"))) {
    if (shape.localName !== "shape") continue;
    if ((/(\d+)$/.exec(shape.getAttribute("id") ?? "") ?? [])[1] !== control.shapeId) continue;
    const client = Array.from(shape.children).find((e) => e.localName === "ClientData");
    if (!client) break;
    const set = (local: string, value: string | undefined): void => {
      const existing = Array.from(client.children).find((e) => e.localName === local);
      if (!value) { if (existing) client.removeChild(existing); return; }
      if (existing) { existing.textContent = value; return; }
      const prefix = client.prefix ? `${client.prefix}:` : "";
      const el = doc.createElementNS(client.namespaceURI, `${prefix}${local}`);
      el.textContent = value;
      client.appendChild(el);
    };
    set("FmlaLink", control.linkedCell);
    set("FmlaRange", control.sourceRange);
    // The cell anchor, in the same from/to form createXlsxControl writes.
    const a = control.anchor;
    if (a) set("Anchor", `${a.fromCol - 1},0,${a.fromRow - 1},0,${a.toCol - 1},0,${a.toRow - 1},0`);
    // Excel qualifies the macro with its workbook; "[0]!" means this one, which is what it writes
    // for a macro living in the same file.
    set("FmlaMacro", control.macro ? `[0]!${control.macro}` : undefined);
    break;
  }
  wb.files[control.vmlPath] = serializeXml(doc);
}

/** A default placement for a new control: two columns wide on the selected row. */
export const placementFor = (r: number, c: number): NewControl["at"] => ({ r1: r, c1: c, r2: r + 1, c2: c + 2 });

/** Used by the UI to name the cell a control defaults to linking. */
export const defaultLink = (r: number, c: number): string => `$${colToLetters(c)}$${r}`;
