import { parseXmlOpt, type ControlVisuals, type SheetControl, type Workbook } from "../../core/model";
import { anchorOf, relMap, resolvePart } from "./chart-read";
import { kindOfClsid, readActiveXStream, type ActiveXControl, type ActiveXKind } from "./activex-read";

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
/** Whether a VML ClientData describes a form control at all: the same part also holds comments. */
const isControlType = (raw: string | null | undefined): boolean => (raw ?? "").toLowerCase() in KINDS;

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

/**
 * The macro a control runs, from `<x:FmlaMacro>`. Excel writes it qualified ("[0]!Button1_Click"
 * names the workbook), and a name with spaces arrives quoted; the model keeps the bare name.
 */
function macroName(client: Element): string | undefined {
  const raw = clientVal(client, "FmlaMacro");
  if (!raw) return undefined;
  const name = raw.replace(/^\[\d+\]!/, "").replace(/^'(.*)'$/, "$1").replace(/^.*!/, "").trim();
  return name || undefined;
}

/** Apply what the VML says, for a file whose control has no ctrlProps part. */
function applyClientData(ctl: SheetControl, client: Element): void {
  ctl.macro ??= macroName(client);
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
    const claimed = new Set<string>();
    // A control is written twice, once under mc:Choice with its placement and once under
    // mc:Fallback without it, so the same shape id appears twice and only the first is wanted.
    const seen = new Set<string>();
    for (const el of Array.from(sheet.doc.getElementsByTagName("*"))) {
      if (el.localName !== "control") continue;
      const dedupe = el.getAttribute("shapeId") ?? el.getAttribute("name") ?? "";
      if (dedupe && seen.has(dedupe)) continue;
      if (dedupe) seen.add(dedupe);
      const shapeId = el.getAttribute("shapeId") ?? undefined;
      const rid = Array.from(el.attributes).find((a) => a.localName === "id" && a.name !== "id")?.value ?? el.getAttribute("r:id") ?? undefined;
      const ctl: SheetControl = { kind: "label", name: el.getAttribute("name") ?? `Control ${controls.length + 1}`, shapeId, vmlPath };

      const target = rid ? rels.byId.get(rid) : undefined;
      const propsPath = target ? resolvePart("xl/worksheets", target) : undefined;
      // Form controls and ActiveX controls share this element and are told apart by the part the
      // relationship lands on. Reading an ActiveX part as a formControlPr yields a kind of "label"
      // with no properties, which is how a workbook full of ActiveX drew a screen of blank labels.
      const isActiveX = /\/activeX\//i.test(propsPath ?? "");
      const propsDoc = propsPath && files[propsPath] ? parseXmlOpt(files[propsPath]) : undefined;
      if (isActiveX) {
        applyActiveX(ctl, files, propsPath!, propsDoc);
      } else if (propsDoc?.documentElement) {
        ctl.propsPath = propsPath;
        applyProps(ctl, propsDoc.documentElement);
      }
      // <controlPr> carries the placement AND the Excel-side properties that are not the control's
      // own: which cell it drives and where a list gets its items. An ActiveX control keeps them
      // here rather than in its binary, so they read without touching it.
      const prEl = Array.from(el.getElementsByTagName("*")).find((e) => e.localName === "controlPr");
      const anchorEl = Array.from(el.getElementsByTagName("*")).find((e) => e.localName === "anchor");
      if (anchorEl) ctl.anchor = anchorOf(anchorEl) ?? undefined;
      if (prEl) {
        const link = prEl.getAttribute("linkedCell");
        if (link) ctl.linkedCell = link;
        const fill = prEl.getAttribute("listFillRange");
        if (fill) ctl.sourceRange = fill;
        // A macro named here wins over the one a button's name implies.
        const macro = prEl.getAttribute("macro");
        if (macro) ctl.macro = macro.replace(/^\[\d+\]!/, "").replace(/^.*!/, "");
      }

      const shape = shapeId ? vml.get(shapeId) : undefined;
      if (shape) {
        if (!propsDoc) ctl.kind = shape.kind;
        ctl.label ??= shape.label;
        applyClientData(ctl, shape.client);
        ctl.anchor ??= vmlAnchor(shape.client);
      }
      controls.push(ctl);
      if (shapeId) claimed.add(shapeId);
    }

    // Older Excel wrote no <controls> element at all: the button exists only as a VML shape, with
    // its macro, its label and its anchor. Those are picked up here, or such a file would show no
    // control at all. Only the known form-control ObjectTypes qualify, because the same VML part
    // also carries cell comments.
    for (const [id, shape] of vml) {
      if (claimed.has(id)) continue;
      if (!isControlType(shape.client.getAttribute("ObjectType"))) continue;
      const ctl: SheetControl = { kind: shape.kind, name: shape.label ?? `Control ${controls.length + 1}`, shapeId: id, vmlPath };
      ctl.label = shape.label;
      applyClientData(ctl, shape.client);
      ctl.anchor = vmlAnchor(shape.client);
      controls.push(ctl);
    }
    if (controls.length) sheet.controls = controls;
  }
}

/** How an ActiveX control's kind maps onto the model's own vocabulary. */
const ACTIVEX_KINDS: Partial<Record<ActiveXKind, SheetControl["kind"]>> = {
  commandButton: "button", checkbox: "checkbox", radio: "radio", textbox: "textbox",
  dropdown: "dropdown", list: "list", toggle: "toggle", label: "label",
  scroll: "scroll", spin: "spin", image: "image",
};

/**
 * An OLE_COLOR to CSS. Only a literal RGB converts: 0x80xxxxxx names a Windows system-palette
 * entry, whose colour is the desktop theme's rather than the document's, so it is left unset and
 * the grid's own colours apply instead of a guess at someone else's.
 */
function oleColorToCss(v: number | undefined): string | undefined {
  if (v === undefined) return undefined;
  if ((v & 0xff000000) !== 0) return undefined;  // system colour, or a palette index
  const hex = (n: number): string => (n & 0xff).toString(16).padStart(2, "0");
  return `#${hex(v)}${hex(v >> 8)}${hex(v >> 16)}`;  // OLE_COLOR is 0x00BBGGRR
}

/**
 * fmMousePointer to a CSS cursor. The two vocabularies line up almost exactly; 99 (custom) points
 * at a MouseIcon picture, which is a cursor image a page cannot install, so it keeps the default.
 */
const CURSORS: Record<number, string> = {
  1: "default", 2: "crosshair", 3: "text", 6: "nesw-resize", 7: "ns-resize", 8: "nwse-resize",
  9: "ew-resize", 10: "n-resize", 11: "wait", 12: "no-drop", 13: "progress", 14: "help", 15: "move",
};

/** PARAFORMAT_Alignment: 1 left, 2 right, 3 centre. */
const ALIGNS: Record<number, "left" | "right" | "center"> = { 1: "left", 2: "right", 3: "center" };

const dataUri = (mime: string, bytes: Uint8Array): string => {
  let binary = "";
  for (const b of bytes) binary += String.fromCharCode(b);
  return `data:${mime};base64,${btoa(binary)}`;
};

/** Everything the binary says about how the control looks, in the model's own vocabulary. */
function visualsOf(parsed: ActiveXControl): ControlVisuals | undefined {
  const v: ControlVisuals = {};
  if (parsed.enabled !== undefined) v.enabled = parsed.enabled;
  if (parsed.locked !== undefined) v.locked = parsed.locked;
  if (parsed.multiLine !== undefined) v.multiLine = parsed.multiLine;
  if (parsed.scrollBars !== undefined) v.scrollBars = parsed.scrollBars;
  if (parsed.maxLength) v.maxLength = parsed.maxLength;
  // PasswordChar is a character code; 0 means the box shows its text.
  if (parsed.passwordChar) v.passwordChar = String.fromCharCode(parsed.passwordChar);
  if (parsed.listRows) v.listRows = parsed.listRows;
  if (parsed.multiSelect) v.multiSelect = parsed.multiSelect !== 0;
  // DisplayStyle tells an editable combo (3) from a drop-list one (7); they share a class id.
  if (parsed.displayStyle !== undefined) v.editable = parsed.displayStyle !== 7;
  if (parsed.wordWrap !== undefined) v.wordWrap = parsed.wordWrap;
  if (parsed.captionLeft !== undefined) v.captionLeft = parsed.captionLeft;
  const fore = oleColorToCss(parsed.foreColor);
  const back = oleColorToCss(parsed.backColor);
  const border = oleColorToCss(parsed.borderColor);
  if (fore) v.color = fore;
  if (back) v.background = back;
  if (border) v.borderColor = border;
  if (parsed.borderStyle) v.borderStyle = parsed.borderStyle;
  if (parsed.specialEffect) v.specialEffect = parsed.specialEffect;
  if (parsed.orientation !== undefined) v.orientation = parsed.orientation;
  if (parsed.smallChange !== undefined) v.smallChange = parsed.smallChange;
  if (parsed.largeChange !== undefined) v.largeChange = parsed.largeChange;
  if (parsed.font) {
    const f = parsed.font;
    const font: NonNullable<ControlVisuals["font"]> = {};
    if (f.name) font.name = f.name;
    if (f.sizePt) font.sizePt = f.sizePt;
    if (f.bold || (f.weight ?? 0) >= 700) font.bold = true;
    if (f.italic) font.italic = true;
    if (f.underline) font.underline = true;
    if (f.strike) font.strike = true;
    const align = f.align !== undefined ? ALIGNS[f.align] : undefined;
    if (align) font.align = align;
    if (Object.keys(font).length) v.font = font;
  }
  if (parsed.picture) v.picture = dataUri(parsed.picture.mime, parsed.picture.bytes);
  const cursor = parsed.mousePointer !== undefined ? CURSORS[parsed.mousePointer] : undefined;
  if (cursor) v.cursor = cursor;
  if (parsed.accelerator) v.accelerator = String.fromCharCode(parsed.accelerator);
  return Object.keys(v).length ? v : undefined;
}

/**
 * Fill a control in from its ActiveX parts: the class id names the kind, and the persisted binary
 * beside it carries the caption and the value. What cannot be read leaves the control as a label
 * rather than as something it is not.
 */
function applyActiveX(ctl: SheetControl, files: Record<string, Uint8Array>, xmlPath: string, doc: Document | undefined): void {
  ctl.activeX = true;
  const clsid = doc?.documentElement?.getAttribute("ax:classid")
    ?? doc?.documentElement?.getAttribute("classid") ?? "";
  const kind = kindOfClsid(clsid);
  ctl.kind = ACTIVEX_KINDS[kind] ?? "label";
  // The binary is a separate part, reached through the activeX part's own relationships.
  const relsPath = xmlPath.replace(/([^/]+)$/, "_rels/$1.rels");
  const bin = relMap(files, relsPath).byType.find((r) => /activeXControlBinary/i.test(r.type));
  const binPath = bin ? resolvePart(xmlPath.replace(/\/[^/]+$/, ""), bin.target) : undefined;
  // An ActiveX button's handler lives in the sheet's own code module, named after the control:
  // CommandButton1 runs CommandButton1_Click. That is the convention, not a guess.
  if (ctl.kind === "button") ctl.macro ??= `${ctl.name}_Click`;
  const parsed = binPath && files[binPath] ? readActiveXStream(files[binPath]) : undefined;
  if (!parsed) return;
  ctl.activeXBinPath = binPath;
  if (parsed.caption) ctl.label = parsed.caption;
  // A range control's bounds come from its binary, where a form control keeps them in ctrlProps.
  if (parsed.min !== undefined) ctl.min = parsed.min;
  if (parsed.max !== undefined) ctl.max = parsed.max;
  if (parsed.position !== undefined) ctl.value = parsed.position;
  if (parsed.value !== undefined) {
    ctl.activeXValue = parsed.value;
    // A checkbox, an option button and a toggle all persist "0"/"1"; anything else is text.
    if (ctl.kind === "checkbox" || ctl.kind === "radio" || ctl.kind === "toggle") ctl.checked = parsed.value === "1";
  }
  const visuals = visualsOf(parsed);
  if (visuals) ctl.visuals = visuals;
}

/** Whether the workbook carries ActiveX controls, which are Windows COM and left untouched. */
export const hasActiveX = (files: Record<string, Uint8Array>): boolean =>
  Object.keys(files).some((p) => /^xl\/activeX\//i.test(p));
