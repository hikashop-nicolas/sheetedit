// Form controls and ActiveX controls in a .xls ([MS-XLS] 2.4.181 Obj): an OBJ record whose
// subrecords hold the state, the linked cell and the macro, an OfficeArt shape for the place, and
// a TXO for the caption. ActiveX controls keep their persisted data in the Ctls stream.
import type { ChartAnchor } from "../../core/chart-model";
import type { SheetControl } from "../../core/model";
import { applyActiveXBin } from "../xlsx/control-read";
import { Bytes } from "./bytes";
import type { Anchor, Blip, DrawNode } from "./escher";
import { xlsSheetPx } from "./to-xlsx";
import { anchorPx, runFont, type XlsRunFont } from "./drawing";
import type { XlsBook, XlsControl, XlsCtlObj, XlsSheet } from "./types";

/** OBJ types that are form controls, as the model's kinds. */
export const CTL_KINDS: Record<number, SheetControl["kind"]> = {
  7: "button", 11: "checkbox", 12: "radio", 13: "textbox", 14: "label", 15: "groupBox",
  16: "spin", 17: "scroll", 18: "list", 19: "groupBox", 20: "dropdown",
};
const KIND_OT: Partial<Record<SheetControl["kind"], number>> = {
  button: 7, checkbox: 11, radio: 12, textbox: 13, label: 14, groupBox: 19, spin: 16, scroll: 17, list: 18, dropdown: 20,
};
const BASE_NAMES: Record<number, string> = {
  7: "Button", 11: "Check Box", 12: "Option Button", 13: "Edit Box", 14: "Label", 15: "Dialog", 16: "Spinner",
  17: "Scroll Bar", 18: "List Box", 19: "Group Box", 20: "Drop Down",
};

/** What a .xls control keeps from the file, for an exact write-back while it is untouched. */
export interface XlsControlExtra {
  /** The shape container's bytes (header included). */
  raw: Uint8Array;
  anchor: ChartAnchor;
  ot: number;
  id: number;
  subs: { ft: number; data: Uint8Array }[];
  /** The state as read, to tell what the user changed. */
  read: { link?: string; range?: string; macro?: string; checked?: boolean };
  /** The caption with its formatting runs, kept while the caption is unchanged. */
  text?: { text: string; runs: { ich: number; font: XlsRunFont }[] };
  txoGrbit?: number;
  /** Read from inside a group: the cell anchor (ClientAnchor bytes) it is written with instead of
      its child anchor while it has not moved. */
  grouped?: Uint8Array;
  /** ActiveX: the class name and the preview picture the shape shows. */
  progId?: string;
  blip?: Blip;
}

const u16 = (d: Uint8Array, o: number): number => (d[o] ?? 0) | ((d[o + 1] ?? 0) << 8);
const i16 = (d: Uint8Array, o: number): number => (u16(d, o) << 16) >> 16;
const u32 = (d: Uint8Array, o: number): number => (u16(d, o) | (u16(d, o + 2) << 16)) >>> 0;

/** The tokens of an ObjectParsedFormula: cce, 4 unused bytes, then the rgce. */
const fmlaRgce = (d: Uint8Array): Uint8Array => d.subarray(6, 6 + (u16(d, 0) & 0x7fff));

/** How long an ftLbsData runs: its own size field is not to be trusted (Excel writes 0x1FEE). */
function lbsLength(d: Uint8Array, at: number, ot: number): number {
  let q = at;
  const cbFmla = u16(d, q);
  q += 2 + cbFmla;
  const lines = u16(d, q), flags = u16(d, q + 4);
  q += 8;
  if (ot === 20) {
    q += 6;
    const cch = u16(d, q), size = 3 + cch * (d[q + 2]! & 1 ? 2 : 1);
    q += size + (size & 1);
  }
  if (flags & 2) for (let i = 0; i < lines && q + 3 <= d.length; i++) q += 3 + u16(d, q) * (d[q + 2]! & 1 ? 2 : 1);
  if ((flags >> 4) & 3) q += lines;
  return Math.min(q, d.length) - at;
}

/** Read an OBJ record as a control, or undefined when it is something else. */
export function readCtlObj(data: Uint8Array, decode: (rgce: Uint8Array) => string | null): XlsCtlObj | undefined {
  if (data.length < 22 || u16(data, 0) !== 0x15) return undefined;
  const ot = u16(data, 4), id = u16(data, 6), flags = u16(data, 8);
  // Autofilter and validation arrows are drop-downs marked as UI objects.
  if (!(ot in CTL_KINDS) && ot !== 8) return undefined;
  if (ot === 20 && flags & 0x100) return undefined;
  const subs: { ft: number; data: Uint8Array }[] = [];
  for (let p = 0; p + 4 <= data.length;) {
    const ft = u16(data, p), cb = u16(data, p + 2);
    if (ft === 0) break;
    const len = ft === 0x13 ? lbsLength(data, p + 4, ot) : Math.min(cb, data.length - p - 4);
    subs.push({ ft, data: data.subarray(p + 4, p + 4 + len).slice() });
    p += 4 + len;
  }
  const out: XlsCtlObj = { ot, id, subs };
  const fm = (d: Uint8Array): string | undefined => { const r = fmlaRgce(d); return r.length ? decode(r) ?? undefined : undefined; };
  for (const { ft, data: d } of subs) {
    if (ft === 0x12 && d.length >= 2) out.checked = u16(d, 0);
    else if (ft === 0x0c && d.length >= 16) out.sbs = { val: i16(d, 4), min: i16(d, 6), max: i16(d, 8), inc: u16(d, 10), page: u16(d, 12), horiz: u16(d, 14) !== 0 };
    else if (ft === 0x04) out.macro = fm(d);
    else if (ft === 0x0e || ft === 0x14) out.link = fm(d);
    else if (ft === 0x13 && d.length >= 2) {
      const cb = u16(d, 0);
      if (cb) out.range = fm(d.subarray(2, 2 + cb));
      out.sel = u16(d, 2 + cb + 2);
    }
  }
  if (ot !== 8) return out;
  // A picture is an ActiveX control when its data lives in the Ctls stream.
  const pio = subs.find((s) => s.ft === 0x08)?.data;
  const pict = subs.find((s) => s.ft === 0x09)?.data;
  if (!pio || !pict || !(u16(pio, 0) & 0x10)) return undefined;
  const ax = readPictFmla(pict, fm);
  if (!ax) return undefined;
  out.activeX = { progId: ax.progId, pos: ax.pos, size: ax.size };
  if (ax.link) out.link = ax.link;
  if (ax.range) out.range = ax.range;
  return out;
}

interface PictFmla { progId: string; pos: number; size: number; link?: string; range?: string; keyAt?: number }

/** An ftPictFmla: the class name, the Ctls stream slice and the linked cell / list range. */
function readPictFmla(d: Uint8Array, fm: (d: Uint8Array) => string | undefined): PictFmla | undefined {
  const cb = u16(d, 0);
  if (d.length < 10 + cb) return undefined;
  const f = d.subarray(2, 2 + cb);
  const cce = u16(f, 0) & 0x7fff;
  let progId = "";
  const e = 6 + cce;
  if (f[e] === 0x03) {
    const n = u16(f, e + 1), wide = f[e + 3]! & 1;
    for (let i = 0; i < n; i++) progId += String.fromCharCode(wide ? u16(f, e + 4 + i * 2) : f[e + 4 + i]!);
  }
  const out: PictFmla = { progId, pos: u32(d, 2 + cb), size: u32(d, 6 + cb) };
  const keyAt = 10 + cb;
  if (d.length >= keyAt + 4) {
    out.keyAt = keyAt;
    let q = keyAt + 4 + u32(d, keyAt);
    const lc = u16(d, q);
    if (lc) out.link = fm(d.subarray(q + 2, q + 2 + lc));
    q += 2 + lc;
    const rc = u16(d, q);
    if (rc) out.range = fm(d.subarray(q + 2, q + 2 + rc));
  }
  return out;
}

/** The OfficeArt shape's name (wzName), as Excel shows it in the name box. */
function shapeName(n: DrawNode): string | undefined {
  const c = n.props.get(0x380)?.complex;
  if (!c) return undefined;
  let s = "";
  for (let i = 0; i + 1 < c.length; i += 2) { const ch = u16(c, i); if (!ch) break; s += String.fromCharCode(ch); }
  return s || undefined;
}

/** A macro reference without its workbook qualifier, as the model names macros. */
export const bareMacro = (m: string): string => m.replace(/^\[\d+\]!/, "").replace(/^'(.*)'$/, "$1").replace(/^.*!/, "");

type Box = { l: number; t: number; r: number; b: number };

/** Where a grouped shape sits on the sheet: its box mapped out through every enclosing group. */
function groupedAnchor(node: DrawNode, groups: DrawNode[], s: XlsSheet): Anchor | undefined {
  const { colPx, rowPx } = xlsSheetPx(s);
  const sum = (n: number, size: (i: number) => number): number => { let v = 0; for (let i = 0; i < n; i++) v += size(i); return v; };
  const abs = (a: Anchor): Box => ({
    l: sum(a.col1, colPx) + (Math.min(1024, a.dx1) / 1024) * colPx(a.col1), t: sum(a.row1, rowPx) + (Math.min(256, a.dy1) / 256) * rowPx(a.row1),
    r: sum(a.col2, colPx) + (Math.min(1024, a.dx2) / 1024) * colPx(a.col2), b: sum(a.row2, rowPx) + (Math.min(256, a.dy2) / 256) * rowPx(a.row2),
  });
  let box: Box | undefined = node.child;
  for (let i = groups.length - 1; i >= 0 && box; i--) {
    const g = groups[i]!, sp = g.space;
    const to = i > 0 ? g.child : g.anchor ? abs(g.anchor) : undefined;
    if (!sp || !to || sp.r === sp.l || sp.b === sp.t) return undefined;
    const fx = (to.r - to.l) / (sp.r - sp.l), fy = (to.b - to.t) / (sp.b - sp.t);
    box = { l: to.l + (box.l - sp.l) * fx, r: to.l + (box.r - sp.l) * fx, t: to.t + (box.t - sp.t) * fy, b: to.t + (box.b - sp.t) * fy };
  }
  if (!box) return undefined;
  // Back to a cell and a fraction of it (1/1024 of a column, 1/256 of a row).
  const cell = (px: number, size: (i: number) => number, max: number, unit: number): [number, number] => {
    let i = 0, x = Math.max(0, px);
    while (i < max - 1 && x >= size(i)) { x -= size(i); i++; }
    return [i, Math.min(unit - 1, Math.round((x / (size(i) || 1)) * unit))];
  };
  const [col1, dx1] = cell(box.l, colPx, 256, 1024), [col2, dx2] = cell(box.r, colPx, 256, 1024);
  const [row1, dy1] = cell(box.t, rowPx, 65536, 256), [row2, dy2] = cell(box.b, rowPx, 65536, 256);
  return { col1, dx1, row1, dy1, col2, dx2, row2, dy2 };
}

/** A grouped shape's container as a top-level one: the cell anchor in place of the child anchor. */
export function ungroupedShape(raw: Uint8Array, anchor: Uint8Array): Uint8Array {
  const out = new Bytes();
  for (let p = 8; p + 8 <= raw.length;) {
    const type = u16(raw, p + 2), len = u32(raw, p + 4);
    const rec = raw.subarray(p, p + 8 + len);
    if (type === 0xf00f) out.u16(0).u16(0xf010).u32(anchor.length).bytes(anchor);
    else if (type === 0xf00a && len >= 8) {
      const c = rec.slice();
      const flags = (u32(c, 12) & ~0x2) | 0x200; // no longer a child; has an anchor
      new DataView(c.buffer).setUint32(12, flags >>> 0, true);
      out.bytes(c);
    } else out.bytes(rec);
    p += 8 + len;
  }
  return new Bytes().bytes(raw.subarray(0, 4)).u32(out.len).bytes(out.out()).out().slice();
}

/** A control read from the file, in the model the xlsx path and the editor share. */
export function xlsControlModel(c: XlsControl, s: XlsSheet, book: XlsBook, files: Record<string, Uint8Array>, binPath: string): SheetControl | null {
  const { node, obj, txo } = c;
  const cellAnchor = c.groups?.length ? groupedAnchor(node, c.groups, s) : node.anchor;
  if (!cellAnchor) return null;
  const anchor = anchorPx(cellAnchor, s);
  const ctl: SheetControl = { kind: CTL_KINDS[obj.ot] ?? "label", name: shapeName(node) ?? `${BASE_NAMES[obj.ot] ?? "Control"} ${obj.id}`, anchor };
  const extra: XlsControlExtra = { raw: node.raw, anchor, ot: obj.ot, id: obj.id, subs: obj.subs, read: { link: obj.link, range: obj.range, macro: obj.macro } };
  if (c.groups?.length) {
    const a = cellAnchor;
    extra.grouped = new Bytes().u16(0).u16(a.col1).u16(a.dx1).u16(a.row1).u16(a.dy1).u16(a.col2).u16(a.dx2).u16(a.row2).u16(a.dy2).out().slice();
  }
  if (obj.activeX) {
    const data = obj.activeX.data;
    if (!data?.length) return null;
    files[binPath] = data;
    applyActiveXBin(ctl, binPath, data);
    extra.progId = obj.activeX.progId;
    if (c.blip) extra.blip = c.blip;
  } else {
    if (txo?.text) {
      ctl.label = txo.text;
      extra.txoGrbit = txo.grbit;
      const runs = txo.runs.map((r) => ({ ich: r.ich, font: runFont(book.fonts[r.ifnt < 4 ? r.ifnt : r.ifnt - 1], book) })).filter((r): r is { ich: number; font: XlsRunFont } => !!r.font);
      if (runs.length) extra.text = { text: txo.text, runs };
    } else if (txo || CAPTION_GRBIT[ctl.kind] != null) { if (txo) extra.txoGrbit = txo.grbit; ctl.label = ""; } // no caption shows no text, not the name
    if (obj.macro) ctl.macro = bareMacro(obj.macro);
    if (ctl.kind === "checkbox" || ctl.kind === "radio") { ctl.checked = obj.checked === 1; extra.read.checked = ctl.checked; }
    if ((ctl.kind === "spin" || ctl.kind === "scroll") && obj.sbs) Object.assign(ctl, { value: obj.sbs.val, min: obj.sbs.min, max: obj.sbs.max, inc: obj.sbs.inc });
    if (ctl.kind === "dropdown" || ctl.kind === "list") ctl.selected = obj.sel ?? 0;
  }
  if (obj.link) ctl.linkedCell = obj.link;
  if (obj.range) ctl.sourceRange = obj.range;
  ctl.xls = extra;
  return ctl;
}

// ---- write -------------------------------------------------------------------------------------

/** What writing a control's OBJ needs from the workbook writer. */
export interface CtlWriteEnv {
  /** A cell or range reference as reference-class tokens, or null when it cannot be compiled. */
  ref(text: string): Uint8Array | null;
  /** The tokens that call a macro by name. */
  macro(text: string): Uint8Array | null;
  /** The new object id of a control by its id in the file (option button groups chain by id). */
  newId(oldId: number): number | undefined;
}

/** Subrecord order the format lays down; anything else goes last, before ftEnd. */
const ORDER = [0x15, 0x06, 0x07, 0x08, 0x0a, 0x0b, 0x0c, 0x0d, 0x04, 0x09, 0x14, 0x0e, 0x12, 0x11, 0x10, 0x13, 0x0f];

/** An ObjFmla's payload (cce, unused, rgce), padded to an even length. */
function objFmla(rgce: Uint8Array): Uint8Array {
  const b = new Bytes().u16(rgce.length).u32(0).bytes(rgce);
  if (b.len & 1) b.u8(0);
  return b.out().slice();
}

/** Single-token references stay valid as written; 3-D ones index the rebuilt EXTERNSHEET. */
const plainRef = (d: Uint8Array): boolean => { const r = fmlaRgce(d); return r.length === 5 || r.length === 9 ? [0x24, 0x25, 0x44, 0x45, 0x64, 0x65].includes(r[0]!) : false; };

const clamp16 = (v: number | undefined, dflt: number): number => Math.max(-32768, Math.min(32767, Math.round(v ?? dflt)));

/** How many rows a source range spans, for a new list's line count. */
function rangeRows(ref: string | undefined): number {
  const m = /\$?[A-Z]+\$?(\d+)(?::\$?[A-Z]+\$?(\d+))?$/i.exec(ref ?? "");
  return m ? Math.max(1, Number(m[2] ?? m[1]) - Number(m[1]) + 1) : 0;
}

/** The OBJ record payload for a form control, ftEnd included. */
export function ctlObjBody(ctl: SheetControl, id: number, env: CtlWriteEnv): Bytes {
  const x = ctl.xls;
  const ot = x?.ot ?? KIND_OT[ctl.kind] ?? 14;
  const old = new Map((x?.subs ?? []).map((s) => [s.ft, s.data]));
  const subs = new Map<number, Uint8Array>();
  const keep = (ft: number): void => { const d = old.get(ft); if (d) subs.set(ft, d); };
  for (const s of x?.subs ?? []) subs.set(s.ft, s.data);
  // ftCmo: the type and the new id; Excel's flags kept, else locked + printed (+ autoline for captions).
  const cmo = (old.get(0x15) ?? new Bytes().u16(ot).u16(0).u16(ot === 7 || ot === 14 || ot === 19 ? 0x4011 : 0x0011).zeros(12).out()).slice();
  cmo[0] = ot & 255; cmo[1] = ot >> 8; cmo[2] = id & 255; cmo[3] = (id >> 8) & 255;
  subs.set(0x15, cmo);
  // A reference subrecord: kept as read while it says the same and holds no sheet index.
  const fmla = (ft: number, text: string | undefined, was: string | undefined, compile: (t: string) => Uint8Array | null): void => {
    const d = old.get(ft);
    if (!text) { subs.delete(ft); return; }
    if (d && text === was && plainRef(d)) return keep(ft);
    const rgce = compile(text);
    if (rgce) subs.set(ft, objFmla(rgce));
    else if (d) keep(ft); else subs.delete(ft);
  };
  const macroText = ctl.macro ? (x?.read.macro && bareMacro(x.read.macro) === ctl.macro ? x.read.macro : ctl.macro) : undefined;
  fmla(0x04, macroText, x?.read.macro, (t) => env.macro(t));
  const k = ctl.kind;
  if (k === "checkbox" || k === "radio") {
    fmla(0x14, ctl.linkedCell, x?.read.link, (t) => env.ref(t));
    const d = (old.get(0x12) ?? new Bytes().u16(0).u16(0).u16(0).u16(1).out()).slice();
    // A mixed (grey) box stays mixed until the user changes it.
    const v = x?.read.checked === ctl.checked && u16(d, 0) === 2 ? 2 : ctl.checked ? 1 : 0;
    d[0] = v; d[1] = 0;
    subs.set(0x12, d);
    // ftCbls states it again (LibreOffice reads it from there).
    const cbls = (old.get(0x0a) ?? new Bytes().u16(0).zeros(8).u16(1).out()).slice();
    if (cbls.length >= 2) { cbls[0] = v; cbls[1] = 0; }
    subs.set(0x0a, cbls);
    if (k === "radio") {
      if (!subs.has(0x0b)) subs.set(0x0b, new Bytes().u32(0).u16(0).out().slice());
      const r = (old.get(0x11) ?? new Bytes().u16(0).u16(0).out()).slice();
      const next = env.newId(u16(r, 0)) ?? 0;
      r[0] = next & 255; r[1] = next >> 8;
      subs.set(0x11, r);
    }
  }
  if (k === "spin" || k === "scroll" || k === "list" || k === "dropdown") {
    const tall = !ctl.anchor || (ctl.anchor.toRow - ctl.anchor.fromRow) >= (ctl.anchor.toCol - ctl.anchor.fromCol);
    const d = (old.get(0x0c) ?? new Bytes().u32(0).u16(0).u16(0).u16(k === "list" || k === "dropdown" ? 0 : 100).u16(1).u16(10).u16(k === "scroll" && !tall ? 1 : 0).u16(16).u16(1).out()).slice();
    if (k === "spin" || k === "scroll") {
      const dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
      dv.setInt16(4, clamp16(ctl.value, 0), true);
      dv.setInt16(6, clamp16(ctl.min, 0), true);
      dv.setInt16(8, clamp16(ctl.max, 100), true);
      dv.setUint16(10, Math.max(0, clamp16(ctl.inc, 1)), true);
    }
    subs.set(0x0c, d);
    fmla(0x0e, ctl.linkedCell, x?.read.link, (t) => env.ref(t));
  }
  if (k === "list" || k === "dropdown") subs.set(0x13, lbsData(ctl, old.get(0x13), x?.read.range, ot, env));
  if (k === "textbox" && !subs.has(0x10)) subs.set(0x10, new Uint8Array(8));
  if (k === "groupBox" && ot === 19 && !subs.has(0x0f)) subs.set(0x0f, new Uint8Array(6));
  const fts = [...subs.keys()].sort((a, b) => (ORDER.indexOf(a) + 1 || 99) - (ORDER.indexOf(b) + 1 || 99));
  const out = new Bytes();
  for (const ft of fts) {
    const d = subs.get(ft)!;
    // ftLbsData's size field is cbFContinued, which Excel always writes as 0x1FEE.
    out.u16(ft).u16(ft === 0x13 ? 0x1fee : d.length).bytes(d);
  }
  return out.u16(0).u16(0);
}

/** ftLbsData: the input range and the selection in front, the rest as read. */
function lbsData(ctl: SheetControl, old: Uint8Array | undefined, wasRange: string | undefined, ot: number, env: CtlWriteEnv): Uint8Array {
  const cb = old ? u16(old, 0) : 0;
  const oldF = old?.subarray(2, 2 + cb);
  let f: Uint8Array = oldF ?? new Uint8Array(0);
  if (!ctl.sourceRange) f = new Uint8Array(0);
  else if (!(oldF && ctl.sourceRange === wasRange && plainRef(oldF))) {
    const rgce = env.ref(ctl.sourceRange);
    if (rgce) f = objFmla(rgce);
  }
  const lines = old && ctl.sourceRange === wasRange ? u16(old, 2 + cb) : rangeRows(ctl.sourceRange);
  const b = new Bytes().u16(f.length).bytes(f).u16(lines).u16(Math.max(0, Math.round(ctl.selected ?? 0)));
  if (old) return b.bytes(old.subarray(2 + cb + 4)).out().slice();
  b.u16(0x0008).u16(0); // fNo3d, no edit box
  if (ot === 20) b.u16(0).u16(8).u16(0).u16(0).u8(0).u8(0); // combo style, 8 lines, empty text, pad
  return b.out().slice();
}

/** The ActiveX OBJ: as read, with the new id and the control's slice of the new Ctls stream. */
export function activeXObjBody(ctl: SheetControl, id: number, pos: number, size: number, env: CtlWriteEnv): Bytes | null {
  const x = ctl.xls;
  const subs = x?.subs ?? [];
  const pict = subs.find((s) => s.ft === 0x09)?.data;
  if (!x || !pict) return null;
  const pf = readPictFmla(pict, () => undefined);
  if (!pf) return null;
  const cb = u16(pict, 0);
  const d = new Bytes().bytes(pict.subarray(0, 2 + cb)).u32(pos).u32(size);
  if (pf.keyAt != null) {
    const keyEnd = pf.keyAt + 4 + u32(pict, pf.keyAt);
    d.bytes(pict.subarray(pf.keyAt, keyEnd));
    let q = keyEnd;
    const part = (text: string | undefined, was: string | undefined): void => {
      const lc = u16(pict, q);
      const oldF = pict.subarray(q + 2, q + 2 + lc);
      q += 2 + lc;
      let f = text && lc && text === was && plainRef(oldF) ? oldF : undefined;
      if (!f && text) { const r = env.ref(text); if (r) f = objFmla(r); }
      d.u16(f?.length ?? 0);
      if (f) d.bytes(f);
    };
    part(ctl.linkedCell, x.read.link);
    part(ctl.sourceRange, x.read.range);
  }
  const out = new Bytes();
  for (const s of subs) {
    let data = s.data;
    if (s.ft === 0x15) { data = data.slice(); data[2] = id & 255; data[3] = (id >> 8) & 255; }
    if (s.ft === 0x09) data = d.out();
    out.u16(s.ft).u16(data.length).bytes(data);
  }
  return out.u16(0).u16(0);
}

/** A shape container for a control the file did not have: Excel's host-control shape (201). */
export function newControlShape(spid: number, anchor: Uint8Array, caption: boolean): Uint8Array {
  const kids = new Bytes();
  const rec = (type: number, ver: number, inst: number, data: Uint8Array): void => { kids.u16((ver & 0xf) | (inst << 4)).u16(type).u32(data.length).bytes(data); };
  rec(0xf00a, 2, 201, new Bytes().u32(spid).u32(0xa00).out());
  // Lock against grouping, fit text, no fill hit-test, a white fill and a hairline (Excel's own).
  const props: [number, number][] = [[0x007f, 0x01040104], [0x00bf, 0x00080008], [0x0181, 0x08000041], [0x0183, 0x08000040], [0x01bf, 0x00010000], [0x01c0, 0x08000040], [0x01ff, 0x00080000], [0x023f, 0x00020000], [0x03bf, 0x00080000]];
  const o = new Bytes();
  for (const [pid, v] of props) o.u16(pid).u32(v);
  rec(0xf00b, 3, props.length, o.out());
  rec(0xf010, 0, 0, anchor);
  rec(0xf011, 0, 0, new Uint8Array(0));
  if (caption) rec(0xf00d, 0, 0, new Uint8Array(0));
  const out = new Bytes().u16(0xf).u16(0xf004).u32(kids.len).bytes(kids.out());
  return out.out().slice();
}

/** A shape container with its picture (pib) pointing at a new BLIP index. */
export function withPib(raw: Uint8Array, pib: number): Uint8Array {
  const out = raw.slice();
  for (let p = 8; p + 8 <= out.length;) {
    const vi = u16(out, p), type = u16(out, p + 2), len = u32(out, p + 4);
    if (type === 0xf00b || type === 0xf122) {
      const n = vi >> 4;
      for (let i = 0; i < n; i++) {
        const at = p + 8 + i * 6;
        if ((u16(out, at) & 0x3fff) === 0x104) new DataView(out.buffer).setUint32(at + 2, pib, true);
      }
    }
    p += 8 + len;
  }
  return out;
}

/** The kinds that show a caption in a TXO, and its default alignment bits. */
export const CAPTION_GRBIT: Partial<Record<SheetControl["kind"], number>> = {
  button: 0x0224, checkbox: 0x0222, radio: 0x0222, label: 0x0212, groupBox: 0x0212,
};
