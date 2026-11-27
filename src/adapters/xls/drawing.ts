// OfficeArt shapes ([MS-ODRAW]) in a .xls drawing, as sheetedit shapes and charts: the shape type
// (MSOSPT) mapped onto DrawingML preset names, the fill / line / text properties, the anchors.
import type { ChartAnchor, ChartModel } from "../../core/chart-model";
import type { SheetShape, ShapeGeom } from "../../core/model";
import { GEOM_PRESET } from "../../core/shape-geom";
import { geomOf } from "../xlsx/shape-read";
import type { Anchor, DrawNode } from "./escher";
import { icvCss } from "./palette";
import { xlsSheetPx } from "./to-xlsx";
import type { XlsBook, XlsFont, XlsSheet, XlsTxo } from "./types";

/** MSOSPT shape types and the DrawingML preset each one is. */
export const SPT_PRESET: Record<number, string> = {
  1: "rect", 2: "roundRect", 3: "ellipse", 4: "diamond", 5: "triangle", 6: "rtTriangle", 7: "parallelogram",
  8: "trapezoid", 9: "hexagon", 10: "octagon", 11: "plus", 12: "star5", 13: "rightArrow", 15: "homePlate",
  16: "cube", 19: "arc", 20: "line", 21: "plaque", 22: "can", 23: "donut", 32: "straightConnector1",
  33: "bentConnector2", 34: "bentConnector3", 35: "bentConnector4", 36: "bentConnector5",
  37: "curvedConnector2", 38: "curvedConnector3", 39: "curvedConnector4", 40: "curvedConnector5",
  41: "callout1", 42: "callout2", 43: "callout3", 44: "accentCallout1", 45: "accentCallout2", 46: "accentCallout3",
  47: "borderCallout1", 48: "borderCallout2", 49: "borderCallout3", 50: "accentBorderCallout1",
  51: "accentBorderCallout2", 52: "accentBorderCallout3", 53: "ribbon", 54: "ribbon2", 55: "chevron",
  56: "pentagon", 57: "noSmoking", 58: "star8", 59: "star16", 60: "star32", 61: "wedgeRectCallout",
  62: "wedgeRoundRectCallout", 63: "wedgeEllipseCallout", 64: "wave", 65: "foldedCorner", 66: "leftArrow",
  67: "downArrow", 68: "upArrow", 69: "leftRightArrow", 70: "upDownArrow", 71: "irregularSeal1",
  72: "irregularSeal2", 73: "lightningBolt", 74: "heart", 76: "quadArrow", 77: "leftArrowCallout",
  78: "rightArrowCallout", 79: "upArrowCallout", 80: "downArrowCallout", 81: "leftRightArrowCallout",
  82: "upDownArrowCallout", 83: "quadArrowCallout", 84: "bevel", 85: "leftBracket", 86: "rightBracket",
  87: "leftBrace", 88: "rightBrace", 89: "leftUpArrow", 90: "bentUpArrow", 91: "bentArrow", 92: "star24",
  93: "stripedRightArrow", 94: "notchedRightArrow", 95: "blockArc", 96: "smileyFace", 97: "verticalScroll",
  98: "horizontalScroll", 99: "circularArrow", 101: "uturnArrow", 102: "curvedRightArrow",
  103: "curvedLeftArrow", 104: "curvedUpArrow", 105: "curvedDownArrow", 106: "cloudCallout",
  107: "ellipseRibbon", 108: "ellipseRibbon2", 109: "flowChartProcess", 110: "flowChartDecision",
  111: "flowChartInputOutput", 112: "flowChartPredefinedProcess", 113: "flowChartInternalStorage",
  114: "flowChartDocument", 115: "flowChartMultidocument", 116: "flowChartTerminator",
  117: "flowChartPreparation", 118: "flowChartManualInput", 119: "flowChartManualOperation",
  120: "flowChartConnector", 121: "flowChartPunchedCard", 122: "flowChartPunchedTape",
  123: "flowChartSummingJunction", 124: "flowChartOr", 125: "flowChartCollate", 126: "flowChartSort",
  127: "flowChartExtract", 128: "flowChartMerge", 129: "flowChartOfflineStorage",
  130: "flowChartOnlineStorage", 131: "flowChartMagneticTape", 132: "flowChartMagneticDisk",
  133: "flowChartMagneticDrum", 134: "flowChartDisplay", 135: "flowChartDelay",
  176: "flowChartAlternateProcess", 177: "flowChartOffpageConnector", 182: "leftRightUpArrow", 183: "sun",
  184: "moon", 185: "bracketPair", 186: "bracePair", 187: "star4", 188: "doubleWave", 202: "rect",
};

const PRESET_SPT: Record<string, number> = {};
for (const [k, v] of Object.entries(SPT_PRESET)) if (!(v in PRESET_SPT)) PRESET_SPT[v] = Number(k);
PRESET_SPT.rect = 1;

/** The MSOSPT a preset (or a geometry) is written as; undefined when OfficeArt has no such shape. */
export function sptOf(preset: string | undefined, geom: ShapeGeom): number | undefined {
  if (preset && PRESET_SPT[preset] != null && geomOf(preset) === geom) return PRESET_SPT[preset];
  return PRESET_SPT[GEOM_PRESET[geom]];
}

/** Whether a geometry can be written to a .xls as itself rather than as its closest stand-in. */
export const xlsCanWrite = (geom: ShapeGeom): boolean => PRESET_SPT[GEOM_PRESET[geom]] != null;

// System colour indexes Excel uses in OfficeArt colours: window text / background, chart, tooltip.
const SYS_ICV: Record<number, string> = { 64: "#000000", 65: "#ffffff", 77: "#000000", 78: "#ffffff", 79: "#000000", 80: "#ffffe1", 81: "#000000" };

/** An OfficeArtCOLORREF as CSS: RGB, or a palette index when flagged. */
export function colorRef(v: number, palette: string[]): string | undefined {
  const flags = v >>> 24;
  if (flags & 0x08 || flags & 0x01) { const i = v & 0xffff; return icvCss(palette, i) ?? SYS_ICV[i]; }
  if (flags & 0x10) return undefined;
  const hex = (n: number): string => n.toString(16).padStart(2, "0");
  return "#" + hex(v & 0xff) + hex((v >> 8) & 0xff) + hex((v >> 16) & 0xff);
}

const LINE_ENDS = ["", "triangle", "stealth", "diamond", "oval", "arrow"];
const DASHING = ["", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot"];

/** A font a shape's text uses, by value, so it can be added to the written FONT table. */
export interface XlsRunFont { height: number; italic: boolean; strike: boolean; weight: number; underline: number; escapement: number; color?: string; name: string }

/** The .xls-only details of a shape read from a file, for an exact write-back while it is untouched. */
export interface XlsShapeExtra {
  /** The shape container's bytes (header included). */
  raw: Uint8Array;
  /** The anchor as read, to tell a moved shape apart. */
  anchor: ChartAnchor;
  /** The enclosing groups' own containers, outermost first, each with its member count. */
  groups?: { id: number; raw: Uint8Array }[];
  /** The text with its formatting runs, kept while the text is unchanged. */
  text?: { text: string; runs: { ich: number; font: XlsRunFont }[] };
  ot: number;
}

const fixed = (v: number): number => ((v | 0) >> 16) + (v & 0xffff) / 65536;

/** A cell anchor (0-based, offsets in 1/1024 column and 1/256 row) as a px anchor. */
export function anchorPx(a: Anchor, s: XlsSheet): ChartAnchor {
  const { colPx, rowPx } = xlsSheetPx(s);
  const off = (frac: number, size: number): number => Math.round(Math.min(1, frac) * size);
  return {
    fromCol: a.col1 + 1, fromColOff: off(a.dx1 / 1024, colPx(a.col1)), fromRow: a.row1 + 1, fromRowOff: off(a.dy1 / 256, rowPx(a.row1)),
    toCol: a.col2 + 1, toColOff: off(a.dx2 / 1024, colPx(a.col2)), toRow: a.row2 + 1, toRowOff: off(a.dy2 / 256, rowPx(a.row2)),
  };
}

export const runFont = (f: XlsFont | undefined, book: XlsBook): XlsRunFont | undefined =>
  f ? { height: f.height, italic: f.italic, strike: f.strike, weight: f.weight, underline: f.underline, escapement: f.escapement, color: icvCss(book.palette, f.color), name: f.name } : undefined;

/** The sheet's drawing shapes and text boxes, groups flattened onto their members. */
export function readXlsShapes(s: XlsSheet, book: XlsBook): SheetShape[] {
  const d = s.drawing;
  if (!d) return [];
  const out: SheetShape[] = [];
  let groupId = 0;
  const visit = (n: DrawNode, anchor: ChartAnchor | undefined, box: { x: number; y: number; w: number; h: number } | undefined, groups: { id: number; raw: Uint8Array }[]): void => {
    const own = n.anchor ? anchorPx(n.anchor, s) : anchor;
    if (!own) return;
    // A member's box in its group's space, as fractions of the outermost group's box.
    let within = box;
    if (n.child && box && groups.length) {
      const sp = spaceOf(groups[groups.length - 1]!.raw);
      if (sp && sp.r !== sp.l && sp.b !== sp.t) {
        const fx = (n.child.l - sp.l) / (sp.r - sp.l), fy = (n.child.t - sp.t) / (sp.b - sp.t);
        const fw = (n.child.r - n.child.l) / (sp.r - sp.l), fh = (n.child.b - n.child.t) / (sp.b - sp.t);
        within = { x: box.x + fx * box.w, y: box.y + fy * box.h, w: fw * box.w, h: fh * box.h };
      }
    }
    if (n.kids) {
      const inner = [...groups, { id: ++groupId, raw: n.raw }];
      for (const k of n.kids) visit(k, own, within ?? { x: 0, y: 0, w: 1, h: 1 }, inner);
      return;
    }
    const sh = toShape(n, own, d.txos[n.txo ?? -1], d.objs[n.obj ?? -1]?.ot ?? 0x1e, book);
    if (!sh) return;
    if (groups.length && within) sh.within = within;
    sh.xls = { raw: n.raw, anchor: own, ot: d.objs[n.obj ?? -1]?.ot ?? 0x1e, ...(groups.length ? { groups } : {}), ...(sh.xls?.text ? { text: sh.xls.text } : {}) };
    out.push(sh);
  };
  for (const n of d.nodes) visit(n, undefined, undefined, []);
  return out;
}

/** A group container's coordinate space (its FSPGR). */
function spaceOf(raw: Uint8Array): { l: number; t: number; r: number; b: number } | undefined {
  for (let p = 8; p + 8 <= raw.length;) {
    const type = raw[p + 2]! | (raw[p + 3]! << 8);
    const len = (raw[p + 4]! | (raw[p + 5]! << 8) | (raw[p + 6]! << 16) | (raw[p + 7]! << 24)) >>> 0;
    if (type === 0xf009 && len >= 16) {
      const dv = new DataView(raw.buffer, raw.byteOffset + p + 8, 16);
      return { l: dv.getInt32(0, true), t: dv.getInt32(4, true), r: dv.getInt32(8, true), b: dv.getInt32(12, true) };
    }
    p += 8 + len;
  }
  return undefined;
}

function toShape(n: DrawNode, anchor: ChartAnchor, txo: XlsTxo | undefined, ot: number, book: XlsBook): SheetShape | null {
  const p = (id: number): number | undefined => n.props.get(id)?.value;
  const has = (id: number, bit: number, use: number, dflt: boolean): boolean => {
    const v = p(id);
    return v != null && v & use ? !!(v & bit) : dflt;
  };
  const isLine = n.spt === 20 || n.spt === 32 || ot === 1;
  const preset = SPT_PRESET[n.spt] ?? (ot === 3 ? "ellipse" : ot === 4 ? "arc" : isLine ? "line" : "rect");
  const geom = geomOf(preset);
  const sh: SheetShape = { geom, preset, anchor };
  // Fill: filled unless switched off; gradients run from the fill colour to the back colour.
  if (!isLine && has(0x1bf, 0x10, 0x100000, true)) {
    const fill = colorRef(p(0x181) ?? 0xffffff, book.palette);
    if (fill) sh.fill = fill;
    const type = p(0x180) ?? 0;
    const back = colorRef(p(0x183) ?? 0xffffff, book.palette);
    if (fill && back && type >= 4 && type <= 8) sh.fillGradient = { angle: (90 - fixed(p(0x18b) ?? 0) + 360) % 360, stops: [{ pos: 0, color: fill }, { pos: 1, color: back }] };
    const op = p(0x182);
    if (op != null && op < 0x10000) sh.fillOpacity = Math.round((op / 65536) * 100) / 100;
  }
  if (has(0x1ff, 0x8, 0x80000, true)) {
    sh.stroke = colorRef(p(0x1c0) ?? 0, book.palette) ?? "#000000";
    const w = p(0x1cb);
    if (w != null) sh.strokeWidth = Math.max(1, Math.round(w / 9525));
    const dash = DASHING[p(0x1ce) ?? 0];
    if (dash) sh.dash = dash;
    const head = LINE_ENDS[p(0x1d0) ?? 0], tail = LINE_ENDS[p(0x1d1) ?? 0];
    if (head) sh.headEnd = head;
    if (tail) sh.tailEnd = tail;
  }
  if (n.flags & 0x40) sh.flipH = true;
  if (n.flags & 0x80) sh.flipV = true;
  const rot = p(0x04);
  if (rot) {
    const deg = ((fixed(rot) % 360) + 360) % 360;
    if (deg) sh.rotation = deg;
  }
  if (txo?.text) {
    sh.text = txo.text;
    const h = (txo.grbit >> 1) & 7, v = (txo.grbit >> 4) & 7;
    if (h === 1) sh.textAlign = "left"; else if (h === 2) sh.textAlign = "center"; else if (h === 3) sh.textAlign = "right";
    if (v === 1) sh.textValign = "top"; else if (v === 2) sh.textValign = "middle"; else if (v === 3) sh.textValign = "bottom";
    const runs = txo.runs.map((r) => ({ ich: r.ich, font: runFont(book.fonts[r.ifnt < 4 ? r.ifnt : r.ifnt - 1], book) })).filter((r): r is { ich: number; font: XlsRunFont } => !!r.font);
    const f0 = runs[0]?.font;
    if (f0) {
      sh.textSize = f0.height / 20;
      if (f0.color) sh.textColor = f0.color;
    }
    if (runs.length) sh.xls = { raw: n.raw, anchor, ot, text: { text: txo.text, runs } };
  }
  return sh;
}

/** A chart read from the file, anchored on its sheet. */
export function xlsChartModel(c: { anchor?: Anchor; chart: { model: Omit<ChartModel, "id" | "anchor"> } }, s: XlsSheet, id: string): ChartModel | null {
  if (!c.anchor) return null;
  return { ...c.chart.model, id, anchor: anchorPx(c.anchor, s) } as ChartModel;
}
