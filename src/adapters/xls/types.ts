// The intermediate shape the BIFF reader produces, before it becomes xlsx parts for the editor.
import type { XlsString } from "./biff";
import type { Anchor, Blip } from "./escher";

export interface XlsFont {
  height: number; // twips
  italic: boolean;
  strike: boolean;
  weight: number;
  underline: number;
  escapement: number;
  color: number;
  name: string;
}

export interface XlsBorder { style: number; color: number }

export interface XlsXf {
  font: number;
  fmt: number;
  locked: boolean;
  hidden: boolean;
  isStyle: boolean;
  parent: number;
  hAlign: number;
  wrap: boolean;
  vAlign: number;
  rot: number;
  indent: number;
  shrink: boolean;
  left: XlsBorder;
  right: XlsBorder;
  top: XlsBorder;
  bottom: XlsBorder;
  diag?: { style: number; color: number; up: boolean; down: boolean };
  pattern: number;
  fore: number;
  back: number;
}

export type XlsValue = { k: "n"; v: number } | { k: "s"; v: string; runs?: XlsString["runs"] } | { k: "b"; v: boolean } | { k: "e"; v: string } | { k: "blank" };

export interface XlsCell {
  r: number; // 0-based
  c: number;
  xf: number;
  val: XlsValue;
  formula?: string;
  arrayRef?: string;
  /** Raw PtgExp target, resolved once the sheet's SHRFMLA / ARRAY records are known. */
  exp?: { r: number; c: number };
}

export interface XlsRow { height?: number; custom?: boolean; hidden?: boolean; level?: number; collapsed?: boolean; xf?: number }
export interface XlsCol { first: number; last: number; width: number; xf: number; hidden: boolean; level: number; collapsed: boolean }
export interface XlsRange { r1: number; c1: number; r2: number; c2: number } // 0-based inclusive

export interface XlsLink { range: XlsRange; url?: string; location?: string; tip?: string }
export interface XlsNote { r: number; c: number; author: string; text: string; obj: number }

export interface XlsDv {
  type: number;
  errStyle: number;
  strList: boolean;
  allowBlank: boolean;
  op: number;
  promptTitle: string;
  errorTitle: string;
  prompt: string;
  error: string;
  formula1?: string;
  formula2?: string;
  ranges: XlsRange[];
}

export interface XlsCf { type: number; op: number; formula1?: string; formula2?: string; fill?: number; color?: number; bold?: boolean; italic?: boolean; strike?: boolean; underline?: boolean }
export interface XlsCondFmt { ranges: XlsRange[]; rules: XlsCf[] }

export interface XlsImage { anchor: Anchor; blip: Blip; name?: string }

export interface XlsPrint {
  paperSize?: number;
  scale?: number;
  pageStart?: number;
  usePage?: boolean;
  fitWidth?: number;
  fitHeight?: number;
  portrait?: boolean;
  noOrient?: boolean;
  overThenDown?: boolean;
  noPls?: boolean;
  headerMargin?: number;
  footerMargin?: number;
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: string;
  footer?: string;
  gridLines?: boolean;
  headings?: boolean;
  hCenter?: boolean;
  vCenter?: boolean;
  fitToPage?: boolean;
  rowBreaks?: number[];
  colBreaks?: number[];
}

export interface XlsSheet {
  name: string;
  state: number; // 0 visible, 1 hidden, 2 very hidden
  cells: Map<number, XlsCell>;
  rows: Map<number, XlsRow>;
  cols: XlsCol[];
  defColWidth?: number; // characters
  stdWidth?: number; // 1/256 characters
  defRowHeight?: number; // twips
  merges: XlsRange[];
  links: XlsLink[];
  notes: XlsNote[];
  dvs: XlsDv[];
  condFmts: XlsCondFmt[];
  images: XlsImage[];
  window2?: number;
  pane?: { x: number; y: number; top: number; left: number };
  protect?: boolean;
  password?: number;
  objProtect?: boolean;
  scenProtect?: boolean;
  /** FEATHEADR enhanced protection: a set bit allows the action. */
  allow?: number;
  print: XlsPrint;
  tabColor?: number;
  summaryBelow?: boolean;
  summaryRight?: boolean;
  charts: number;
  shapes: number;
  /** Position in the BOUNDSHEET list (chart sheets included), which names and 3D refs index. */
  boundIndex?: number;
  hasAutoFilter?: boolean;
  autoFilter?: XlsRange;
}

export interface XlsName {
  name: string;
  /** 0 = workbook scope, else 1-based sheet index. */
  itab: number;
  hidden: boolean;
  builtin?: number;
  func: boolean;
  formula?: string;
  rgce?: Uint8Array;
  rgcb?: Uint8Array;
}

export interface XlsSupbook {
  kind: "self" | "addin" | "external" | "ole";
  path?: string;
  sheets: string[];
  names: string[];
  /** External books: the EXTERNNAME payloads and the XCT / CRN cached values, kept verbatim. */
  rawNames?: Uint8Array[];
  cache?: { type: number; data: Uint8Array }[];
}

export interface XlsBook {
  biff: number;
  codepage: number;
  date1904: boolean;
  fonts: XlsFont[];
  formats: Map<number, string>;
  xfs: XlsXf[];
  palette: string[]; // index 8.. -> "#rrggbb"
  sheets: XlsSheet[];
  names: XlsName[];
  supbooks: XlsSupbook[];
  xti: { book: number; first: number; last: number }[];
  protect?: boolean;
  windowProtect?: boolean;
  password?: number;
  activeTab: number;
  sst: XlsString[];
  /** Chart sheets and dialog / VB sheets that were not loaded. */
  skippedSheets: string[];
  encrypted?: "xor" | "rc4" | "cryptoapi";
  summary?: Map<string, Uint8Array>;
}
