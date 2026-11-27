// The shape the HTML and Excel 2003 XML readers produce, before it becomes xlsx parts for the editor.
// Rows and columns are 0-based; colours are "#rrggbb".

export interface TFont {
  name?: string;
  size?: number; // points
  bold?: boolean;
  italic?: boolean;
  underline?: "single" | "double" | "singleAccounting" | "doubleAccounting";
  strike?: boolean;
  color?: string;
  vertAlign?: "superscript" | "subscript";
}

export interface TBorder { style: string; color?: string }

export interface TStyle {
  font?: TFont;
  fill?: { pattern: string; fg?: string; bg?: string };
  border?: { left?: TBorder; right?: TBorder; top?: TBorder; bottom?: TBorder; diagonal?: TBorder; diagUp?: boolean; diagDown?: boolean };
  align?: { h?: string; v?: string; wrap?: boolean; indent?: number; rotate?: number; shrink?: boolean };
  numFmt?: string;
  locked?: boolean;
  hidden?: boolean;
}

export interface TRun { text: string; font?: TFont }

export type TValue =
  | { t: "n"; v: number }
  | { t: "s"; v: string; runs?: TRun[] }
  | { t: "b"; v: boolean }
  | { t: "e"; v: string };

export interface TCell {
  r: number;
  c: number;
  v?: TValue;
  /** A1 formula text without the leading "=". */
  f?: string;
  /** The range an array formula spans (A1), on its top-left cell. */
  array?: string;
  style?: TStyle;
  link?: { href?: string; location?: string; tip?: string };
  comment?: { author?: string; text: string };
}

export interface TRange { r1: number; c1: number; r2: number; c2: number }

export interface TValidation {
  ranges: TRange[];
  type: "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom";
  op?: string;
  f1?: string;
  f2?: string;
  allowBlank?: boolean;
  noDropdown?: boolean;
  promptTitle?: string;
  prompt?: string;
  errorTitle?: string;
  error?: string;
  errorStyle?: "stop" | "warning" | "information";
}

export interface TCondRule { kind: "cellIs" | "expression"; op?: string; f1: string; f2?: string; style: TStyle }
export interface TCondFormat { ranges: TRange[]; rules: TCondRule[] }

export interface TPrint {
  landscape?: boolean;
  fitToPage?: boolean;
  fitWidth?: number;
  fitHeight?: number;
  scale?: number;
  paperSize?: number;
  margins?: { left?: number; right?: number; top?: number; bottom?: number; header?: number; footer?: number };
  header?: string;
  footer?: string;
  gridLines?: boolean;
  hCenter?: boolean;
  vCenter?: boolean;
}

export interface TSheet {
  name: string;
  cells: Map<number, TCell>;
  merges: TRange[];
  /** Column widths in px, hidden flags and whole-column styles. */
  cols: Map<number, { width?: number; hidden?: boolean; style?: TStyle }>;
  /** Row heights in points, hidden flags and whole-row styles. */
  rows: Map<number, { height?: number; hidden?: boolean; style?: TStyle }>;
  defaultColWidth?: number; // px
  defaultRowHeight?: number; // pt
  state?: "hidden" | "veryHidden";
  tabColor?: string;
  freeze?: { rows: number; cols: number; top?: number; left?: number };
  autoFilter?: TRange;
  validations: TValidation[];
  condFormats: TCondFormat[];
  print?: TPrint;
  hideGrid?: boolean;
  hideZeros?: boolean;
  rtl?: boolean;
}

export interface TName { name: string; ref: string; sheet?: number; hidden?: boolean }

export interface TBook {
  sheets: TSheet[];
  names: TName[];
  activeSheet?: number;
  defaultFont?: { name: string; size: number };
  date1904?: boolean;
}

export const cellKey = (r: number, c: number): number => r * 16384 + c;

export function newSheet(name: string): TSheet {
  return { name, cells: new Map(), merges: [], cols: new Map(), rows: new Map(), validations: [], condFormats: [] };
}

/** Make sheet names valid and unique (31 chars, no []:*?/\). */
export function uniqueSheetNames(names: string[], base: (n: number) => string): string[] {
  const used = new Set<string>();
  return names.map((raw, i) => {
    let name = raw.replace(/[[\]:*?/\\]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31) || base(i + 1);
    const stem = name;
    for (let k = 2; used.has(name.toLowerCase()); k++) name = `${stem.slice(0, 31 - String(k).length - 3)} (${k})`;
    used.add(name.toLowerCase());
    return name;
  });
}
