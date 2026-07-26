import type { Sheet } from "./model";

// ---------------------------------------------------------------------------
// Print settings
// ---------------------------------------------------------------------------
// sheetedit does not print, so these are carried, shown and authored for the benefit of whatever
// does: Excel, Calc, or a PDF export. The model is stated in xlsx's terms (inches, a paper-size id,
// scale as a percentage) because that is the more explicit of the two formats; the ods adapter
// converts to ODF's page-layout / master-page pair.
//
// Two things the grid can show without printing anything: the print area, and where the pages
// break. Both are drawn as an overlay, which is what makes these settings checkable at all.

/** A page-break line: the 1-based row / column that starts a new page. */
export type Breaks = number[];

/** The three regions of a header or footer, in the file's own field-code syntax (&P, &N, &D, ...). */
export interface HeaderFooter {
  left?: string;
  center?: string;
  right?: string;
}

export interface PrintSetup {
  orientation?: "portrait" | "landscape";
  /** xlsx paperSize id (9 = A4, 1 = Letter, ...). See PAPER_SIZES. */
  paperSize?: number;
  /** Zoom percentage (10..400). Ignored by both apps while a fit-to is in force. */
  scale?: number;
  /** Fit the printout to this many pages across / down; 0 = do not constrain that axis. */
  fitToWidth?: number;
  fitToHeight?: number;
  /** Whether fit-to wins over scale (xlsx sheetPr/pageSetUpPr@fitToPage). */
  fitToPage?: boolean;
  /** Which way multi-page sheets paginate. */
  pageOrder?: "downThenOver" | "overThenDown";
  /** Page number to print on the first page, when the file overrides the default 1. */
  firstPageNumber?: number;
  /** Page margins in inches. header/footer are the distance from the paper edge to that block. */
  margins?: { left: number; right: number; top: number; bottom: number; header: number; footer: number };
  /** Print the gridlines / the row and column headings. */
  gridLines?: boolean;
  headings?: boolean;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
  header?: HeaderFooter;
  footer?: HeaderFooter;
  /** The ranges to print (1-based, inclusive). Absent = the whole used range. */
  printArea?: { r1: number; c1: number; r2: number; c2: number }[];
  /** Rows / columns repeated at the top or left of every page (1-based, inclusive). */
  titleRows?: { from: number; to: number };
  titleCols?: { from: number; to: number };
  /** Manual page breaks: the 1-based line that starts a new page. */
  rowBreaks?: Breaks;
  colBreaks?: Breaks;
}

/** The paper sizes both formats agree on, as xlsx id -> millimetres (portrait). */
export const PAPER_SIZES: Record<number, { name: string; w: number; h: number }> = {
  1: { name: "Letter", w: 215.9, h: 279.4 },
  3: { name: "Tabloid", w: 279.4, h: 431.8 },
  5: { name: "Legal", w: 215.9, h: 355.6 },
  7: { name: "Executive", w: 184.15, h: 266.7 },
  8: { name: "A3", w: 297, h: 420 },
  9: { name: "A4", w: 210, h: 297 },
  11: { name: "A5", w: 148, h: 210 },
  12: { name: "B4", w: 250, h: 353 },
  13: { name: "B5", w: 176, h: 250 },
};

export const DEFAULT_PAPER = 9; // A4
export const MM_PER_INCH = 25.4;

/** Excel's default margins, in inches: the values a sheet with no <pageMargins> is printed with. */
export const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

/** The paper-size id whose dimensions match (within a tolerance), or undefined. */
export function paperSizeFor(widthMm: number, heightMm: number, tolMm = 3): number | undefined {
  // Landscape states the same paper turned round, so compare on the sorted pair.
  const [w, h] = widthMm <= heightMm ? [widthMm, heightMm] : [heightMm, widthMm];
  for (const [id, p] of Object.entries(PAPER_SIZES))
    if (Math.abs(p.w - w) <= tolMm && Math.abs(p.h - h) <= tolMm) return Number(id);
  return undefined;
}

/**
 * Split an xlsx header/footer string into its three regions. The syntax is a stream of &-codes in
 * which &L / &C / &R switch region; everything else (fields like &P, and the font/size codes) stays
 * in the text verbatim, so a header sheetedit did not author round-trips exactly.
 */
export function parseHeaderFooter(s: string | undefined): HeaderFooter | undefined {
  if (!s) return undefined;
  const out: HeaderFooter = {};
  let region: keyof HeaderFooter = "center"; // no leading code = centre, as Excel reads it
  let buf = "";
  const flush = () => { if (buf) out[region] = (out[region] ?? "") + buf; buf = ""; };
  for (let i = 0; i < s.length; i++) {
    if (s[i] === "&" && i + 1 < s.length) {
      const c = s[i + 1]!;
      if (c === "L" || c === "C" || c === "R") {
        flush();
        region = c === "L" ? "left" : c === "C" ? "center" : "right";
        i++;
        continue;
      }
      // &" and && are escapes/literals that must not be mistaken for a region switch.
      if (c === "&") { buf += "&&"; i++; continue; }
    }
    buf += s[i];
  }
  flush();
  return out.left || out.center || out.right ? out : undefined;
}

/** The inverse of parseHeaderFooter: the three regions back into one &-coded string. */
export function formatHeaderFooter(hf: HeaderFooter | undefined): string | undefined {
  if (!hf) return undefined;
  let s = "";
  if (hf.left) s += "&L" + hf.left;
  if (hf.center) s += "&C" + hf.center;
  if (hf.right) s += "&R" + hf.right;
  return s || undefined;
}

/** Whether anything at all is set, so an untouched sheet is never given a print block. */
export const hasPrintSetup = (p: PrintSetup | undefined): boolean =>
  !!p && Object.values(p).some((v) => (Array.isArray(v) ? v.length > 0 : v !== undefined));

/** The effective print area: what the file says, else the whole used range. */
export function effectivePrintArea(sheet: Sheet): { r1: number; c1: number; r2: number; c2: number }[] {
  const areas = sheet.printSetup?.printArea;
  if (areas?.length) return areas;
  if (!sheet.maxRow || !sheet.maxCol) return [];
  return [{ r1: 1, c1: 1, r2: sheet.maxRow, c2: sheet.maxCol }];
}

/** Add a manual break at `line`, or remove it when it is already there. Kept sorted and unique. */
export function toggleBreak(breaks: Breaks | undefined, line: number): Breaks {
  const set = new Set(breaks ?? []);
  if (set.has(line)) set.delete(line);
  else set.add(line);
  return [...set].sort((a, b) => a - b);
}
