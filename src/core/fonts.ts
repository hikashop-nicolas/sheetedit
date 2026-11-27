import { isSheetLoaded } from "./lazy-sheet";
import type { CellStyle, Workbook } from "./model";

/** The workbook's mount option for the font picker (same shape as richdoc's). */
export interface FontOptions {
  /** Offer "Show my fonts" (installed fonts, found on the device; nothing is sent). Default true. */
  local?: boolean;
  /** Offer "Load a font file" (.ttf, .otf, .ttc, .woff, .woff2, kept for later sessions). Default true. */
  file?: boolean;
  /** Offer "Get more fonts" from Google Fonts after an in-app consent. Default off. */
  remote?: false | { provider: "google" };
}

/**
 * Every font the workbook names: the Normal font, the theme's heading and body fonts, the style
 * pool, and the cell styles and rich-text runs of the sheets parsed so far (an unparsed sheet is
 * left alone; the editor adds its fonts once it loads).
 */
export function workbookFonts(wb: Workbook): string[] {
  const out = new Map<string, string>();
  const add = (name: string | undefined): void => {
    const n = name?.trim();
    if (n && !out.has(n.toLowerCase())) out.set(n.toLowerCase(), n);
  };
  const seen = new Set<CellStyle>();
  const addStyle = (st: CellStyle | undefined): void => {
    if (!st || seen.has(st)) return;
    seen.add(st);
    add(st.fontFamily);
  };
  add(wb.defaultFontName);
  add(wb.theme?.minorFont);
  add(wb.theme?.majorFont);
  for (const st of wb.themeStyles ?? []) addStyle(st);
  for (const sheet of wb.sheets) {
    if (!isSheetLoaded(sheet)) continue;
    for (const cell of sheet.cells.values()) {
      addStyle(cell.cellStyle);
      if (cell.richRuns) for (const run of cell.richRuns) add(run.font);
    }
  }
  return [...out.values()];
}
