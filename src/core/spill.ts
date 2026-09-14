// Deciding cheaply whether a cell's text could reach past its column.

/** Grid text size when a cell sets none (table.sheetedit-table is 13px). */
export const GRID_FONT_PX = 13;
/** The cell's own side padding, 8px each side. */
export const CELL_PAD_PX = 16;
/** Widest a character is taken to be, in em. Full-width CJK is 1em and some emoji a little more;
    Latin text is far narrower. Generous on purpose: this only rules out texts that cannot spill. */
export const WIDEST_CHAR_EM = 1.35;

/** A cell style's font size (points) in px, or the grid default. */
export const fontPx = (sizePt: number | undefined): number => (sizePt ? (sizePt * 96) / 72 : GRID_FONT_PX);

/**
 * True when `text` could not be wider than `colWidthPx` even if every character were as wide as the
 * widest script: then there is no spill to find and nothing needs measuring. Measuring means a forced
 * layout per cell, which was the largest single cost of drawing a workbook's grid.
 */
export function cannotOverflow(text: string, sizePx: number, colWidthPx: number): boolean {
  let chars = 0;
  for (const _ of text) chars++; // code points, so a surrogate pair counts once
  return CELL_PAD_PX + chars * sizePx * WIDEST_CHAR_EM <= colWidthPx;
}
