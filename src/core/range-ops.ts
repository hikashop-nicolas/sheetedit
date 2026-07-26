import { cellDisplay, getCell, type Sheet } from "./model";
import { setCellInput } from "./workbook";

// ---------------------------------------------------------------------------
// Sorting and filtering a range
// ---------------------------------------------------------------------------
// Shared by the grid's own filter menu and by VBA's Range.Sort / Range.AutoFilter, so there is one
// definition of what "sorted" and "filtered" mean rather than two that drift.

export interface Rect { r1: number; c1: number; r2: number; c2: number }

/** The text a sort compares: the formatted display, which is what the grid shows and Excel sorts. */
export const sortText = (sheet: Sheet, r: number, c: number): string => cellDisplay(getCell(sheet, r, c));

/** What a cell round-trips as when a sort moves it: its formula, or its raw value. */
export const editText = (sheet: Sheet, r: number, c: number): string => {
  const cell = getCell(sheet, r, c);
  return cell ? (cell.formula != null ? `=${cell.formula}` : cell.value) : "";
};

/**
 * Excel's sort order for two cell texts. Numbers sort numerically before text, and blanks go last
 * whichever direction the sort runs, which is why the caller must not simply negate this.
 */
export function compareSortKeys(a: string, b: string): number {
  const na = Number(a), nb = Number(b);
  const aNum = a !== "" && !Number.isNaN(na), bNum = b !== "" && !Number.isNaN(nb);
  if (a === "" && b !== "") return 1;
  if (b === "" && a !== "") return -1;
  if (aNum && bNum) return na - nb;
  if (aNum !== bNum) return aNum ? -1 : 1; // numbers come before text
  return a.localeCompare(b, undefined, { numeric: true });
}

/** Every cell a sort of this rectangle can move, which is what an undo step has to record. */
export function sortedPositions(rect: Rect, firstRowIsHeader = false): { r: number; c: number }[] {
  const out: { r: number; c: number }[] = [];
  for (let r = rect.r1 + (firstRowIsHeader ? 1 : 0); r <= rect.r2; r++) {
    for (let c = rect.c1; c <= rect.c2; c++) out.push({ r, c });
  }
  return out;
}

export interface SortKey {
  /** Absolute column index to sort on. */
  col: number;
  ascending: boolean;
}

/**
 * Reorder a rectangle's rows by one or more key columns. Values move; per-cell styles stay where
 * they are, and a formula travels as text without its references being adjusted, which matches
 * what the grid's own column sort has always done.
 *
 * Returns the positions it wrote, so a caller can record them for undo.
 */
export function sortRange(sheet: Sheet, rect: Rect, keys: SortKey[], firstRowIsHeader = false): { r: number; c: number }[] {
  const r0 = rect.r1 + (firstRowIsHeader ? 1 : 0);
  const cols: number[] = [];
  for (let c = rect.c1; c <= rect.c2; c++) cols.push(c);
  const positions = sortedPositions(rect, firstRowIsHeader);
  if (r0 >= rect.r2 || !keys.length) return positions;

  const rows = [];
  for (let r = r0; r <= rect.r2; r++) {
    const texts = new Map<number, string>();
    for (const c of cols) texts.set(c, editText(sheet, r, c));
    rows.push({ keys: keys.map((k) => sortText(sheet, r, k.col)), texts });
  }
  rows.sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      // Blanks stay last in both directions, so the direction is applied to the comparison of two
      // non-blank keys only.
      const ka = a.keys[i]!, kb = b.keys[i]!;
      if (ka === "" || kb === "") {
        const c = compareSortKeys(ka, kb);
        if (c !== 0) return c;
        continue;
      }
      const c = compareSortKeys(ka, kb) * (keys[i]!.ascending ? 1 : -1);
      if (c !== 0) return c;
    }
    return 0;
  });
  rows.forEach((row, i) => {
    const tr = r0 + i;
    for (const c of cols) setCellInput(sheet, tr, c, row.texts.get(c) ?? "");
  });
  return positions;
}

/**
 * Which data rows the sheet's per-column value filters hide. Pure: the caller applies the result,
 * because writing it back differs by format (xlsx flags the row element, ODF re-emits the sheet).
 */
export function filterHiddenRows(sheet: Sheet): Set<number> {
  const hidden = new Set<number>();
  const af = sheet.autoFilter;
  if (!af || !sheet.filters?.size) return hidden;
  for (let r = af.r1 + 1; r <= af.r2; r++) {
    for (const [col, allowed] of sheet.filters) {
      if (!allowed.has(sortText(sheet, r, col))) { hidden.add(r); break; }
    }
  }
  return hidden;
}

/** Every distinct display value in a column's data rows, which is what a filter offers. */
export function columnFilterValues(sheet: Sheet, col: number): string[] {
  const af = sheet.autoFilter;
  if (!af) return [];
  const seen = new Set<string>();
  for (let r = af.r1 + 1; r <= af.r2; r++) seen.add(sortText(sheet, r, col));
  return [...seen].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}
