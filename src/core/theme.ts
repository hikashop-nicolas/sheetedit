import type { CellStyle, Workbook } from "./model";

// ---------------------------------------------------------------------------
// Workbook themes (xl/theme/theme1.xml)
// ---------------------------------------------------------------------------
// A workbook's theme is a named palette of twelve colours plus a heading and a body font. Cells do
// not store a theme colour, they store a *reference* to one (`<color theme="4" tint="-0.25"/>`), so
// switching the theme recolours every cell that used it while leaving the cells themselves alone.
//
// That is why the model keeps the reference next to the resolved colour: reading resolves it once
// for the grid, and a theme switch re-resolves from the new palette. Only theme1.xml is rewritten;
// styles.xml is untouched, exactly as Excel does it.

/** The twelve slots, in <clrScheme> element order. */
export const THEME_SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"] as const;
export type ThemeSlot = (typeof THEME_SLOTS)[number];

/**
 * The order a `<color theme="N">` index maps onto. It is NOT the clrScheme element order: the
 * first two pairs are swapped, which is a long-standing wart of the format.
 */
export const THEME_INDEX_ORDER: ThemeSlot[] = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

export type ThemeColors = Record<ThemeSlot, string>;

export interface WorkbookTheme {
  /** The scheme name Excel shows (Office, Berlin, ...). */
  name: string;
  colors: ThemeColors;
  /** Heading and body typefaces; a cell font tagged `<scheme val="major|minor"/>` follows these. */
  majorFont?: string;
  minorFont?: string;
}

const t = (name: string, hexes: string[], majorFont: string, minorFont: string): WorkbookTheme => ({
  name,
  colors: Object.fromEntries(THEME_SLOTS.map((s, i) => [s, hexes[i]!])) as ThemeColors,
  majorFont,
  minorFont,
});

/**
 * The palettes Excel ships, in clrScheme order (dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink).
 * Enough to offer a real choice without claiming to be the full Office set.
 */
export const BUILTIN_THEMES: WorkbookTheme[] = [
  t("Office", ["#000000", "#ffffff", "#44546a", "#e7e6e6", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#0563c1", "#954f72"], "Calibri Light", "Calibri"),
  t("Office 2007-2010", ["#000000", "#ffffff", "#1f497d", "#eeece1", "#4f81bd", "#c0504d", "#9bbb59", "#8064a2", "#4bacc6", "#f79646", "#0000ff", "#800080"], "Cambria", "Calibri"),
  t("Berlin", ["#000000", "#ffffff", "#3e3d2d", "#ceddc0", "#e97b1f", "#a0cf1c", "#26aad2", "#e4c60c", "#c74547", "#8b5fbf", "#0088cc", "#6a3f9e"], "Trebuchet MS", "Trebuchet MS"),
  t("Slice", ["#000000", "#ffffff", "#33456b", "#e8e9ea", "#052f61", "#a71930", "#b2c0d0", "#7d8da5", "#4a5a70", "#2e3e50", "#0563c1", "#954f72"], "Century Gothic", "Century Gothic"),
  t("Ion", ["#000000", "#ffffff", "#1e5155", "#e3ded1", "#b01513", "#ea6312", "#e6b729", "#6aac90", "#54849a", "#9e5e9b", "#58c1bf", "#dd8e17"], "Century Gothic", "Century Gothic"),
  t("Grayscale", ["#000000", "#ffffff", "#494949", "#f2f2f2", "#dddddd", "#b2b2b2", "#969696", "#808080", "#5f5f5f", "#4d4d4d", "#5f5f5f", "#919191"], "Corbel", "Corbel"),
];

export const DEFAULT_THEME = BUILTIN_THEMES[0]!;

/** A theme colour reference held on a cell style, so a palette switch can re-resolve it. */
export interface ThemeColorRef {
  /** The `<color theme="N">` index (into THEME_INDEX_ORDER). */
  index: number;
  /** Excel's tint: negative darkens toward black, positive lightens toward white. */
  tint?: number;
}

/** Excel tint applied to a "#rrggbb". */
export function applyTint(hex: string, tint: number): string {
  const ch = (i: number): string => {
    const c = parseInt(hex.slice(i, i + 2), 16);
    const v = tint < 0 ? c * (1 + tint) : c * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0");
  };
  return "#" + ch(1) + ch(3) + ch(5);
}

/** The palette as the flat array a `<color theme="N">` index addresses. */
export const themeIndexColors = (theme: WorkbookTheme): string[] => THEME_INDEX_ORDER.map((slot) => theme.colors[slot]);

/** Resolve one reference against a palette. */
export function resolveThemeRef(ref: ThemeColorRef, theme: WorkbookTheme): string {
  const base = themeIndexColors(theme)[ref.index] ?? "#000000";
  return ref.tint ? applyTint(base, ref.tint) : base;
}

/** Re-resolve a single cell style against a palette, in place. Returns true when anything changed. */
export function restyleForTheme(st: CellStyle, theme: WorkbookTheme): boolean {
  let changed = false;
  const set = <K extends "color" | "bg" | "fontFamily">(key: K, value: string | undefined): void => {
    if (value !== undefined && st[key] !== value) { st[key] = value; changed = true; }
  };
  if (st.colorRef) set("color", resolveThemeRef(st.colorRef, theme));
  if (st.bgRef) set("bg", resolveThemeRef(st.bgRef, theme));
  if (st.borderRefs) {
    const sides = { ...st.borders };
    for (const [side, ref] of Object.entries(st.borderRefs) as [keyof typeof sides, ThemeColorRef][]) {
      const next = resolveThemeRef(ref, theme);
      if (sides[side] !== next) { sides[side] = next; changed = true; }
    }
    st.borders = sides;
  }
  // A font tagged with a scheme follows the theme's typeface; the name in the file is only a cache.
  if (st.fontScheme) {
    const font = st.fontScheme === "major" ? theme.majorFont : theme.minorFont;
    if (font) set("fontFamily", font);
  }
  return changed;
}

/**
 * Switch the workbook's theme: re-resolve every style that referenced the old palette and flag the
 * theme part for rewriting. The cells keep their theme references untouched, which is what makes
 * this reversible and what Excel itself does.
 */
export function setWorkbookTheme(wb: Workbook, theme: WorkbookTheme): void {
  wb.theme = theme;
  wb.themeDirty = true;
  for (const sheet of wb.sheets)
    for (const cell of sheet.cells.values()) if (cell.cellStyle) restyleForTheme(cell.cellStyle, theme);
  // The resolved pool is what a freshly rendered cell reads, so it has to follow too.
  for (const st of wb.themeStyles ?? []) if (st) restyleForTheme(st, theme);
}

/** Whether a cell style takes any of its appearance from the theme. */
export const usesTheme = (st: CellStyle | undefined): boolean =>
  !!(st && (st.colorRef || st.bgRef || st.borderRefs || st.fontScheme));
