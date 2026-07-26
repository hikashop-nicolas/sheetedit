import { parseXmlOpt } from "../../core/model";
import { DEFAULT_THEME, THEME_SLOTS, type ThemeColors, type WorkbookTheme } from "../../core/theme";
import { argbToCss, findByLocal } from "./read";

// xl/theme/theme1.xml holds the workbook's <a:clrScheme> (twelve colours) and <a:fontScheme> (the
// heading and body typefaces). Cells reference the palette by index, so this is the one place the
// actual colours live.

/** The CSS colour of a <a:dk1>/<a:accent1>/... wrapper, which holds either an srgbClr or a sysClr. */
function slotColor(el: Element | undefined): string | undefined {
  const c = el?.firstElementChild;
  if (!c) return undefined;
  // A sysClr (windowText / window) carries the resolved value in @lastClr.
  return c.localName === "srgbClr" ? argbToCss(c.getAttribute("val")) : argbToCss(c.getAttribute("lastClr"));
}

/** The latin typeface of an <a:majorFont> / <a:minorFont>. */
function schemeFont(el: Element | undefined): string | undefined {
  if (!el) return undefined;
  const latin = Array.from(el.children).find((e) => e.localName === "latin");
  return latin?.getAttribute("typeface") || undefined;
}

/** Parse xl/theme/theme1.xml into the model, falling back to Office for anything it omits. */
export function readXlsxTheme(file: Uint8Array | undefined): WorkbookTheme {
  const doc = file ? parseXmlOpt(file) : undefined;
  if (!doc) return { ...DEFAULT_THEME, colors: { ...DEFAULT_THEME.colors } };
  const scheme = findByLocal(doc, "clrScheme");
  const colors = { ...DEFAULT_THEME.colors } as ThemeColors;
  if (scheme) {
    for (const slot of THEME_SLOTS) {
      const el = Array.from(scheme.children).find((e) => e.localName === slot);
      const css = slotColor(el);
      if (css) colors[slot] = css;
    }
  }
  const fontScheme = findByLocal(doc, "fontScheme");
  const major = fontScheme ? Array.from(fontScheme.children).find((e) => e.localName === "majorFont") : undefined;
  const minor = fontScheme ? Array.from(fontScheme.children).find((e) => e.localName === "minorFont") : undefined;
  // The theme element's own name is what Excel shows in the theme gallery.
  const themeEl = doc.documentElement;
  return {
    name: scheme?.getAttribute("name") || themeEl.getAttribute("name") || DEFAULT_THEME.name,
    colors,
    majorFont: schemeFont(major) ?? DEFAULT_THEME.majorFont,
    minorFont: schemeFont(minor) ?? DEFAULT_THEME.minorFont,
  };
}
