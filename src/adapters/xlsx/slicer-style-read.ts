import { parseXmlOpt, type Workbook } from "../../core/model";
import { resolveColor } from "./read";

// User-defined slicer styles. Excel keeps them in xl/styles.xml's extLst: an <x14:slicerStyles>
// group under {EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}, where each <x14:slicerStyle name="..."> lists
// <x14:slicerStyleElement type dxfId>. The dxfId indexes the x14 dxfs list, which sits in its own
// ext of the same extLst. Built-in styles (SlicerStyleLight1 and friends) are NOT defined here -
// only the name is stored on the slicer, so those stay mapped to a theme accent.

/** The colours a custom slicer style gives its items, resolved to CSS. */
export interface SlicerStyleDef {
  selectedFill?: string;
  selectedText?: string;
  unselectedFill?: string;
  unselectedText?: string;
  /** Only set when the style names a table style that supplies the rest of the look. */
  tableStyle?: string;
}

const byLocal = (root: Element | Document, local: string): Element[] =>
  Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);
const firstLocal = (root: Element, local: string): Element | undefined =>
  Array.from(root.getElementsByTagName("*")).find((e) => e.localName === local);

/** Fill and font colour of one dxf, as CSS. */
function dxfColors(dxf: Element, theme: string[]): { fill?: string; text?: string } {
  const pat = firstLocal(dxf, "patternFill");
  // In a dxf the solid colour lives in bgColor, not fgColor (same quirk the cf reader handles).
  const fill = pat
    ? resolveColor(firstLocal(pat, "bgColor"), theme) ?? resolveColor(firstLocal(pat, "fgColor"), theme)
    : undefined;
  const font = firstLocal(dxf, "font");
  const text = font ? resolveColor(firstLocal(font, "color"), theme) : undefined;
  return { fill, text };
}

/**
 * Read the workbook's user-defined slicer styles, keyed by name.
 * The x14 dxfs list is found by element name inside the styles extLst rather than by ext URI, so a
 * file that groups the extensions differently still resolves.
 */
export function readSlicerStyles(files: Record<string, Uint8Array>, theme: string[]): Map<string, SlicerStyleDef> {
  const out = new Map<string, SlicerStyleDef>();
  const doc = files["xl/styles.xml"] ? parseXmlOpt(files["xl/styles.xml"]!) : undefined;
  const extLst = doc ? Array.from(doc.documentElement.children).find((e) => e.localName === "extLst") : undefined;
  if (!extLst) return out;
  const group = byLocal(extLst, "slicerStyles")[0];
  if (!group) return out;
  // The x14 dxfs: the <dxfs> inside the extension list, not the styleSheet's own top-level one.
  const dxfs = byLocal(extLst, "dxfs")[0];
  const dxfList = dxfs ? Array.from(dxfs.children).filter((e) => e.localName === "dxf") : [];
  for (const style of byLocal(group, "slicerStyle")) {
    const name = style.getAttribute("name");
    if (!name) continue;
    const def: SlicerStyleDef = { tableStyle: name };
    for (const el of byLocal(style, "slicerStyleElement")) {
      const id = Number(el.getAttribute("dxfId"));
      const dxf = Number.isFinite(id) ? dxfList[id] : undefined;
      if (!dxf) continue;
      const { fill, text } = dxfColors(dxf, theme);
      switch (el.getAttribute("type")) {
        case "selectedItemWithData": if (fill) def.selectedFill = fill; if (text) def.selectedText = text; break;
        case "unselectedItemWithData": if (fill) def.unselectedFill = fill; if (text) def.unselectedText = text; break;
        // The no-data and hovered variants are read past: sheetedit shows every cache item and has
        // no hover state of its own, so they would never be applied.
        default: break;
      }
    }
    out.set(name, def);
  }
  return out;
}

/** Attach the workbook's custom slicer styles so the overlay can colour by name. */
export function readXlsxSlicerStyles(wb: Workbook, files: Record<string, Uint8Array>, theme: string[]): void {
  const styles = readSlicerStyles(files, theme);
  if (styles.size) wb.slicerStyles = styles;
}
