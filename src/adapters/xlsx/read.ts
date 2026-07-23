import type { Cell, CellKind, CellStyle, Phonetic, Sheet, Workbook } from "../../core/model";
import { ensureCell, firstByLocal, formatNumber, key, noteExtent, numToStr, parseA1Ref, parseXml, parseXmlOpt, shiftFormula } from "../../core/model";
import { parseDxfs, readCondFormats } from "./condformat";
import { readCharts } from "./chart-read";
import { readImages } from "./image-read";
import { isDateFmt, isoToSerial } from "../../core/dates";

/** "A1:D10" (or "A1") -> a 1-based inclusive range, or null. */
function parseRangeRef(ref: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const [a, b] = ref.replace(/\$/g, "").split(":");
  const p1 = parseA1Ref(a ?? "");
  const p2 = b ? parseA1Ref(b) : p1;
  if (!p1 || !p2) return null;
  return { r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) };
}
// ---------------------------------------------------------------------------
// xlsx read: workbook/worksheet parsing, style pools resolution
// ---------------------------------------------------------------------------

export interface RichString {
  text: string; // base text (plain <t> and <r> runs)
  phonetic?: Phonetic[]; // <rPh> furigana runs, if any
}

// Parse a rich-text container (<si> in sharedStrings or <is> for an inline string) into its
// base text and any phonetic (furigana) runs. rPh runs carry the reading over base[sb..eb);
// they must not be folded into the text (that is the "東京トウキョウ" concatenation bug).
export function parseRichString(el: Element): RichString {
  let text = "";
  const phonetic: Phonetic[] = [];
  for (const ch of Array.from(el.children)) {
    if (ch.localName === "t") text += ch.textContent ?? "";
    else if (ch.localName === "r") {
      for (const rc of Array.from(ch.children)) if (rc.localName === "t") text += rc.textContent ?? "";
    } else if (ch.localName === "rPh") {
      const rt = Array.from(ch.children).find((x) => x.localName === "t");
      phonetic.push({ sb: Number(ch.getAttribute("sb") || "0"), eb: Number(ch.getAttribute("eb") || "0"), reading: rt?.textContent ?? "" });
    }
  }
  return phonetic.length ? { text, phonetic } : { text };
}

export function readSharedStrings(file: Uint8Array | undefined): RichString[] {
  if (!file) return [];
  const doc = parseXmlOpt(file);
  if (!doc) return [];
  return Array.from(doc.getElementsByTagName("si")).map(parseRichString);
}

export interface XlsxStyles {
  customFmt: Map<number, string>; // numFmtId -> format code (custom, id >= 164)
  xfNumFmtIds: number[]; // cellXfs index (the cell @s) -> numFmtId
  xfStyles: (CellStyle | undefined)[]; // cellXfs index (the cell @s) -> resolved style
}

// ARGB ("FFRRGGBB" or "RRGGBB") -> CSS "#rrggbb".
export function argbToCss(argb: string | null | undefined): string | undefined {
  if (!argb) return undefined;
  const h = argb.length === 8 ? argb.slice(2) : argb;
  return /^[0-9a-fA-F]{6}$/.test(h) ? "#" + h.toLowerCase() : undefined;
}

// Excel tint: negative darkens toward black, positive lightens toward white.
export function applyTint(hex: string, tint: number): string {
  const ch = (i: number) => {
    const c = parseInt(hex.slice(i, i + 2), 16);
    const v = tint < 0 ? c * (1 + tint) : c * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(v)))
      .toString(16)
      .padStart(2, "0");
  };
  return "#" + ch(1) + ch(3) + ch(5);
}

export const findByLocal = (doc: Document, local: string): Element | undefined =>
  Array.from(doc.getElementsByTagName("*")).find((e) => e.localName === local);

/** theme1.xml <clrScheme> -> a name->CSS map (accent1, dk1, lt1, ... plus tx1/bg1 aliases), for
    resolving DrawingML <a:schemeClr val="accentN"/> references in charts. */
export function readThemeMap(file: Uint8Array | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  const doc = file ? parseXmlOpt(file) : undefined;
  const scheme = doc && findByLocal(doc, "clrScheme");
  if (scheme) for (const el of Array.from(scheme.children)) {
    const c = el.firstElementChild;
    const css = c && (c.localName === "srgbClr" ? argbToCss(c.getAttribute("val")) : argbToCss(c.getAttribute("lastClr")));
    if (el.localName && css) out[el.localName] = css;
  }
  // Chart schemeClr often uses tx1/bg1/tx2/bg2 aliases for dk1/lt1/dk2/lt2.
  out.tx1 ??= out.dk1; out.bg1 ??= out.lt1; out.tx2 ??= out.dk2; out.bg2 ??= out.lt2;
  return out;
}

// theme1.xml <clrScheme> -> array indexed by a <color theme="N"> index.
export function readTheme(file: Uint8Array | undefined): string[] {
  const fallback = ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#0563c1", "#954f72"];
  if (!file) return fallback;
  try {
    const themeDoc = parseXmlOpt(file);
    if (!themeDoc) return [];
    const scheme = findByLocal(themeDoc, "clrScheme");
    if (!scheme) return fallback;
    const byName: Record<string, string> = {};
    for (const el of Array.from(scheme.children)) {
      const c = el.firstElementChild;
      const css = c && (c.localName === "srgbClr" ? argbToCss(c.getAttribute("val")) : argbToCss(c.getAttribute("lastClr")));
      if (el.localName && css) byName[el.localName] = css;
    }
    // theme index order swaps dk/lt 1 and 2 vs the clrScheme element order.
    const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    return order.map((nm, i) => byName[nm] ?? fallback[i]!);
  } catch {
    return fallback;
  }
}

// Resolve a <color> element (rgb, or theme + tint) to a CSS colour.
export function resolveColor(el: Element | undefined, theme: string[]): string | undefined {
  if (!el) return undefined;
  const rgb = el.getAttribute("rgb");
  if (rgb) return argbToCss(rgb);
  const t = el.getAttribute("theme");
  if (t != null) {
    const base = theme[Number(t)] ?? "#000000";
    const tint = Number(el.getAttribute("tint") || "0");
    return tint ? applyTint(base, tint) : base;
  }
  return undefined;
}

export function readXlsxStyles(doc: Document | undefined, theme: string[]): XlsxStyles {
  const customFmt = new Map<number, string>();
  const xfNumFmtIds: number[] = [];
  const xfStyles: (CellStyle | undefined)[] = [];
  if (!doc) return { customFmt, xfNumFmtIds, xfStyles };
  for (const nf of Array.from(doc.getElementsByTagName("numFmt"))) {
    const id = Number(nf.getAttribute("numFmtId"));
    const code = nf.getAttribute("formatCode");
    if (Number.isFinite(id) && code != null) customFmt.set(id, code);
  }

  const pool = (local: string) => {
    const parent = firstByLocal(doc.documentElement, local);
    return parent ? Array.from(parent.children).filter((e) => e.localName === local.replace(/s$/, "")) : [];
  };
  const fonts = pool("fonts").map((f) => ({
    bold: !!firstByLocal(f, "b"),
    italic: !!firstByLocal(f, "i"),
    underline: !!firstByLocal(f, "u"),
    // <u val="double"/> etc.; xlsx has no dotted/dashed, so only double is a distinct flavour.
    underlineStyle: firstByLocal(f, "u") && /double/i.test(firstByLocal(f, "u")!.getAttribute("val") || "single") ? "double" : undefined,
    strike: !!firstByLocal(f, "strike"),
    size: Number(firstByLocal(f, "sz")?.getAttribute("val")) || undefined,
    name: firstByLocal(f, "name")?.getAttribute("val") || undefined,
    color: resolveColor(firstByLocal(f, "color"), theme),
  }));
  const fills = pool("fills").map((fl) => {
    const pat = firstByLocal(fl, "patternFill");
    return pat?.getAttribute("patternType") === "solid" ? resolveColor(firstByLocal(pat, "fgColor"), theme) : undefined;
  });
  const borders = pool("borders").map((bd) => {
    const side = (name: string): string | undefined => {
      const s = firstByLocal(bd, name);
      return s?.getAttribute("style") ? (resolveColor(firstByLocal(s, "color"), theme) ?? "#444") : undefined;
    };
    const b = { top: side("top"), right: side("right"), bottom: side("bottom"), left: side("left") };
    return b.top || b.right || b.bottom || b.left ? b : undefined;
  });

  // The cell @s indexes <cellXfs>, not <cellStyleXfs>; read that list specifically.
  const cellXfs = doc.getElementsByTagName("cellXfs")[0];
  if (cellXfs) {
    for (const xf of Array.from(cellXfs.children)) {
      if (xf.localName !== "xf") continue;
      xfNumFmtIds.push(Number(xf.getAttribute("numFmtId") || "0"));
      const st: CellStyle = {};
      const font = fonts[Number(xf.getAttribute("fontId") || "0")];
      if (font) {
        if (font.bold) st.bold = true;
        if (font.italic) st.italic = true;
        if (font.underline) st.underline = true;
        if (font.underlineStyle) st.underlineStyle = font.underlineStyle;
        if (font.strike) st.strike = true;
        // Size/name only when they differ from the default font, so plain cells
        // keep an undefined cellStyle and the grid's own defaults.
        if (font.size && font.size !== fonts[0]?.size) st.fontSize = font.size;
        if (font.name && font.name !== fonts[0]?.name) st.fontFamily = font.name;
        if (font.color) st.color = font.color;
      }
      const fill = fills[Number(xf.getAttribute("fillId") || "0")];
      if (fill) st.bg = fill;
      const border = borders[Number(xf.getAttribute("borderId") || "0")];
      if (border) st.borders = border;
      const alignEl = firstByLocal(xf, "alignment");
      const align = alignEl?.getAttribute("horizontal");
      if (align === "center" || align === "right" || align === "left") st.align = align;
      const valign = alignEl?.getAttribute("vertical");
      if (valign === "top" || valign === "bottom") st.valign = valign;
      else if (valign === "center") st.valign = "middle";
      const wrap = alignEl?.getAttribute("wrapText");
      if (wrap === "1" || wrap === "true") st.wrap = true;
      xfStyles.push(Object.keys(st).length ? st : undefined);
    }
  }
  return { customFmt, xfNumFmtIds, xfStyles };
}

/** Resolve a cell's number format (code or built-in id), or undefined for General. */
export function resolveXlsxFmt(styles: XlsxStyles, s: string | undefined): string | number | undefined {
  if (s == null) return undefined;
  const numFmtId = styles.xfNumFmtIds[Number(s)];
  if (numFmtId == null || numFmtId === 0) return undefined; // 0 = General
  const custom = styles.customFmt.get(numFmtId);
  if (custom != null) return custom === "General" ? undefined : custom;
  return numFmtId; // built-in id; SSF resolves it
}

export function readXlsx(files: Record<string, Uint8Array>): Workbook {
  const wb: Workbook = { kind: "xlsx", sheets: [], files };
  const wbXml = files["xl/workbook.xml"];
  if (!wbXml) throw new Error("not an .xlsx: xl/workbook.xml missing");
  const wbDoc = parseXml(wbXml);
  // Workbook-level defined names: <definedNames><definedName name="X">Sheet1!$A$1:$A$10<...
  // Skip sheet-scoped names (localSheetId) and function/print built-ins; recalc reads the map.
  const definedNames = new Map<string, string>();
  for (const dn of Array.from(wbDoc.getElementsByTagName("definedName"))) {
    const name = dn.getAttribute("name");
    const target = dn.textContent?.trim();
    if (name && target && dn.getAttribute("localSheetId") == null && !name.startsWith("_xlnm")) definedNames.set(name, target);
  }
  if (definedNames.size) wb.definedNames = definedNames;
  const rels = new Map<string, string>();
  const relsFile = files["xl/_rels/workbook.xml.rels"];
  const relsDoc = relsFile ? parseXmlOpt(relsFile) : undefined;
  if (relsDoc) {
    for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
      const id = r.getAttribute("Id");
      const target = r.getAttribute("Target");
      if (id && target) rels.set(id, target);
    }
  }
  const shared = readSharedStrings(files["xl/sharedStrings.xml"]);
  const theme = readTheme(files["xl/theme/theme1.xml"]);
  const themeMap = readThemeMap(files["xl/theme/theme1.xml"]);
  wb.stylesDoc = files["xl/styles.xml"] ? parseXmlOpt(files["xl/styles.xml"]) : undefined;
  const styles = readXlsxStyles(wb.stylesDoc, theme);
  const dxfs = parseDxfs(wb.stylesDoc, theme);

  let n = 0;
  for (const sheetEl of Array.from(wbDoc.getElementsByTagName("sheet"))) {
    n++;
    const name = sheetEl.getAttribute("name") ?? `Sheet${n}`;
    const rid = sheetEl.getAttribute("r:id") ?? sheetEl.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "id",
    );
    let target = (rid && rels.get(rid)) || `worksheets/sheet${n}.xml`;
    const path = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
    const wsFile = files[path];
    const sheet: Sheet = { name, cells: new Map(), maxRow: 0, maxCol: 0, path };
    if (wsFile) {
      const doc = parseXml(wsFile);
      const sheetData = doc.getElementsByTagName("sheetData")[0];
      sheet.doc = doc;
      sheet.sheetData = sheetData;
      // Column widths: <cols><col min max width/></cols>. Width is in character units;
      // convert to px (~7px per char + padding for the default font).
      const colsEl = doc.getElementsByTagName("cols")[0];
      if (colsEl) {
        const cw = new Map<number, number>();
        const hiddenCols = new Set<number>();
        for (const col of Array.from(colsEl.children)) {
          if (col.localName !== "col") continue;
          const min = Number(col.getAttribute("min") || "0");
          const max = Number(col.getAttribute("max") || "0");
          const width = Number(col.getAttribute("width") || "0");
          if (!min) continue;
          const last = Math.min(max || min, min + 1000);
          const hidden = col.getAttribute("hidden") === "1" || col.getAttribute("hidden") === "true";
          for (let c = min; c <= last; c++) {
            if (width) cw.set(c, Math.round(width * 7 + 5));
            if (hidden) hiddenCols.add(c);
          }
        }
        if (cw.size) sheet.colWidths = cw;
        if (hiddenCols.size) sheet.hiddenCols = hiddenCols;
      }
      // Row heights: <row r ht customHeight hidden/>. ht is in points; convert to px (~4/3 px/pt).
      if (sheetData) {
        const rh = new Map<number, number>();
        const hiddenRows = new Set<number>();
        for (const rowEl of Array.from(sheetData.children)) {
          if (rowEl.localName !== "row") continue;
          const r = Number(rowEl.getAttribute("r") || "0");
          if (!r) continue;
          const ht = Number(rowEl.getAttribute("ht") || "0");
          if (ht) rh.set(r, Math.round((ht * 4) / 3));
          if (rowEl.getAttribute("hidden") === "1" || rowEl.getAttribute("hidden") === "true") hiddenRows.add(r);
        }
        if (rh.size) sheet.rowHeights = rh;
        if (hiddenRows.size) sheet.hiddenRows = hiddenRows;
      }
      // Autofilter range: <autoFilter ref="A1:D10"/>.
      const afEl = doc.getElementsByTagName("autoFilter")[0];
      const afRef = afEl?.getAttribute("ref");
      if (afRef) {
        const rng = parseRangeRef(afRef);
        if (rng) sheet.autoFilter = rng;
      }
      // Frozen panes: <sheetView><pane xSplit ySplit state="frozen"/></sheetView>.
      // xSplit / ySplit are the counts of frozen leading columns / rows.
      const pane = doc.getElementsByTagName("pane")[0];
      if (pane) {
        const state = pane.getAttribute("state");
        if (state === "frozen" || state === "frozenSplit") {
          const rows = Math.max(0, Math.floor(Number(pane.getAttribute("ySplit") || "0")));
          const cols = Math.max(0, Math.floor(Number(pane.getAttribute("xSplit") || "0")));
          if (rows > 0 || cols > 0) sheet.freeze = { rows, cols };
        }
      }
      // Merged ranges: <mergeCells><mergeCell ref="B1:C1"/></mergeCells>.
      const mergeEls = doc.getElementsByTagName("mergeCell");
      if (mergeEls.length) {
        const merges: { r1: number; c1: number; r2: number; c2: number }[] = [];
        for (const m of Array.from(mergeEls)) {
          const ref = m.getAttribute("ref");
          const [a, b] = (ref ?? "").split(":");
          const p1 = a ? parseA1Ref(a) : null;
          const p2 = b ? parseA1Ref(b) : null;
          if (p1 && p2) merges.push({ r1: p1.row, c1: p1.col, r2: p2.row, c2: p2.col });
        }
        if (merges.length) sheet.merges = merges;
      }
      if (sheetData) readSheetData(sheet, sheetData, shared, styles);
      readHyperlinks(sheet, doc, files, path);
      readDataValidations(sheet, doc);
      readCondFormats(sheet, doc, dxfs, theme);
      readComments(sheet, files, path);
      readCharts(sheet, files, path, themeMap);
      readImages(sheet, files, path);
    }
    wb.sheets.push(sheet);
  }
  return wb;
}

/** Read the worksheet's <hyperlinks> and attach a link to each cell in every referenced range.
    External targets resolve through the sheet's rels; internal ones use the location attribute. */
function readHyperlinks(sheet: Sheet, doc: Document, files: Record<string, Uint8Array>, path: string): void {
  const links = doc.getElementsByTagName("hyperlink");
  if (!links.length) return;
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const relsPath = path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const relsDoc = files[relsPath] ? parseXmlOpt(files[relsPath]) : undefined;
  const url = new Map<string, string>();
  if (relsDoc)
    for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
      const id = r.getAttribute("Id");
      const tgt = r.getAttribute("Target");
      if (id && tgt) url.set(id, tgt);
    }
  for (const h of Array.from(links)) {
    const ref = h.getAttribute("ref");
    if (!ref) continue;
    const rid = h.getAttribute("r:id") ?? h.getAttributeNS(REL, "id");
    const location = h.getAttribute("location");
    const tip = h.getAttribute("tooltip") ?? undefined;
    let href: string | undefined;
    let internal = false;
    if (rid && url.has(rid)) {
      href = url.get(rid)!;
      if (location) href += `#${location}`;
    } else if (location) {
      href = location;
      internal = true;
    }
    if (!href) continue;
    const [a, b] = ref.split(":");
    const p1 = parseA1Ref(a ?? "");
    const p2 = b ? parseA1Ref(b) : p1;
    if (!p1 || !p2) continue;
    for (let r = p1.row; r <= p2.row; r++)
      for (let c = p1.col; c <= p2.col; c++) ensureCell(sheet, r, c).link = { href, internal, tip };
  }
}

/** Read list-type <dataValidation> entries (dropdowns) into sheet.validations. Other validation
    types (whole/decimal/date/custom) are left to round-trip via the untouched sheet XML. */
function readDataValidations(sheet: Sheet, doc: Document): void {
  const dvs = doc.getElementsByTagName("dataValidation");
  if (!dvs.length) return;
  const out: NonNullable<Sheet["validations"]> = [];
  for (const dv of Array.from(dvs)) {
    if ((dv.getAttribute("type") || "") !== "list") continue;
    const sqref = dv.getAttribute("sqref") || (dv.getAttributeNS("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", "sqref") ?? "");
    if (!sqref) continue;
    const allowBlank = dv.getAttribute("allowBlank") === "1" || dv.getAttribute("allowBlank") === "true";
    const f1 = firstByLocal(dv, "formula1")?.textContent?.trim() ?? "";
    let values: string[] | undefined;
    let rangeRef: string | undefined;
    if (f1.startsWith('"') && f1.endsWith('"')) values = f1.slice(1, -1).split(",").map((s) => s.trim());
    else if (f1) rangeRef = f1;
    else continue;
    const ranges: { r1: number; c1: number; r2: number; c2: number }[] = [];
    for (const range of sqref.split(/\s+/)) {
      const [a, b] = range.split(":");
      const p1 = parseA1Ref(a ?? "");
      const p2 = b ? parseA1Ref(b) : p1;
      if (p1 && p2) ranges.push({ r1: p1.row, c1: p1.col, r2: p2.row, c2: p2.col });
    }
    if (ranges.length) out.push({ ranges, values, rangeRef, allowBlank });
  }
  if (out.length) sheet.validations = out;
}

/** Read legacy comments (comments*.xml) and threaded comments for a sheet, attaching each to its
    cell. Both parts are referenced from the sheet's rels; threaded-comment authors resolve
    through xl/persons/*.xml. The comment parts round-trip untouched; this only reads them. */
function readComments(sheet: Sheet, files: Record<string, Uint8Array>, path: string): void {
  const relsPath = path.replace(/worksheets\/(sheet[^/]+\.xml)$/i, "worksheets/_rels/$1.rels");
  const relsDoc = files[relsPath] ? parseXmlOpt(files[relsPath]) : undefined;
  if (!relsDoc) return;
  const resolve = (tgt: string): string => {
    if (tgt.startsWith("/")) return tgt.slice(1);
    const parts: string[] = [];
    for (const seg of `xl/worksheets/${tgt}`.split("/")) { if (seg === "..") parts.pop(); else if (seg !== ".") parts.push(seg); }
    return parts.join("/");
  };
  const targets = new Set<string>();
  for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
    const tgt = r.getAttribute("Target") ?? "";
    if (/comments\d*\.xml$/i.test(tgt) || /threadedComment/i.test(tgt)) targets.add(resolve(tgt));
  }
  if (!targets.size) return;
  const allByLocal = (root: Element | Document, local: string): Element[] =>
    Array.from(root.getElementsByTagName("*")).filter((e) => e.localName === local);
  // Persons map (threaded-comment authors).
  const persons = new Map<string, string>();
  for (const [fp, data] of Object.entries(files))
    if (/^xl\/persons\/.*\.xml$/i.test(fp)) { const pd = parseXmlOpt(data); if (pd) for (const p of allByLocal(pd, "person")) { const id = p.getAttribute("id"); if (id) persons.set(id, p.getAttribute("displayName") ?? ""); } }
  const add = (ref: string, author: string | undefined, text: string): void => {
    const p = parseA1Ref(ref);
    if (!p || !text) return;
    const cell = ensureCell(sheet, p.row, p.col);
    (cell.comments ??= []).push({ author: author || undefined, text });
  };
  for (const cp of targets) {
    const data = files[cp];
    const cd = data ? parseXmlOpt(data) : undefined;
    if (!cd) continue;
    if (allByLocal(cd, "commentList").length) {
      const authors = allByLocal(cd, "author").map((a) => a.textContent ?? "");
      for (const cm of allByLocal(cd, "comment")) {
        const ref = cm.getAttribute("ref");
        if (!ref) continue;
        const text = allByLocal(cm, "t").map((t) => t.textContent ?? "").join("").trim();
        add(ref, authors[Number(cm.getAttribute("authorId") || "0")], text);
      }
    } else {
      for (const tc of allByLocal(cd, "threadedComment")) {
        const ref = tc.getAttribute("ref");
        if (!ref) continue;
        const text = (allByLocal(tc, "text")[0]?.textContent ?? "").trim();
        add(ref, persons.get(tc.getAttribute("personId") ?? ""), text);
      }
    }
  }
}

export function readSheetData(sheet: Sheet, sheetData: Element, shared: RichString[], styles: XlsxStyles): void {
  // Shared formulas: the master <f t="shared" si ref> holds the text; children are
  // empty <f t="shared" si/>. Resolve each child to the master's formula shifted by
  // its offset so recalc (and a possible de-share on save) can treat it normally.
  const sharedMasters = new Map<string, { row: number; col: number; formula: string }>();
  const sharedChildren: Cell[] = [];
  for (const rowEl of Array.from(sheetData.getElementsByTagName("row"))) {
    const rAttr = rowEl.getAttribute("r");
    let rowNum = rAttr ? Number(rAttr) : 0;
    let colCursor = 0;
    for (const c of Array.from(rowEl.children)) {
      if (c.localName !== "c") continue;
      const ref = c.getAttribute("r");
      let row = rowNum;
      let col: number;
      if (ref) {
        const p = parseA1Ref(ref);
        if (!p) continue;
        row = p.row;
        col = p.col;
        colCursor = col;
      } else {
        col = ++colCursor;
      }
      if (!row) continue;
      const t = c.getAttribute("t");
      const fEl = firstByLocal(c, "f");
      const vEl = firstByLocal(c, "v");
      const isEl = firstByLocal(c, "is");
      const formulaText = fEl?.textContent?.trim();
      const formula = formulaText ? formulaText : undefined;

      let value = "";
      let kind: CellKind = "blank";
      let phonetic: Phonetic[] | undefined;
      if (t === "s") {
        const ss = shared[Number(vEl?.textContent ?? "0")] ?? { text: "" };
        value = ss.text;
        phonetic = ss.phonetic;
        kind = "s";
      } else if (t === "inlineStr") {
        const ss = isEl ? parseRichString(isEl) : { text: "" };
        value = ss.text;
        phonetic = ss.phonetic;
        kind = "s";
      } else if (t === "str") {
        value = vEl?.textContent ?? "";
        kind = "s";
      } else if (t === "b") {
        value = vEl?.textContent === "1" ? "TRUE" : "FALSE";
        kind = "b";
      } else if (t === "e") {
        value = vEl?.textContent ?? "";
        kind = "e";
      } else if (t === "d") {
        // ISO 8601 date cell: model it as a serial so formulas and formats work.
        const iso = vEl?.textContent ?? "";
        const serial = iso === "" ? null : isoToSerial(iso);
        if (serial != null) {
          value = numToStr(serial);
          kind = "n";
        } else {
          value = iso;
          kind = iso === "" ? "blank" : "s";
        }
      } else {
        value = vEl?.textContent ?? "";
        kind = value === "" ? "blank" : "n";
      }

      const cell: Cell = {
        row,
        col,
        value,
        kind,
        formula,
        el: c,
        style: c.getAttribute("s") ?? undefined,
        phonetic,
      };
      // Legacy array formula: the top-left cell carries <f t="array" ref="A1:C3">.
      if (fEl?.getAttribute("t") === "array") cell.arrayRef = fEl.getAttribute("ref") ?? undefined;
      if (kind === "n") {
        const fmt = resolveXlsxFmt(styles, cell.style);
        if (fmt != null) {
          cell.numFmt = fmt;
          const d = formatNumber(fmt, value);
          if (d != null) cell.display = d;
        }
        // A t="d" cell whose style has no date format still must display as a date.
        if (t === "d" && !isDateFmt(cell.numFmt)) {
          cell.numFmt = (vEl?.textContent ?? "").includes("T") ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd";
          cell.numFmtDirty = true;
          const d = formatNumber(cell.numFmt, value);
          if (d != null) cell.display = d;
        }
      }
      if (cell.style != null) cell.cellStyle = styles.xfStyles[Number(cell.style)];
      if (fEl?.getAttribute("t") === "shared") {
        const si = fEl.getAttribute("si");
        if (si != null) {
          cell.sharedSi = si;
          if (formula) sharedMasters.set(si, { row, col, formula });
          else sharedChildren.push(cell);
        }
      }
      sheet.cells.set(key(row, col), cell);
      noteExtent(sheet, row, col);
    }
  }
  for (const cell of sharedChildren) {
    const m = sharedMasters.get(cell.sharedSi!);
    if (m) cell.formula = shiftFormula(m.formula, cell.row - m.row, cell.col - m.col);
  }
}
