import type { Cell, CellKind, CellStyle, Sheet, StyleChange, Workbook } from "./model";
import { colToLetters, firstByLocal, formatNumber, key, noteExtent, parseA1Ref, parseXml, parseXmlOpt, removeByLocal, serializeXml, shiftFormula } from "./model";
// ---------------------------------------------------------------------------
// xlsx (OOXML SpreadsheetML)
// ---------------------------------------------------------------------------

export const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export function readSharedStrings(file: Uint8Array | undefined): string[] {
  if (!file) return [];
  const doc = parseXmlOpt(file);
  if (!doc) return [];
  return Array.from(doc.getElementsByTagName("si")).map((si) =>
    Array.from(si.getElementsByTagName("t"))
      .map((t) => t.textContent ?? "")
      .join(""),
  );
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
        if (font.color) st.color = font.color;
      }
      const fill = fills[Number(xf.getAttribute("fillId") || "0")];
      if (fill) st.bg = fill;
      const border = borders[Number(xf.getAttribute("borderId") || "0")];
      if (border) st.borders = border;
      const align = firstByLocal(xf, "alignment")?.getAttribute("horizontal");
      if (align === "center" || align === "right" || align === "left") st.align = align;
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
  wb.stylesDoc = files["xl/styles.xml"] ? parseXmlOpt(files["xl/styles.xml"]) : undefined;
  const styles = readXlsxStyles(wb.stylesDoc, theme);

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
        for (const col of Array.from(colsEl.children)) {
          if (col.localName !== "col") continue;
          const min = Number(col.getAttribute("min") || "0");
          const max = Number(col.getAttribute("max") || "0");
          const width = Number(col.getAttribute("width") || "0");
          if (!min || !width) continue;
          const px = Math.round(width * 7 + 5);
          for (let c = min; c <= Math.min(max || min, min + 1000); c++) cw.set(c, px);
        }
        if (cw.size) sheet.colWidths = cw;
      }
      // Row heights: <row r ht customHeight/>. ht is in points; convert to px (~4/3 px/pt).
      if (sheetData) {
        const rh = new Map<number, number>();
        for (const rowEl of Array.from(sheetData.children)) {
          if (rowEl.localName !== "row") continue;
          const r = Number(rowEl.getAttribute("r") || "0");
          const ht = Number(rowEl.getAttribute("ht") || "0");
          if (r && ht) rh.set(r, Math.round((ht * 4) / 3));
        }
        if (rh.size) sheet.rowHeights = rh;
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
    }
    wb.sheets.push(sheet);
  }
  return wb;
}

export function readSheetData(sheet: Sheet, sheetData: Element, shared: string[], styles: XlsxStyles): void {
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
      if (t === "s") {
        value = shared[Number(vEl?.textContent ?? "0")] ?? "";
        kind = "s";
      } else if (t === "inlineStr") {
        value = isEl ? Array.from(isEl.getElementsByTagName("t")).map((x) => x.textContent ?? "").join("") : "";
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
      };
      if (kind === "n") {
        const fmt = resolveXlsxFmt(styles, cell.style);
        if (fmt != null) {
          cell.numFmt = fmt;
          const d = formatNumber(fmt, value);
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

export function ensureXlsxCellEl(sheet: Sheet, cell: Cell): Element {
  if (cell.el) return cell.el;
  const doc = sheet.doc!;
  const sheetData = sheet.sheetData!;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  // find or create the <row>
  let rowEl: Element | undefined;
  let insertRowBefore: Element | null = null;
  for (const r of Array.from(sheetData.children)) {
    if (r.localName !== "row") continue;
    const rn = Number(r.getAttribute("r") || "0");
    if (rn === cell.row) {
      rowEl = r;
      break;
    }
    if (rn > cell.row) {
      insertRowBefore = r;
      break;
    }
  }
  if (!rowEl) {
    rowEl = doc.createElementNS(ns, "row");
    rowEl.setAttribute("r", String(cell.row));
    sheetData.insertBefore(rowEl, insertRowBefore);
  }
  // find or create the <c> in column order
  const ref = colToLetters(cell.col) + cell.row;
  let insertCellBefore: Element | null = null;
  for (const c of Array.from(rowEl.children)) {
    if (c.localName !== "c") continue;
    const cref = c.getAttribute("r");
    const p = cref ? parseA1Ref(cref) : null;
    if (p && p.col === cell.col) return (cell.el = c);
    if (p && p.col > cell.col) {
      insertCellBefore = c;
      break;
    }
  }
  const cEl = doc.createElementNS(ns, "c");
  cEl.setAttribute("r", ref);
  if (cell.style) cEl.setAttribute("s", cell.style);
  rowEl.insertBefore(cEl, insertCellBefore);
  cell.el = cEl;
  return cEl;
}

export function writeXlsxCell(sheet: Sheet, cell: Cell, plainFormula = false): void {
  const doc = sheet.doc!;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const c = ensureXlsxCellEl(sheet, cell);
  // When only the cached value changed, keep the original <f> untouched so
  // t="shared"/"array", @si and @ref survive. A formula edit (or a group
  // de-share) rewrites <f> as a plain formula instead.
  const oldF = firstByLocal(c, "f");
  const keepF = cell.formula != null && oldF != null && !cell.fDirty && !plainFormula;
  if (!keepF) removeByLocal(c, "f");
  removeByLocal(c, "v");
  removeByLocal(c, "is");
  const addV = (text: string) => {
    const v = doc.createElementNS(ns, "v");
    v.textContent = text;
    c.appendChild(v);
  };
  if (cell.formula != null) {
    if (!keepF) {
      const f = doc.createElementNS(ns, "f");
      f.textContent = cell.formula;
      c.appendChild(f);
    }
    if (cell.kind === "n") {
      c.removeAttribute("t");
      if (cell.value !== "") addV(cell.value);
    } else if (cell.kind === "b") {
      c.setAttribute("t", "b");
      addV(cell.value === "TRUE" || cell.value === "1" ? "1" : "0");
    } else if (cell.kind === "e") {
      c.setAttribute("t", "e");
      addV(cell.value);
    } else if (cell.kind === "blank" || cell.value === "") {
      c.removeAttribute("t");
    } else {
      c.setAttribute("t", "str");
      addV(cell.value);
    }
    return;
  }
  // literal
  if (cell.value === "" || cell.kind === "blank") {
    c.removeAttribute("t");
  } else if (cell.kind === "n") {
    c.removeAttribute("t");
    addV(cell.value);
  } else if (cell.kind === "b") {
    c.setAttribute("t", "b");
    addV(cell.value === "TRUE" || cell.value === "1" ? "1" : "0");
  } else if (cell.kind === "e") {
    c.setAttribute("t", "e");
    addV(cell.value);
  } else {
    c.setAttribute("t", "inlineStr");
    const is = doc.createElementNS(ns, "is");
    const t = doc.createElementNS(ns, "t");
    t.setAttribute("xml:space", "preserve");
    t.textContent = cell.value;
    is.appendChild(t);
    c.appendChild(is);
  }
}

// Set a single column's width (px) in the worksheet's <cols>, creating <cols>/<col>
// as needed and splitting any existing run that covers the column. Keeps colWidths in sync.
export function setXlsxColWidth(sheet: Sheet, col: number, px: number): void {
  if (!sheet.colWidths) sheet.colWidths = new Map();
  sheet.colWidths.set(col, px);
  const doc = sheet.doc;
  if (!doc) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const width = Math.max(0, (px - 5) / 7);
  let colsEl = doc.getElementsByTagName("cols")[0] as Element | undefined;
  if (!colsEl) {
    colsEl = doc.createElementNS(ns, "cols");
    // <cols> must precede <sheetData> per the schema.
    sheet.sheetData?.parentNode?.insertBefore(colsEl, sheet.sheetData);
  }
  // Narrow any run that spans `col` so we can give `col` its own entry.
  for (const c of Array.from(colsEl.children)) {
    if (c.localName !== "col") continue;
    const min = Number(c.getAttribute("min") || "0");
    const max = Number(c.getAttribute("max") || String(min));
    if (col < min || col > max) continue;
    if (min === max) {
      c.setAttribute("width", String(width));
      c.setAttribute("customWidth", "1");
      sheet.layoutDirty = true;
      return;
    }
    // Split: left part [min..col-1], right part [col+1..max], plus the single col.
    if (col > min) {
      const left = c.cloneNode(true) as Element;
      left.setAttribute("min", String(min));
      left.setAttribute("max", String(col - 1));
      colsEl.insertBefore(left, c);
    }
    if (col < max) {
      const right = c.cloneNode(true) as Element;
      right.setAttribute("min", String(col + 1));
      right.setAttribute("max", String(max));
      colsEl.insertBefore(right, c);
    }
    c.setAttribute("min", String(col));
    c.setAttribute("max", String(col));
    c.setAttribute("width", String(width));
    c.setAttribute("customWidth", "1");
    sheet.layoutDirty = true;
    return;
  }
  const colEl = doc.createElementNS(ns, "col");
  colEl.setAttribute("min", String(col));
  colEl.setAttribute("max", String(col));
  colEl.setAttribute("width", String(width));
  colEl.setAttribute("customWidth", "1");
  colsEl.appendChild(colEl);
  sheet.layoutDirty = true;
}

// Set a single row's height (px) on its <row>, creating the row element if absent.
export function setXlsxRowHeight(sheet: Sheet, row: number, px: number): void {
  if (!sheet.rowHeights) sheet.rowHeights = new Map();
  sheet.rowHeights.set(row, px);
  const doc = sheet.doc;
  const sd = sheet.sheetData;
  if (!doc || !sd) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const pt = (px * 3) / 4;
  let rowEl: Element | undefined;
  for (const re of Array.from(sd.children)) {
    if (re.localName === "row" && Number(re.getAttribute("r") || "0") === row) {
      rowEl = re;
      break;
    }
  }
  if (!rowEl) {
    rowEl = doc.createElementNS(ns, "row");
    rowEl.setAttribute("r", String(row));
    // Insert keeping rows in ascending order.
    let next: Element | null = null;
    for (const re of Array.from(sd.children)) {
      if (re.localName === "row" && Number(re.getAttribute("r") || "0") > row) {
        next = re;
        break;
      }
    }
    sd.insertBefore(rowEl, next);
  }
  rowEl.setAttribute("ht", String(pt));
  rowEl.setAttribute("customHeight", "1");
  sheet.layoutDirty = true;
}

// Add or remove a merged range (1-based, inclusive). The top-left cell shows through;
// any cells the merge hides keep their data (so unmerging restores it). Updates the
// worksheet's <mergeCells> element and the in-memory merge list.
export function setXlsxMerge(
  sheet: Sheet,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  merge: boolean,
): void {
  const top = Math.min(r1, r2),
    left = Math.min(c1, c2),
    bottom = Math.max(r1, r2),
    right = Math.max(c1, c2);
  const ref = `${colToLetters(left)}${top}:${colToLetters(right)}${bottom}`;
  const merges = (sheet.merges ??= []);
  const idx = merges.findIndex((m) => m.r1 === top && m.c1 === left && m.r2 === bottom && m.c2 === right);
  if (merge) {
    if (idx === -1) merges.push({ r1: top, c1: left, r2: bottom, c2: right });
  } else if (idx !== -1) {
    merges.splice(idx, 1);
  }

  const doc = sheet.doc;
  if (!doc) return;
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  let mcEl = doc.getElementsByTagName("mergeCells")[0] as Element | undefined;
  if (merge) {
    if (!mcEl) {
      mcEl = doc.createElementNS(ns, "mergeCells");
      // <mergeCells> follows <sheetData> in the schema.
      sheet.sheetData?.parentNode?.insertBefore(mcEl, sheet.sheetData.nextSibling);
    }
    const exists = Array.from(mcEl.children).some((m) => m.getAttribute("ref") === ref);
    if (!exists) {
      const mc = doc.createElementNS(ns, "mergeCell");
      mc.setAttribute("ref", ref);
      mcEl.appendChild(mc);
    }
  } else if (mcEl) {
    for (const m of Array.from(mcEl.children))
      if (m.getAttribute("ref") === ref) mcEl.removeChild(m);
  }
  if (mcEl) {
    if (mcEl.children.length === 0) mcEl.parentNode?.removeChild(mcEl);
    else mcEl.setAttribute("count", String(mcEl.children.length));
  }
  sheet.layoutDirty = true;
}

export const xmlOf = (el: Element): string => new XMLSerializer().serializeToString(el);

// Find a matching child in a style pool (deduped by serialized form) or append it;
// returns its index and keeps the pool's count attribute in sync.
export function poolIndex(parent: Element, candidate: Element): number {
  const want = xmlOf(candidate);
  const kids = Array.from(parent.children);
  for (let i = 0; i < kids.length; i++) if (xmlOf(kids[i]!) === want) return i;
  parent.appendChild(candidate);
  parent.setAttribute("count", String(parent.children.length));
  return parent.children.length - 1;
}

export const argbOf = (css: string): string => "FF" + css.replace("#", "").toUpperCase();

// Compute the resulting CellStyle from the current one plus a change. Shared by the
// xlsx and ods style writers so both stay consistent; affected borders become black.
export function mergeCellStyle(cur: CellStyle, change: StyleChange): CellStyle {
  const sides = {
    top: !!cur.borders?.top,
    right: !!cur.borders?.right,
    bottom: !!cur.borders?.bottom,
    left: !!cur.borders?.left,
  };
  if (change.border !== undefined) sides.top = sides.right = sides.bottom = sides.left = change.border;
  if (change.borderSides) Object.assign(sides, change.borderSides);
  const any = sides.top || sides.right || sides.bottom || sides.left;
  return {
    bold: change.bold ?? cur.bold,
    italic: change.italic ?? cur.italic,
    color: change.color ?? cur.color,
    bg: change.bg ?? cur.bg,
    align: change.align ?? cur.align,
    borders: any
      ? {
          top: sides.top ? "#000000" : undefined,
          right: sides.right ? "#000000" : undefined,
          bottom: sides.bottom ? "#000000" : undefined,
          left: sides.left ? "#000000" : undefined,
        }
      : undefined,
  };
}

/**
 * Apply a style change to a cell, managing the xlsx style pools: derive a new font /
 * fill / border from the cell's current format plus the change, find-or-create each in
 * styles.xml, find-or-create the combined <xf>, and point the cell at it.
 */
export function setXlsxCellStyle(wb: Workbook, sheet: Sheet, cell: Cell, change: StyleChange): void {
  const doc = wb.stylesDoc;
  if (!doc || !sheet.doc || !sheet.sheetData) return; // chartsheets have no cells to style
  const ns = doc.documentElement.namespaceURI || SS_MAIN;
  const ce = (name: string) => doc.createElementNS(ns, name);
  const root = doc.documentElement;
  const pool = (name: string): Element => firstByLocal(root, name) ?? (root.appendChild(ce(name)) as Element);
  const fontsEl = pool("fonts");
  const fillsEl = pool("fills");
  const bordersEl = pool("borders");
  const cellXfsEl = pool("cellXfs");

  const curXf = cellXfsEl.children[cell.style ? Number(cell.style) : 0];
  const numFmtId = curXf?.getAttribute("numFmtId") || "0";
  const curFontId = Number(curXf?.getAttribute("fontId") || "0");
  const curFillId = Number(curXf?.getAttribute("fillId") || "0");
  const curBorderId = Number(curXf?.getAttribute("borderId") || "0");

  const cur = cell.cellStyle ?? {};
  const bold = change.bold ?? cur.bold;
  const italic = change.italic ?? cur.italic;
  const color = change.color ?? cur.color;
  const bg = change.bg ?? cur.bg;
  const align = change.align ?? cur.align;
  // Border sides: start from the current borders, apply the all-sides toggle and/or per-side change.
  const curSides = {
    top: !!cur.borders?.top,
    right: !!cur.borders?.right,
    bottom: !!cur.borders?.bottom,
    left: !!cur.borders?.left,
  };
  let sides = curSides;
  let borderChanged = false;
  if (change.border !== undefined) {
    sides = { top: change.border, right: change.border, bottom: change.border, left: change.border };
    borderChanged = true;
  }
  if (change.borderSides) {
    sides = { ...sides, ...change.borderSides };
    borderChanged = true;
  }

  // Font: clone the current one and toggle bold/italic/colour.
  const baseFont = fontsEl.children[curFontId];
  const font = baseFont ? (baseFont.cloneNode(true) as Element) : ce("font");
  const flag = (tag: string, on: boolean | undefined) => {
    const ex = firstByLocal(font, tag);
    if (on && !ex) font.appendChild(ce(tag));
    else if (!on && ex) font.removeChild(ex);
  };
  flag("b", bold);
  flag("i", italic);
  if (color) {
    const col = firstByLocal(font, "color") ?? (font.appendChild(ce("color")) as Element);
    col.removeAttribute("theme");
    col.removeAttribute("tint");
    col.setAttribute("rgb", argbOf(color));
  }
  const fontId = poolIndex(fontsEl, font);

  // Fill (solid) when set; else keep the current fill.
  let fillId = curFillId;
  if (bg) {
    const fill = ce("fill");
    const pat = ce("patternFill");
    pat.setAttribute("patternType", "solid");
    const fg = ce("fgColor");
    fg.setAttribute("rgb", argbOf(bg));
    pat.appendChild(fg);
    fill.appendChild(pat);
    fillId = poolIndex(fillsEl, fill);
  }

  // Border: rebuild the per-side border element only when the change touches borders.
  let borderId = curBorderId;
  if (borderChanged) {
    const bd = ce("border");
    for (const side of ["left", "right", "top", "bottom"] as const) {
      const s = ce(side);
      if (sides[side]) {
        s.setAttribute("style", "thin");
        const cc = ce("color");
        cc.setAttribute("rgb", "FF000000");
        s.appendChild(cc);
      }
      bd.appendChild(s);
    }
    borderId = poolIndex(bordersEl, bd);
  }

  const xf = ce("xf");
  xf.setAttribute("numFmtId", numFmtId);
  xf.setAttribute("fontId", String(fontId));
  xf.setAttribute("fillId", String(fillId));
  xf.setAttribute("borderId", String(borderId));
  xf.setAttribute("xfId", "0");
  xf.setAttribute("applyFont", "1");
  if (bg) xf.setAttribute("applyFill", "1");
  if (borderId) xf.setAttribute("applyBorder", "1");
  if (align) {
    xf.setAttribute("applyAlignment", "1");
    const a = ce("alignment");
    a.setAttribute("horizontal", align);
    xf.appendChild(a);
  }
  const sIdx = poolIndex(cellXfsEl, xf);

  cell.style = String(sIdx);
  ensureXlsxCellEl(sheet, cell).setAttribute("s", String(sIdx));
  cell.cellStyle = mergeCellStyle(cur, change);
  cell.edited = true;
  wb.stylesDirty = true;
}

export function writeXlsx(wb: Workbook): void {
  for (const sheet of wb.sheets) {
    if (!sheet.doc || !sheet.sheetData) continue;
    // A formula change inside a shared group would leave the other members'
    // @si dangling, so rewrite the whole group as plain formulas (de-share).
    const dirtySi = new Set<string>();
    for (const cell of sheet.cells.values()) if (cell.sharedSi != null && cell.fDirty) dirtySi.add(cell.sharedSi);
    let touched = false;
    for (const cell of sheet.cells.values()) {
      const deshare = cell.sharedSi != null && dirtySi.has(cell.sharedSi);
      if (cell.edited || cell.recomputed || deshare) {
        writeXlsxCell(sheet, cell, deshare);
        if (deshare) cell.sharedSi = undefined;
        touched = true;
      }
    }
    if ((touched || sheet.layoutDirty) && sheet.path) wb.files[sheet.path] = serializeXml(sheet.doc);
  }
  if (wb.stylesDirty && wb.stylesDoc) wb.files["xl/styles.xml"] = serializeXml(wb.stylesDoc);
}

