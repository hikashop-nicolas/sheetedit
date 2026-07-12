import type { Cell, CellKind, CellStyle, Sheet, Workbook } from "../../core/model";
import { firstByLocal, formatNumber, key, noteExtent, numToStr, parseA1Ref, parseXml, parseXmlOpt, shiftFormula } from "../../core/model";
import { isDateFmt, isoToSerial } from "../../core/dates";
// ---------------------------------------------------------------------------
// xlsx read: workbook/worksheet parsing, style pools resolution
// ---------------------------------------------------------------------------

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
    underline: !!firstByLocal(f, "u"),
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
      };
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
