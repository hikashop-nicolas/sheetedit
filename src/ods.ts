import type { Cell, CellKind, CellStyle, Sheet, StyleChange, Workbook } from "./model";
import { getCell, key, noteExtent, parseXml, parseXmlOpt, serializeXml } from "./model";
import { mergeCellStyle, xmlOf } from "./xlsx";
// ---------------------------------------------------------------------------
// ods (ODF spreadsheet)
// ---------------------------------------------------------------------------

export const ODS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
};
export const REPEAT_CAP = 1024;

/** Replace text outside single-quoted string literals. */
export function replaceOutsideStrings(s: string, fn: (chunk: string) => string): string {
  let out = "";
  let i = 0;
  while (i < s.length) {
    const q = s.indexOf('"', i);
    const qq = s.indexOf("'", i);
    let next = -1;
    let quote = '"';
    if (q === -1 && qq === -1) next = -1;
    else if (q === -1) {
      next = qq;
      quote = "'";
    } else if (qq === -1) {
      next = q;
      quote = '"';
    } else if (q < qq) {
      next = q;
      quote = '"';
    } else {
      next = qq;
      quote = "'";
    }
    if (next === -1) {
      out += fn(s.slice(i));
      break;
    }
    out += fn(s.slice(i, next));
    const end = s.indexOf(quote, next + 1);
    if (end === -1) {
      out += s.slice(next);
      break;
    }
    out += s.slice(next, end + 1);
    i = end + 1;
  }
  return out;
}

/** ODF formula (`of:=[.A1]+[.B1]`) -> A1 (`A1+B1`). */
export function odfToA1(odf: string): string {
  let core = odf.replace(/^of:=/, "").replace(/^=/, "");
  core = core.replace(/\[([^\]]*)\]/g, (_, inner: string) => {
    if (!inner || inner.includes("#")) return inner.replace(/\./g, ""); // #REF! etc.
    const parts = inner.split(":");
    const mapped = parts.map((part) => {
      const dot = part.lastIndexOf(".");
      const sheet = dot >= 0 ? part.slice(0, dot) : "";
      const ref = dot >= 0 ? part.slice(dot + 1) : part;
      return { sheet, ref };
    });
    const sheet = mapped[0]!.sheet;
    const cells = mapped.map((m) => m.ref).join(":");
    return (sheet ? sheet + "!" : "") + cells;
  });
  return replaceOutsideStrings(core, (chunk) => chunk.replace(/;/g, ","));
}

/** A1 (`A1+B1`) -> ODF formula (`of:=[.A1]+[.B1]`). Used only for user-typed formulas. */
export function a1ToOdf(a1: string): string {
  const refRe =
    /(?<![A-Za-z0-9_.$])(?:('[^']*'|[A-Za-z_][A-Za-z0-9_]*)!)?(\$?[A-Z]{1,3}\$?[0-9]+)(?::(\$?[A-Z]{1,3}\$?[0-9]+))?(?![A-Za-z0-9_(])/g;
  const converted = replaceOutsideStrings(a1, (chunk) => {
    const semi = chunk.replace(/,/g, ";");
    return semi.replace(refRe, (_m, sheet: string | undefined, c1: string, c2?: string) => {
      const sh = sheet ?? "";
      const range = c2 ? `${c1}:.${c2}` : c1;
      return `[${sh}.${range}]`;
    });
  });
  return "of:=" + converted;
}

export function odsCellText(cell: Element): string {
  return Array.from(cell.getElementsByTagName("text:p"))
    .map((p) => p.textContent ?? "")
    .join("\n");
}

// Convert an ODF length ("2.5cm", "96pt", "1in", "0.45mm", "12px") to CSS px.
export function odsLenToPx(len: string | null | undefined): number | undefined {
  if (!len) return undefined;
  const m = /^([\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(len.trim());
  if (!m) return undefined;
  const v = parseFloat(m[1]!);
  switch (m[2]) {
    case "cm":
      return Math.round((v * 96) / 2.54);
    case "mm":
      return Math.round((v * 96) / 25.4);
    case "in":
      return Math.round(v * 96);
    case "pt":
      return Math.round((v * 96) / 72);
    case "pc":
      return Math.round(v * 16);
    default:
      return Math.round(v); // px or unitless
  }
}

export const odsColorOf = (v: string | null): string | undefined =>
  v && v !== "transparent" && v !== "none" ? (/^#[0-9a-fA-F]{6}$/.test(v) ? v.toLowerCase() : v) : undefined;

// A border value is like "0.5pt solid #000000" or "none"; keep the colour if it draws.
export function odsBorderColor(v: string | null): string | undefined {
  if (!v || v === "none" || /(^|\s)0(\.0+)?(pt|cm|mm|in|px)?\s/.test(" " + v + " ")) return undefined;
  const m = /#[0-9a-fA-F]{6}/.exec(v);
  return m ? m[0].toLowerCase() : "#000000";
}

export interface OdsStyles {
  cell: Map<string, CellStyle>; // family table-cell -> resolved style
  colW: Map<string, number>; // family table-column -> width px
  rowH: Map<string, number>; // family table-row -> height px
}

// Parse <style:style> from the given docs (content.xml automatic styles + styles.xml),
// resolving table-cell parent chains, and the column/row dimension styles.
export function parseOdsStyles(docs: Document[]): OdsStyles {
  const raw = new Map<string, { el: Element; parent?: string }>();
  const colW = new Map<string, number>();
  const rowH = new Map<string, number>();
  for (const doc of docs) {
    for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
      const name = st.getAttribute("style:name");
      if (!name) continue;
      const family = st.getAttribute("style:family");
      if (family === "table-cell") {
        raw.set(name, { el: st, parent: st.getAttribute("style:parent-style-name") || undefined });
      } else if (family === "table-column") {
        const p = st.getElementsByTagName("style:table-column-properties")[0];
        const w = odsLenToPx(p?.getAttribute("style:column-width"));
        if (w) colW.set(name, w);
      } else if (family === "table-row") {
        const p = st.getElementsByTagName("style:table-row-properties")[0];
        const h = odsLenToPx(p?.getAttribute("style:row-height"));
        if (h) rowH.set(name, h);
      }
    }
  }
  const ownStyle = (el: Element): CellStyle => {
    const s: CellStyle = {};
    const cp = el.getElementsByTagName("style:table-cell-properties")[0];
    if (cp) {
      const bg = odsColorOf(cp.getAttribute("fo:background-color"));
      if (bg) s.bg = bg;
      const all = cp.getAttribute("fo:border");
      const sides: CellStyle["borders"] = {};
      const top = odsBorderColor(cp.getAttribute("fo:border-top") ?? all);
      const right = odsBorderColor(cp.getAttribute("fo:border-right") ?? all);
      const bottom = odsBorderColor(cp.getAttribute("fo:border-bottom") ?? all);
      const left = odsBorderColor(cp.getAttribute("fo:border-left") ?? all);
      if (top) sides.top = top;
      if (right) sides.right = right;
      if (bottom) sides.bottom = bottom;
      if (left) sides.left = left;
      if (top || right || bottom || left) s.borders = sides;
    }
    const tp = el.getElementsByTagName("style:text-properties")[0];
    if (tp) {
      if (tp.getAttribute("fo:font-weight") === "bold") s.bold = true;
      if (tp.getAttribute("fo:font-style") === "italic") s.italic = true;
      const col = odsColorOf(tp.getAttribute("fo:color"));
      if (col) s.color = col;
    }
    const pp = el.getElementsByTagName("style:paragraph-properties")[0];
    const ta = pp?.getAttribute("fo:text-align");
    if (ta === "center") s.align = "center";
    else if (ta === "end" || ta === "right") s.align = "right";
    else if (ta === "start" || ta === "left") s.align = "left";
    return s;
  };
  const cell = new Map<string, CellStyle>();
  const resolve = (name: string, depth = 0): CellStyle => {
    const cached = cell.get(name);
    if (cached) return cached;
    const entry = raw.get(name);
    if (!entry || depth > 8) return {};
    const base = entry.parent ? resolve(entry.parent, depth + 1) : {};
    const merged = { ...base, ...ownStyle(entry.el) };
    cell.set(name, merged);
    return merged;
  };
  for (const name of raw.keys()) resolve(name);
  return { cell, colW, rowH };
}

export function readOds(files: Record<string, Uint8Array>): Workbook {
  const contentFile = files["content.xml"];
  if (!contentFile) throw new Error("not an .ods: content.xml missing");
  const contentDoc = parseXml(contentFile);
  const docs = [contentDoc];
  const odsStylesDoc = files["styles.xml"] ? parseXmlOpt(files["styles.xml"]) : undefined;
  if (odsStylesDoc) docs.push(odsStylesDoc);
  const styles = parseOdsStyles(docs);
  const wb: Workbook = { kind: "ods", sheets: [], files, contentDoc, contentPath: "content.xml" };
  for (const table of Array.from(contentDoc.getElementsByTagName("table:table"))) {
    const name = table.getAttribute("table:name") ?? `Sheet${wb.sheets.length + 1}`;
    const sheet: Sheet = { name, cells: new Map(), maxRow: 0, maxCol: 0, tableEl: table };
    readOdsTable(sheet, table, styles);
    wb.sheets.push(sheet);
  }
  return wb;
}

export function readOdsTable(sheet: Sheet, table: Element, styles: OdsStyles): void {
  // Column widths: walk <table:table-column> (each may repeat) and map to px.
  const cols = new Map<number, number>();
  let colIdx = 0;
  const collectCols = (parent: Element) => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-column") {
        const rep = Math.max(1, Number(ch.getAttribute("table:number-columns-repeated") || "1"));
        const w = styles.colW.get(ch.getAttribute("table:style-name") ?? "");
        for (let i = 0; i < Math.min(rep, REPEAT_CAP); i++) {
          colIdx++;
          if (w) cols.set(colIdx, w);
        }
      } else if (ch.localName === "table-header-columns" || ch.localName === "table-columns") {
        collectCols(ch);
      }
    }
  };
  collectCols(table);
  if (cols.size) sheet.colWidths = cols;

  let rowNum = 0;
  const rows: { el: Element; header: boolean }[] = [];
  let headerGroupEl: Element | undefined;
  const collect = (parent: Element, header = false) => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-row") rows.push({ el: ch, header });
      else if (ch.localName === "table-header-rows" || ch.localName === "table-rows") {
        if (ch.localName === "table-header-rows") headerGroupEl = ch;
        collect(ch, header || ch.localName === "table-header-rows");
      }
    }
  };
  collect(table);
  const rowHeights = new Map<number, number>();
  const rowStyles = new Map<number, string>();
  const rowEls = new Map<number, Element>();
  const rowRuns: { from: number; to: number; el: Element }[] = [];
  const coveredEls = new Map<string, Element>();
  const merges: { r1: number; c1: number; r2: number; c2: number }[] = [];
  let headerFrom = 0;
  let headerTo = 0;
  for (const { el: rowEl, header } of rows) {
    const rrep = Math.max(1, Number(rowEl.getAttribute("table:number-rows-repeated") || "1"));
    const rowStyle = rowEl.getAttribute("table:style-name") ?? undefined;
    const rh = styles.rowH.get(rowStyle ?? "");
    const parsedCells = parseOdsRow(rowEl, styles);
    const rowHasContent = parsedCells.some((c) => c.has);
    const copies = rowHasContent ? Math.min(rrep, REPEAT_CAP) : 0;
    if (header) {
      if (!headerFrom) headerFrom = rowNum + 1;
      headerTo = Math.max(headerTo, rowNum + rrep);
    }
    // Keep the original row element per expanded index so its attributes
    // (visibility, default-cell-style-name, ...) survive a rebuild.
    const worthKeeping = Array.from(rowEl.attributes).some((a) => a.name !== "table:number-rows-repeated");
    if (worthKeeping) for (let k = 0; k < Math.min(rrep, REPEAT_CAP); k++) rowEls.set(rowNum + 1 + k, rowEl);
    // A content run repeated beyond the cap: the un-expanded tail is preserved verbatim.
    if (rowHasContent && rrep > REPEAT_CAP) rowRuns.push({ from: rowNum + 1 + REPEAT_CAP, to: rowNum + rrep, el: rowEl });
    for (let k = 0; k < copies; k++) {
      const r = rowNum + 1 + k;
      if (rh) rowHeights.set(r, rh);
      if (rowStyle) rowStyles.set(r, rowStyle);
      for (const pc of parsedCells) {
        if (!pc.has) {
          // A covered (merged-away) cell may still carry content/attributes; keep it.
          if (pc.coveredEl) for (let j = 0; j < pc.span; j++) coveredEls.set(key(r, pc.startCol + j), pc.coveredEl);
          continue;
        }
        const c = pc.cell!;
        if ((pc.colSpan ?? 1) > 1 || (pc.rowSpan ?? 1) > 1) {
          merges.push({ r1: r, c1: c.col, r2: r + (pc.rowSpan ?? 1) - 1, c2: c.col + (pc.colSpan ?? 1) - 1 });
        }
        for (let j = 0; j < pc.span; j++) {
          sheet.cells.set(key(r, c.col + j), { ...c, row: r, col: c.col + j });
          noteExtent(sheet, r, c.col + j);
        }
      }
    }
    rowNum += rrep;
  }
  if (rowHeights.size) sheet.rowHeights = rowHeights;
  if (rowStyles.size) sheet.odsRowStyles = rowStyles;
  if (rowEls.size) sheet.odsRowEls = rowEls;
  if (rowRuns.length) sheet.odsRowRuns = rowRuns;
  if (coveredEls.size) sheet.odsCoveredEls = coveredEls;
  if (headerGroupEl && headerFrom) sheet.odsHeaderRows = { el: headerGroupEl, from: headerFrom, to: headerTo };
  if (merges.length) sheet.merges = merges;
}

export interface ParsedOdsCell {
  has: boolean;
  span: number;
  startCol: number;
  colSpan?: number;
  rowSpan?: number;
  cell?: Cell;
  /** A covered-table-cell that carries content or attributes, preserved on save. */
  coveredEl?: Element;
}

export function parseOdsRow(rowEl: Element, styles: OdsStyles): ParsedOdsCell[] {
  const out: ParsedOdsCell[] = [];
  let col = 0;
  for (const cellEl of Array.from(rowEl.children)) {
    const local = cellEl.localName;
    if (local !== "table-cell" && local !== "covered-table-cell") continue;
    const crep = Math.max(1, Number(cellEl.getAttribute("table:number-columns-repeated") || "1"));
    const startCol = col + 1;
    col += crep;
    if (local === "covered-table-cell") {
      // Merged-away cell: keep it when it carries anything beyond the repeat count.
      const worth = cellEl.children.length > 0 || Array.from(cellEl.attributes).some((a) => a.name !== "table:number-columns-repeated");
      out.push({ has: false, span: Math.min(crep, REPEAT_CAP), startCol, coveredEl: worth ? cellEl : undefined });
      continue;
    }
    const colSpan = Math.max(1, Number(cellEl.getAttribute("table:number-columns-spanned") || "1"));
    const rowSpan = Math.max(1, Number(cellEl.getAttribute("table:number-rows-spanned") || "1"));
    const valueType = cellEl.getAttribute("office:value-type");
    const formulaRaw = cellEl.getAttribute("table:formula") ?? undefined;
    const style = cellEl.getAttribute("table:style-name") ?? undefined;
    const text = odsCellText(cellEl);
    let value = "";
    let kind: CellKind = "blank";
    let display: string | undefined;
    if (valueType === "float" || valueType === "percentage" || valueType === "currency") {
      value = cellEl.getAttribute("office:value") ?? text;
      // ODF stores the producer's formatted text in <text:p>; use it as the display.
      if (text !== "" && text !== value) display = text;
      kind = "n";
    } else if (valueType === "boolean") {
      value = cellEl.getAttribute("office:boolean-value") === "true" ? "TRUE" : "FALSE";
      kind = "b";
    } else if (valueType === "string") {
      value = cellEl.getAttribute("office:string-value") ?? odsCellText(cellEl);
      kind = "s";
    } else if (valueType === "date") {
      value = cellEl.getAttribute("office:date-value") ?? text;
      if (text !== "" && text !== value) display = text;
      kind = "s";
    } else if (valueType === "time") {
      value = cellEl.getAttribute("office:time-value") ?? text;
      if (text !== "" && text !== value) display = text;
      kind = "s";
    } else {
      value = odsCellText(cellEl);
      kind = value === "" ? "blank" : "s";
    }
    const has = value !== "" || formulaRaw != null || style != null;
    if (!has) {
      out.push({ has: false, span: crep, startCol });
      continue;
    }
    const cell: Cell = {
      row: 0,
      col: startCol,
      value,
      kind,
      display,
      formula: formulaRaw ? odfToA1(formulaRaw) : undefined,
      odfFormula: formulaRaw,
      style,
      cellStyle: style ? styles.cell.get(style) : undefined,
      el: cellEl,
    };
    out.push({ has: true, span: Math.min(crep, REPEAT_CAP), startCol, colSpan, rowSpan, cell });
  }
  return out;
}

export function makeOdsCell(doc: Document, cell: Cell, edited: boolean): Element {
  // Untouched cell: clone the original verbatim (preserves dates, formats, rich text).
  if (cell.el && !cell.edited && !cell.recomputed) {
    const clone = cell.el.cloneNode(true) as Element;
    clone.removeAttribute("table:number-columns-repeated");
    clone.removeAttribute("table:number-rows-repeated");
    return clone;
  }
  const c = doc.createElementNS(ODS.table, "table:table-cell");
  if (cell.style) c.setAttributeNS(ODS.table, "table:style-name", cell.style);
  const formulaToWrite = edited && cell.formula != null ? a1ToOdf(cell.formula) : cell.odfFormula;
  if (formulaToWrite) c.setAttributeNS(ODS.table, "table:formula", formulaToWrite);
  const addText = (text: string) => {
    if (text === "") return;
    const p = doc.createElementNS(ODS.text, "text:p");
    p.textContent = text;
    c.appendChild(p);
  };
  if (cell.kind === "n") {
    c.setAttributeNS(ODS.office, "office:value-type", "float");
    c.setAttributeNS(ODS.office, "office:value", cell.value);
    addText(cell.value);
  } else if (cell.kind === "b") {
    c.setAttributeNS(ODS.office, "office:value-type", "boolean");
    c.setAttributeNS(ODS.office, "office:boolean-value", cell.value === "TRUE" ? "true" : "false");
    addText(cell.value);
  } else if (cell.kind === "s" || cell.kind === "e") {
    c.setAttributeNS(ODS.office, "office:value-type", "string");
    c.setAttributeNS(ODS.office, "office:string-value", cell.value);
    addText(cell.value);
  }
  return c;
}

// --- ods style write-back -------------------------------------------------

export function ensureOdsAutoStyles(doc: Document): Element {
  let el = doc.getElementsByTagName("office:automatic-styles")[0] as Element | undefined;
  if (!el) {
    el = doc.createElementNS(ODS.office, "office:automatic-styles");
    const body = doc.getElementsByTagName("office:body")[0];
    doc.documentElement.insertBefore(el, body ?? null);
  }
  return el;
}

export function findOdsStyleByName(doc: Document, name: string): Element | undefined {
  for (const s of Array.from(doc.getElementsByTagName("style:style")))
    if (s.getAttribute("style:name") === name) return s;
  return undefined;
}

// Add a built style element to <office:automatic-styles>, reusing an existing one with
// the same family + serialized properties. Returns the (existing or new) style name.
export function internOdsStyle(doc: Document, autoStyles: Element, family: string, prefix: string, styleEl: Element): string {
  styleEl.setAttributeNS(ODS.style, "style:family", family);
  const sig = Array.from(styleEl.children).map(xmlOf).join("");
  for (const ex of Array.from(autoStyles.children)) {
    if (ex.localName !== "style" || ex.getAttribute("style:family") !== family) continue;
    if (Array.from(ex.children).map(xmlOf).join("") === sig) return ex.getAttribute("style:name")!;
  }
  const used = new Set(Array.from(doc.getElementsByTagName("style:style")).map((s) => s.getAttribute("style:name")));
  let n = 1;
  while (used.has(prefix + n)) n++;
  const name = prefix + n;
  styleEl.setAttributeNS(ODS.style, "style:name", name);
  autoStyles.appendChild(styleEl);
  return name;
}

export const odsSetOrRemove = (el: Element, qn: string, v: string | undefined) => {
  if (v == null) el.removeAttribute(qn);
  else el.setAttributeNS(ODS.fo, qn, v);
};

// Apply a resolved CellStyle onto an ods cell style element (cloned from the original
// so number formats / parents survive), creating the property children as needed.
export function applyCellStyleToOds(doc: Document, st: Element, cs: CellStyle): void {
  const child = (tag: string): Element => {
    const ex = st.getElementsByTagName(tag)[0];
    if (ex) return ex;
    const el = doc.createElementNS(ODS.style, tag);
    st.appendChild(el);
    return el;
  };
  const cp = child("style:table-cell-properties");
  odsSetOrRemove(cp, "fo:background-color", cs.bg);
  cp.removeAttribute("fo:border"); // use per-side so partial borders are exact
  const bv = (c?: string) => (c ? `0.5pt solid ${c}` : undefined);
  odsSetOrRemove(cp, "fo:border-top", bv(cs.borders?.top));
  odsSetOrRemove(cp, "fo:border-right", bv(cs.borders?.right));
  odsSetOrRemove(cp, "fo:border-bottom", bv(cs.borders?.bottom));
  odsSetOrRemove(cp, "fo:border-left", bv(cs.borders?.left));
  const tp = child("style:text-properties");
  odsSetOrRemove(tp, "fo:font-weight", cs.bold ? "bold" : undefined);
  odsSetOrRemove(tp, "fo:font-style", cs.italic ? "italic" : undefined);
  odsSetOrRemove(tp, "fo:color", cs.color);
  const pp = child("style:paragraph-properties");
  odsSetOrRemove(pp, "fo:text-align", cs.align === "center" ? "center" : cs.align === "right" ? "end" : cs.align === "left" ? "start" : undefined);
  // Drop property children that ended up empty so dedup stays tight.
  for (const el of [cp, tp, pp]) if (el.attributes.length === 0 && el.children.length === 0) st.removeChild(el);
}

export function setOdsCellStyle(wb: Workbook, _sheet: Sheet, cell: Cell, change: StyleChange): void {
  const doc = wb.contentDoc;
  if (!doc) return;
  const autoStyles = ensureOdsAutoStyles(doc);
  const desired = mergeCellStyle(cell.cellStyle ?? {}, change);
  const orig = cell.style ? findOdsStyleByName(doc, cell.style) : undefined;
  const st = orig
    ? (orig.cloneNode(true) as Element)
    : doc.createElementNS(ODS.style, "style:style");
  st.removeAttribute("style:name");
  applyCellStyleToOds(doc, st, desired);
  cell.style = internOdsStyle(doc, autoStyles, "table-cell", "ce", st);
  cell.cellStyle = desired;
  cell.edited = true;
}

// Build (or reuse) a table-column style of the given width and return its name.
export function odsColStyle(doc: Document, autoStyles: Element, px: number): string {
  const st = doc.createElementNS(ODS.style, "style:style");
  const p = doc.createElementNS(ODS.style, "style:table-column-properties");
  p.setAttributeNS(ODS.fo, "fo:break-before", "auto");
  p.setAttributeNS(ODS.style, "style:column-width", `${(px / 96) * 2.54}cm`);
  st.appendChild(p);
  return internOdsStyle(doc, autoStyles, "table-column", "co", st);
}

// Set one column's width (px), splitting the <table:table-column> run that covers it.
export function setOdsColWidth(wb: Workbook, sheet: Sheet, col: number, px: number): void {
  (sheet.colWidths ??= new Map()).set(col, px);
  const doc = wb.contentDoc;
  const table = sheet.tableEl;
  if (!doc || !table) return;
  const styleName = odsColStyle(doc, ensureOdsAutoStyles(doc), px);
  // Walk the column elements, tracking the running column index, and split the run at `col`.
  let idx = 0;
  for (const ch of Array.from(table.children)) {
    if (ch.localName !== "table-column") continue;
    const rep = Math.max(1, Number(ch.getAttribute("table:number-columns-repeated") || "1"));
    const start = idx + 1,
      end = idx + rep;
    idx = end;
    if (col < start || col > end) continue;
    const mk = (from: number, to: number, style?: string) => {
      const c = ch.cloneNode(false) as Element;
      const n = to - from + 1;
      if (n > 1) c.setAttributeNS(ODS.table, "table:number-columns-repeated", String(n));
      else c.removeAttribute("table:number-columns-repeated");
      if (style) c.setAttributeNS(ODS.table, "table:style-name", style);
      return c;
    };
    const parent = ch.parentNode!;
    if (start < col) parent.insertBefore(mk(start, col - 1), ch);
    parent.insertBefore(mk(col, col, styleName), ch);
    if (end > col) parent.insertBefore(mk(col + 1, end), ch);
    parent.removeChild(ch);
    return;
  }
}

// Set one row's height (px) by giving its row a row style (re-emitted by writeOds).
export function setOdsRowHeight(wb: Workbook, sheet: Sheet, row: number, px: number): void {
  (sheet.rowHeights ??= new Map()).set(row, px);
  const doc = wb.contentDoc;
  if (!doc) return;
  const st = doc.createElementNS(ODS.style, "style:style");
  const p = doc.createElementNS(ODS.style, "style:table-row-properties");
  p.setAttributeNS(ODS.style, "style:row-height", `${(px / 96) * 2.54}cm`);
  p.setAttributeNS(ODS.fo, "fo:break-before", "auto");
  st.appendChild(p);
  const name = internOdsStyle(doc, ensureOdsAutoStyles(doc), "table-row", "ro", st);
  (sheet.odsRowStyles ??= new Map()).set(row, name);
  sheet.odsDirty = true;
}

// Add or remove a merged range; writeOds emits the spans and covered cells.
export function setOdsMerge(sheet: Sheet, r1: number, c1: number, r2: number, c2: number, merge: boolean): void {
  sheet.odsDirty = true;
  const top = Math.min(r1, r2),
    left = Math.min(c1, c2),
    bottom = Math.max(r1, r2),
    right = Math.max(c1, c2);
  const merges = (sheet.merges ??= []);
  const idx = merges.findIndex((m) => m.r1 === top && m.c1 === left && m.r2 === bottom && m.c2 === right);
  if (merge) {
    if (idx === -1) merges.push({ r1: top, c1: left, r2: bottom, c2: right });
    // Mark covered cells edited so writeOds regenerates the row with covered-table-cells.
    for (let r = top; r <= bottom; r++)
      for (let c = left; c <= right; c++) {
        const cell = sheet.cells.get(key(r, c));
        if (cell) cell.edited = true;
      }
    const tl = sheet.cells.get(key(top, left));
    if (tl) tl.edited = true;
  } else if (idx !== -1) {
    merges.splice(idx, 1);
  }
  const tl = sheet.cells.get(key(top, left));
  if (tl) tl.edited = true;
}

export function writeOds(wb: Workbook): void {
  const doc = wb.contentDoc!;
  for (const sheet of wb.sheets) {
    const table = sheet.tableEl;
    if (!table) continue;
    // An untouched sheet keeps its original XML verbatim (repeats, header groups,
    // row attributes, covered cells); only touched sheets are re-emitted.
    let cellsDirty = false;
    for (const cell of sheet.cells.values())
      if (cell.edited || cell.recomputed) {
        cellsDirty = true;
        break;
      }
    if (!cellsDirty && !sheet.odsDirty) continue;
    // preserve structural children (column definitions etc.), drop existing rows
    const keep: Element[] = [];
    for (const ch of Array.from(table.children)) {
      if (ch.localName !== "table-row" && ch.localName !== "table-header-rows" && ch.localName !== "table-rows") {
        keep.push(ch);
      }
    }
    while (table.firstChild) table.removeChild(table.firstChild);
    for (const k of keep) table.appendChild(k);
    // Merge bookkeeping: covered positions and the span at each top-left.
    const covered = new Set<string>();
    const spanAt = new Map<string, { cs: number; rs: number }>();
    let mergeMaxCol = 0;
    for (const m of sheet.merges ?? []) {
      spanAt.set(key(m.r1, m.c1), { cs: m.c2 - m.c1 + 1, rs: m.r2 - m.r1 + 1 });
      mergeMaxCol = Math.max(mergeMaxCol, m.c2);
      for (let r = m.r1; r <= m.r2; r++)
        for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(key(r, c));
    }
    const maxRow = Math.max(1, sheet.maxRow);
    const maxCol = Math.max(1, sheet.maxCol, mergeMaxCol);
    // Repeated content runs beyond the expansion cap, re-emitted verbatim.
    const runAt = new Map<number, { to: number; el: Element }>();
    let lastRow = maxRow;
    for (const rr of sheet.odsRowRuns ?? []) {
      runAt.set(rr.from, { to: rr.to, el: rr.el });
      lastRow = Math.max(lastRow, rr.to);
    }
    // Original header-rows group: its rows are re-wrapped in a fresh group element.
    const hdr = sheet.odsHeaderRows;
    const headerEl = hdr ? (hdr.el.cloneNode(false) as Element) : null;
    const appendRow = (rowEl: Element, r: number) => {
      if (hdr && headerEl && r >= hdr.from && r <= hdr.to) {
        if (!headerEl.parentNode) table.appendChild(headerEl);
        headerEl.appendChild(rowEl);
      } else table.appendChild(rowEl);
    };
    for (let r = 1; r <= lastRow; r++) {
      const run = runAt.get(r);
      if (run) {
        // The preserved tail of a repeated content run: one clone with its repeat count.
        const clone = run.el.cloneNode(true) as Element;
        const n = run.to - r + 1;
        if (n > 1) clone.setAttributeNS(ODS.table, "table:number-rows-repeated", String(n));
        else clone.removeAttribute("table:number-rows-repeated");
        appendRow(clone, r);
        r = run.to;
        continue;
      }
      if (r > maxRow) continue; // gaps past the content extent existed only as empty runs
      const orig = sheet.odsRowEls?.get(r);
      const rowEl = orig ? (orig.cloneNode(false) as Element) : doc.createElementNS(ODS.table, "table:table-row");
      rowEl.removeAttribute("table:number-rows-repeated");
      const rowStyle = sheet.odsRowStyles?.get(r);
      if (rowStyle) rowEl.setAttributeNS(ODS.table, "table:style-name", rowStyle);
      // The last column needing an explicit cell: content or a merge edge on this row.
      let lastContent = 0;
      for (let c = maxCol; c >= 1; c--) {
        if (getCell(sheet, r, c) || covered.has(key(r, c)) || spanAt.has(key(r, c))) {
          lastContent = c;
          break;
        }
      }
      for (let c = 1; c <= lastContent; c++) {
        if (covered.has(key(r, c))) {
          const origCov = sheet.odsCoveredEls?.get(key(r, c));
          const cov = origCov ? (origCov.cloneNode(true) as Element) : doc.createElementNS(ODS.table, "table:covered-table-cell");
          cov.removeAttribute("table:number-columns-repeated");
          rowEl.appendChild(cov);
          continue;
        }
        const cell = getCell(sheet, r, c);
        const el = cell ? makeOdsCell(doc, cell, !!cell.edited) : doc.createElementNS(ODS.table, "table:table-cell");
        const span = spanAt.get(key(r, c));
        if (span) {
          if (span.cs > 1) el.setAttributeNS(ODS.table, "table:number-columns-spanned", String(span.cs));
          if (span.rs > 1) el.setAttributeNS(ODS.table, "table:number-rows-spanned", String(span.rs));
        }
        rowEl.appendChild(el);
      }
      if (lastContent < maxCol) {
        const filler = doc.createElementNS(ODS.table, "table:table-cell");
        filler.setAttributeNS(ODS.table, "table:number-columns-repeated", String(maxCol - lastContent));
        rowEl.appendChild(filler);
      }
      appendRow(rowEl, r);
    }
  }
  wb.files["content.xml"] = serializeXml(doc);
}

