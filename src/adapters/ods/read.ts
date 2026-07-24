import type { Cell, CellKind, CellStyle, Phonetic, Sheet, Workbook } from "../../core/model";
import { formatNumber, key, noteExtent, numToStr, parseXml, parseXmlOpt } from "../../core/model";
import { durationToSerial, isoToSerial } from "../../core/dates";
import { readOdsCharts } from "./chart-read";
import { REPEAT_CAP, odfToA1, odsBorderColor, odsCellComments, odsCellLink, odsCellRich, odsCellText, odsColorOf, odsLenToPx } from "./shared";
// ---------------------------------------------------------------------------
// ods read: content.xml parsing (tables, rows, styles)
// ---------------------------------------------------------------------------

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
      if (cp.getAttribute("fo:wrap-option") === "wrap") s.wrap = true;
      const va = cp.getAttribute("style:vertical-align");
      if (va === "top" || va === "middle" || va === "bottom") s.valign = va;
    }
    const tp = el.getElementsByTagName("style:text-properties")[0];
    if (tp) {
      if (tp.getAttribute("fo:font-weight") === "bold") s.bold = true;
      if (tp.getAttribute("fo:font-style") === "italic") s.italic = true;
      const us = tp.getAttribute("style:text-underline-style");
      if (us && us !== "none") {
        s.underline = true;
        // ODF underline style/type -> CSS text-decoration-style (double via text-underline-type).
        const flavour =
          tp.getAttribute("style:text-underline-type") === "double"
            ? "double"
            : us === "dotted"
              ? "dotted"
              : /dash/.test(us)
                ? "dashed"
                : us === "wave"
                  ? "wavy"
                  : undefined;
        if (flavour) s.underlineStyle = flavour;
      }
      const ls = tp.getAttribute("style:text-line-through-style");
      if (ls && ls !== "none") s.strike = true;
      const fs = parseFloat(tp.getAttribute("fo:font-size") ?? "");
      if (fs && (tp.getAttribute("fo:font-size") ?? "").endsWith("pt")) s.fontSize = fs;
      const ff = tp.getAttribute("fo:font-family") ?? tp.getAttribute("style:font-name");
      if (ff) s.fontFamily = ff.replace(/^['"]|['"]$/g, "");
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

// Frozen panes live in settings.xml (ODF view settings), keyed by sheet name. SplitMode 2 =
// "frozen" (1 = split, 0 = none); the SplitPosition is the count of frozen columns / rows.
function readOdsFreeze(files: Record<string, Uint8Array>): Map<string, { rows: number; cols: number }> {
  const out = new Map<string, { rows: number; cols: number }>();
  const f = files["settings.xml"];
  if (!f) return out;
  const doc = parseXmlOpt(f);
  if (!doc) return out;
  const tablesMaps = Array.from(doc.getElementsByTagName("config:config-item-map-named")).filter(
    (m) => m.getAttribute("config:name") === "Tables",
  );
  for (const tables of tablesMaps) {
    for (const entry of Array.from(tables.children)) {
      if (entry.localName !== "config-item-map-entry") continue;
      const name = entry.getAttribute("config:name");
      if (!name) continue;
      const item = (key: string): number => {
        for (const ci of Array.from(entry.children))
          if (ci.localName === "config-item" && ci.getAttribute("config:name") === key) return Number(ci.textContent || "0");
        return 0;
      };
      const cols = item("HorizontalSplitMode") === 2 ? Math.max(0, Math.floor(item("HorizontalSplitPosition"))) : 0;
      const rows = item("VerticalSplitMode") === 2 ? Math.max(0, Math.floor(item("VerticalSplitPosition"))) : 0;
      if (rows > 0 || cols > 0) out.set(name, { rows, cols });
    }
  }
  return out;
}

// One ODF cell address ("$Sheet1.$A$1" or ".$A$1") -> A1 ("Sheet1!A1" or "A1").
function odfRefToA1(part: string): string {
  const dot = part.indexOf(".");
  const sheet = dot > 0 ? part.slice(0, dot).replace(/^\$/, "").replace(/^'|'$/g, "") : "";
  const cell = part.slice(dot + 1).replace(/\$/g, "");
  return sheet ? `${sheet}!${cell}` : cell;
}
// An ODF cell-range-address -> an A1 reference ("Sheet1!A1:B2"); the end cell drops its sheet.
function odfAddrToA1(addr: string): string {
  const parts = addr.split(":");
  const first = odfRefToA1(parts[0]!);
  if (parts.length === 1) return first;
  const second = odfRefToA1(parts[1]!);
  return `${first}:${second.includes("!") ? second.split("!")[1] : second}`;
}

export function readOds(files: Record<string, Uint8Array>): Workbook {
  const contentFile = files["content.xml"];
  if (!contentFile) throw new Error("not an .ods: content.xml missing");
  const contentDoc = parseXml(contentFile);
  const docs = [contentDoc];
  const odsStylesDoc = files["styles.xml"] ? parseXmlOpt(files["styles.xml"]) : undefined;
  if (odsStylesDoc) docs.push(odsStylesDoc);
  const styles = parseOdsStyles(docs);
  const freezeByName = readOdsFreeze(files);
  const wb: Workbook = { kind: "ods", sheets: [], files, contentDoc, contentPath: "content.xml" };
  // Defined names: <table:named-range table:name="X" table:cell-range-address="$Sheet1.$A$1:.$B$2"/>.
  // Convert the ODF address to an A1 reference ("Sheet1!A1:B2") for recalc.
  const definedNames = new Map<string, string>();
  for (const nr of Array.from(contentDoc.getElementsByTagName("table:named-range"))) {
    const name = nr.getAttribute("table:name");
    const addr = nr.getAttribute("table:cell-range-address");
    if (name && addr) definedNames.set(name, odfAddrToA1(addr));
  }
  if (definedNames.size) wb.definedNames = definedNames;
  for (const table of Array.from(contentDoc.getElementsByTagName("table:table"))) {
    const name = table.getAttribute("table:name") ?? `Sheet${wb.sheets.length + 1}`;
    const sheet: Sheet = { name, cells: new Map(), maxRow: 0, maxCol: 0, tableEl: table };
    const fz = freezeByName.get(name);
    if (fz) sheet.freeze = fz;
    readOdsTable(sheet, table, styles);
    wb.sheets.push(sheet);
  }
  readOdsCharts(wb, files);
  return wb;
}

export function readOdsTable(sheet: Sheet, table: Element, styles: OdsStyles): void {
  // Column widths + hidden state: walk <table:table-column> (each may repeat) and map to px.
  const cols = new Map<number, number>();
  const hiddenCols = new Set<number>();
  const isHidden = (el: Element): boolean => {
    const v = el.getAttribute("table:visibility");
    return v === "collapse" || v === "filter";
  };
  let colIdx = 0;
  const collectCols = (parent: Element) => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-column") {
        const rep = Math.max(1, Number(ch.getAttribute("table:number-columns-repeated") || "1"));
        const w = styles.colW.get(ch.getAttribute("table:style-name") ?? "");
        const hidden = isHidden(ch);
        for (let i = 0; i < Math.min(rep, REPEAT_CAP); i++) {
          colIdx++;
          if (w) cols.set(colIdx, w);
          if (hidden) hiddenCols.add(colIdx);
        }
      } else if (ch.localName === "table-header-columns" || ch.localName === "table-columns") {
        collectCols(ch);
      }
    }
  };
  collectCols(table);
  if (cols.size) sheet.colWidths = cols;
  if (hiddenCols.size) sheet.hiddenCols = hiddenCols;

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
  const hiddenRows = new Set<number>();
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
    const rowHidden = isHidden(rowEl);
    for (let k = 0; k < copies; k++) {
      const r = rowNum + 1 + k;
      if (rh) rowHeights.set(r, rh);
      if (rowHidden) hiddenRows.add(r);
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
  if (hiddenRows.size) sheet.hiddenRows = hiddenRows;
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
    let numFmt: string | undefined;
    let odsValueType: string | undefined;
    let odsCurrency: string | undefined;
    let phonetic: Phonetic[] | undefined;
    if (valueType === "float" || valueType === "percentage" || valueType === "currency") {
      value = cellEl.getAttribute("office:value") ?? text;
      // ODF stores the producer's formatted text in <text:p>; use it as the display.
      if (text !== "" && text !== value) display = text;
      kind = "n";
      if (valueType !== "float") {
        odsValueType = valueType;
        odsCurrency = cellEl.getAttribute("office:currency") ?? undefined;
        if (valueType === "percentage") numFmt = "0.00%";
      }
    } else if (valueType === "boolean") {
      value = cellEl.getAttribute("office:boolean-value") === "true" ? "TRUE" : "FALSE";
      kind = "b";
    } else if (valueType === "string") {
      const rich = odsCellRich(cellEl); // base text + furigana (text:ruby), not the reading folded in
      value = cellEl.getAttribute("office:string-value") ?? rich.text;
      phonetic = rich.phonetic;
      kind = "s";
    } else if (valueType === "date") {
      // Model dates as serials so arithmetic and re-formatting work; the
      // producer's formatted <text:p> stays as the display.
      const iso = cellEl.getAttribute("office:date-value") ?? "";
      const serial = iso === "" ? null : isoToSerial(iso);
      if (serial != null) {
        value = numToStr(serial);
        kind = "n";
        odsValueType = "date";
        numFmt = iso.includes("T") ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd";
      } else {
        value = iso || text;
        kind = value === "" ? "blank" : "s";
      }
      if (text !== "" && text !== value) display = text;
    } else if (valueType === "time") {
      const dur = cellEl.getAttribute("office:time-value") ?? "";
      const serial = dur === "" ? null : durationToSerial(dur);
      if (serial != null) {
        value = numToStr(serial);
        kind = "n";
        odsValueType = "time";
        numFmt = "hh:mm:ss";
      } else {
        value = dur || text;
        kind = value === "" ? "blank" : "s";
      }
      if (text !== "" && text !== value) display = text;
    } else {
      const rich = odsCellRich(cellEl);
      value = rich.text;
      phonetic = rich.phonetic;
      kind = value === "" ? "blank" : "s";
    }
    const link = odsCellLink(cellEl);
    const comments = odsCellComments(cellEl);
    const has = value !== "" || formulaRaw != null || style != null || link != null || comments != null;
    if (!has) {
      out.push({ has: false, span: crep, startCol });
      continue;
    }
    const cell: Cell = {
      row: 0,
      col: startCol,
      value,
      kind,
      display: display ?? (kind === "n" && numFmt != null ? formatNumber(numFmt, value) ?? undefined : undefined),
      numFmt,
      odsValueType,
      odsCurrency,
      formula: formulaRaw ? odfToA1(formulaRaw) : undefined,
      odfFormula: formulaRaw,
      style,
      cellStyle: style ? styles.cell.get(style) : undefined,
      el: cellEl,
      phonetic,
      link,
      comments,
    };
    out.push({ has: true, span: Math.min(crep, REPEAT_CAP), startCol, colSpan, rowSpan, cell });
  }
  return out;
}

