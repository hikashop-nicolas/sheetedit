import type { Cell, Phonetic, Sheet, Workbook } from "../../core/model";
import { ensureCell, getCell, key, serializeXml } from "../../core/model";
import { isDateFmt, isTimeOnlyFmt, serialToDuration, serialToIso } from "../../core/dates";
import { ODS, a1ToOdf } from "./shared";
import { ensureOdsAutoStyles, findOdsStyleByName, internOdsStyle, odsColStyle } from "./styles";
// ---------------------------------------------------------------------------
// ods write: cell/row emission and the save pass
// ---------------------------------------------------------------------------

/** Any pending write for a cell (value edit, recalc, or authored link/note/validation). */
export function odsCellDirty(cell: Cell): boolean {
  return !!(cell.edited || cell.recomputed || cell.linkDirty || cell.commentsDirty || cell.dvDirty);
}

export function makeOdsCell(doc: Document, cell: Cell, edited: boolean): Element {
  // Untouched cell: clone the original verbatim (preserves dates, formats, rich text, notes).
  if (cell.el && !odsCellDirty(cell)) {
    const clone = cell.el.cloneNode(true) as Element;
    clone.removeAttribute("table:number-columns-repeated");
    clone.removeAttribute("table:number-rows-repeated");
    return clone;
  }
  // Touched cell that has an original element: patch a clone so everything we did not explicitly
  // change (note position/formatting, extra links, unmodelled structure) is preserved.
  if (cell.el) return patchOdsCell(doc, cell);
  // Brand-new cell: build from the model.
  return buildOdsCellFresh(doc, cell, edited);
}

// Remove the value <text:p> children (leaving <office:annotation> etc.) then re-emit the value
// attributes and text, wrapping in a hyperlink and applying furigana as needed.
function applyOdsValue(doc: Document, c: Element, cell: Cell): void {
  for (const ch of Array.from(c.children)) if (ch.localName === "p") c.removeChild(ch);
  for (const a of ["value", "value-type", "date-value", "time-value", "boolean-value", "string-value", "currency"]) c.removeAttributeNS(ODS.office, a);
  const addText = (text: string) => {
    if (text === "") return;
    const p = doc.createElementNS(ODS.text, "text:p");
    if (cell.link) {
      const a = doc.createElementNS(ODS.text, "text:a");
      a.setAttributeNS(ODS.xlink, "xlink:href", cell.link.internal ? `#${cell.link.href.replace("!", ".")}` : cell.link.href);
      a.setAttributeNS(ODS.xlink, "xlink:type", "simple");
      a.textContent = text;
      p.appendChild(a);
    } else {
      p.textContent = text;
    }
    c.appendChild(p);
  };
  if (cell.kind === "n") {
    // Keep the cell's original ODF type (date/time/percentage/currency), or the type a typed date
    // adopted, so an edit does not degrade the cell to a float.
    const vt = cell.odsValueType ?? (isDateFmt(cell.numFmt) ? (isTimeOnlyFmt(cell.numFmt) ? "time" : "date") : undefined);
    const serial = Number(cell.value);
    if (vt === "date" && Number.isFinite(serial) && serialToIso(serial) != null) {
      c.setAttributeNS(ODS.office, "office:value-type", "date");
      c.setAttributeNS(ODS.office, "office:date-value", serialToIso(serial)!);
    } else if (vt === "time" && Number.isFinite(serial)) {
      c.setAttributeNS(ODS.office, "office:value-type", "time");
      c.setAttributeNS(ODS.office, "office:time-value", serialToDuration(serial));
    } else if (vt === "percentage" || vt === "currency") {
      c.setAttributeNS(ODS.office, "office:value-type", vt);
      if (vt === "currency" && cell.odsCurrency) c.setAttributeNS(ODS.office, "office:currency", cell.odsCurrency);
      c.setAttributeNS(ODS.office, "office:value", cell.value);
    } else {
      c.setAttributeNS(ODS.office, "office:value-type", "float");
      c.setAttributeNS(ODS.office, "office:value", cell.value);
    }
    addText(cell.display ?? cell.value);
  } else if (cell.kind === "b") {
    c.setAttributeNS(ODS.office, "office:value-type", "boolean");
    c.setAttributeNS(ODS.office, "office:boolean-value", cell.value === "TRUE" ? "true" : "false");
    addText(cell.value);
  } else if (cell.kind === "s" || cell.kind === "e") {
    c.setAttributeNS(ODS.office, "office:value-type", "string");
    // A linked cell must NOT carry office:string-value: LibreOffice treats it as authoritative and
    // discards the rich <text:a>, so the value lives only in the text:p (the anchor text).
    if (!cell.link) c.setAttributeNS(ODS.office, "office:string-value", cell.value);
    if (cell.phonetic?.length) c.appendChild(makeRubyP(doc, cell.value, cell.phonetic));
    else addText(cell.value);
  }
}

// Add / replace / remove a single note, preserving the first annotation's position + creator/date.
function patchOdsAnnotations(doc: Document, c: Element, cell: Cell): void {
  const anns = Array.from(c.children).filter((x) => x.localName === "annotation");
  if (cell.comments?.length) {
    const cm = cell.comments[0]!;
    let ann = anns[0];
    if (!ann) { ann = doc.createElementNS(ODS.office, "office:annotation"); c.insertBefore(ann, c.firstChild); }
    for (const p of Array.from(ann.children).filter((x) => x.localName === "p")) ann.removeChild(p);
    if (cm.author && !ann.getElementsByTagName("dc:creator").length) { const cr = doc.createElementNS(ODS.dc, "dc:creator"); cr.textContent = cm.author; ann.appendChild(cr); }
    if (!ann.getElementsByTagName("dc:date").length) { const dt = doc.createElementNS(ODS.dc, "dc:date"); dt.textContent = new Date().toISOString().slice(0, 19); ann.appendChild(dt); }
    const ap = doc.createElementNS(ODS.text, "text:p"); ap.textContent = cm.text; ann.appendChild(ap);
    for (let i = 1; i < anns.length; i++) c.removeChild(anns[i]!);
  } else {
    for (const ann of anns) c.removeChild(ann);
  }
}

function patchOdsCell(doc: Document, cell: Cell): Element {
  const c = cell.el!.cloneNode(true) as Element;
  c.removeAttribute("table:number-columns-repeated");
  c.removeAttribute("table:number-rows-repeated");
  if (cell.style) c.setAttributeNS(ODS.table, "table:style-name", cell.style);
  else c.removeAttributeNS(ODS.table, "style-name");
  if (cell.dvDirty) {
    if (cell.odsValidationName) c.setAttributeNS(ODS.table, "table:content-validation-name", cell.odsValidationName);
    else c.removeAttributeNS(ODS.table, "content-validation-name");
  }
  if (cell.commentsDirty) patchOdsAnnotations(doc, c, cell);
  if (cell.edited || cell.recomputed || cell.linkDirty) {
    if (cell.edited) {
      if (cell.formula != null) c.setAttributeNS(ODS.table, "table:formula", a1ToOdf(cell.formula));
      else c.removeAttributeNS(ODS.table, "formula");
    }
    applyOdsValue(doc, c, cell);
  }
  return c;
}

function buildOdsCellFresh(doc: Document, cell: Cell, edited: boolean): Element {
  const c = doc.createElementNS(ODS.table, "table:table-cell");
  if (cell.style) c.setAttributeNS(ODS.table, "table:style-name", cell.style);
  if (cell.odsValidationName) c.setAttributeNS(ODS.table, "table:content-validation-name", cell.odsValidationName);
  const formulaToWrite = edited && cell.formula != null ? a1ToOdf(cell.formula) : cell.odfFormula;
  if (formulaToWrite) c.setAttributeNS(ODS.table, "table:formula", formulaToWrite);
  // A note is an <office:annotation> child, emitted before the value text.
  for (const cm of cell.comments ?? []) {
    const an = doc.createElementNS(ODS.office, "office:annotation");
    if (cm.author) { const cr = doc.createElementNS(ODS.dc, "dc:creator"); cr.textContent = cm.author; an.appendChild(cr); }
    const dt = doc.createElementNS(ODS.dc, "dc:date"); dt.textContent = new Date().toISOString().slice(0, 19); an.appendChild(dt);
    const ap = doc.createElementNS(ODS.text, "text:p"); ap.textContent = cm.text; an.appendChild(ap);
    c.appendChild(an);
  }
  applyOdsValue(doc, c, cell);
  return c;
}

// A <text:p> that renders base text with furigana as <text:ruby> over base[sb..eb).
function makeRubyP(doc: Document, base: string, runs: Phonetic[]): Element {
  const p = doc.createElementNS(ODS.text, "text:p");
  let pos = 0;
  for (const ph of [...runs].filter((x) => x.reading).sort((a, b) => a.sb - b.sb)) {
    const sb = Math.max(pos, Math.min(base.length, ph.sb));
    const eb = Math.max(sb, Math.min(base.length, ph.eb || base.length));
    if (sb > pos) p.appendChild(doc.createTextNode(base.slice(pos, sb)));
    const ruby = doc.createElementNS(ODS.text, "text:ruby");
    const rb = doc.createElementNS(ODS.text, "text:ruby-base");
    rb.textContent = base.slice(sb, eb);
    const rt = doc.createElementNS(ODS.text, "text:ruby-text");
    rt.textContent = ph.reading;
    ruby.appendChild(rb);
    ruby.appendChild(rt);
    p.appendChild(ruby);
    pos = eb;
  }
  if (pos < base.length) p.appendChild(doc.createTextNode(base.slice(pos)));
  return p;
}

// --- ods cell-content authoring (hyperlinks, notes) -----------------------
// These set a per-aspect dirty flag (not `edited`), so makeOdsCell patches the original cell
// element in place: only the link / note / validation is changed, everything else (value, other
// notes' position + formatting, unmodelled structure) is preserved.

/** Set or (link === null) remove a cell hyperlink. */
export function setOdsHyperlink(sheet: Sheet, r: number, c: number, link: Cell["link"] | null): void {
  const cell = ensureCell(sheet, r, c);
  cell.link = link ?? undefined;
  cell.linkDirty = true;
}

/** Add/replace or (text === null) remove a single note on a cell. */
export function setOdsComment(sheet: Sheet, r: number, c: number, text: string | null, author = "sheetedit"): void {
  const cell = ensureCell(sheet, r, c);
  cell.comments = text ? [{ author, text }] : undefined;
  cell.commentsDirty = true;
}

// An A1 range ("Sheet2!A1:A9" / "A1:A9") -> an ODF list address "[Sheet2.A1:.A9]" / "[.A1:.A9]".
function a1RangeToOdfAddr(ref: string): string {
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(ref);
  const sheetName = m ? (m[1] ?? m[2]) : "";
  const body = (m ? m[3]! : ref).replace(/\$/g, "");
  const [a, b] = body.split(":");
  const first = sheetName ? `${sheetName}.${a}` : `.${a}`;
  return b ? `[${first}:.${b}]` : `[${first}]`;
}

/** Add or (spec === null) remove a list data validation over the given ranges (1-based inclusive). */
export function setOdsDataValidation(
  wb: Workbook,
  sheet: Sheet,
  ranges: { r1: number; c1: number; r2: number; c2: number }[],
  spec: { values?: string[]; rangeRef?: string; allowBlank?: boolean } | null,
): void {
  const doc = wb.contentDoc;
  const inRange = (r: number, c: number): boolean => ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2);
  const sameRanges = (v: NonNullable<Sheet["validations"]>[number]): boolean =>
    JSON.stringify(v.ranges) === JSON.stringify(ranges);
  sheet.validations = (sheet.validations ?? []).filter((v) => !sameRanges(v));

  if (!spec) {
    for (const cell of sheet.cells.values())
      if (cell.odsValidationName && inRange(cell.row, cell.col)) { cell.odsValidationName = undefined; cell.dvDirty = true; }
    if (!sheet.validations.length) sheet.validations = undefined;
    return;
  }

  sheet.validations.push({ ranges, values: spec.values, rangeRef: spec.rangeRef, allowBlank: spec.allowBlank });
  if (!doc) return;
  const spreadsheet = doc.getElementsByTagName("office:spreadsheet")[0];
  if (!spreadsheet) return;
  // Unique name across existing definitions.
  const used = new Set(Array.from(doc.getElementsByTagName("table:content-validation")).map((e) => e.getAttribute("table:name")));
  let n = 1;
  while (used.has(`val${n}`)) n++;
  const name = `val${n}`;
  // Condition: an inline quoted list (ODF escapes an inner quote as ""), or a cell range.
  const cond = spec.rangeRef
    ? `of:cell-content-is-in-list(${a1RangeToOdfAddr(spec.rangeRef)})`
    : `of:cell-content-is-in-list(${(spec.values ?? []).map((v) => `"${v.replace(/"/g, '""')}"`).join(";")})`;
  let container = spreadsheet.getElementsByTagName("table:content-validations")[0];
  if (!container) {
    container = doc.createElementNS(ODS.table, "table:content-validations");
    const firstTable = spreadsheet.getElementsByTagName("table:table")[0];
    spreadsheet.insertBefore(container, firstTable ?? spreadsheet.firstChild);
  }
  const cv = doc.createElementNS(ODS.table, "table:content-validation");
  cv.setAttributeNS(ODS.table, "table:name", name);
  cv.setAttributeNS(ODS.table, "table:condition", cond);
  cv.setAttributeNS(ODS.table, "table:allow-empty-cell", spec.allowBlank ? "true" : "false");
  cv.setAttributeNS(ODS.table, "table:display-list", "unsorted");
  container.appendChild(cv);
  // Tag every cell in the range so makeOdsCell re-emits the reference.
  for (const g of ranges)
    for (let r = g.r1; r <= g.r2; r++)
      for (let c = g.c1; c <= g.c2; c++) { const cell = ensureCell(sheet, r, c); cell.odsValidationName = name; cell.dvDirty = true; }
}

// An A1 range -> an ODF conditional-format target address "Sheet1.A1:Sheet1.A5".
function a1RangeToOdfTarget(sheetName: string, g: { r1: number; c1: number; r2: number; c2: number }): string {
  const a1 = (r: number, c: number): string => {
    let s = "", n = c;
    while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); }
    return `${s}${r}`;
  };
  return `${sheetName}.${a1(g.r1, g.c1)}:${sheetName}.${a1(g.r2, g.c2)}`;
}

type OdsCfSpec =
  | { kind: "cellIs"; operator: string; value: string; fill: string }
  | { kind: "colorScale"; colors: string[] }
  | { kind: "dataBar"; color: string };

// A cellIs operator + operand -> a standard ODF style:condition ("cell-content()>5").
function styleMapCondition(operator: string, value: string): string {
  const op: Record<string, string> = { greaterThan: ">", lessThan: "<", equal: "=", notEqual: "!=", greaterThanOrEqual: ">=", lessThanOrEqual: "<=" };
  return `cell-content()${op[operator] ?? "="}${value}`;
}

/** Add or (spec === null) remove a conditional format over the given ranges. */
export function setOdsCondFormat(
  wb: Workbook,
  sheet: Sheet,
  ranges: { r1: number; c1: number; r2: number; c2: number }[],
  spec: OdsCfSpec | null,
): void {
  const doc = wb.contentDoc;
  const inRange = (r: number, c: number): boolean => ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2);
  const target = ranges.map((g) => a1RangeToOdfTarget(sheet.name, g)).join(" ");
  const sameRanges = (cf: NonNullable<Sheet["condFormats"]>[number]): boolean => JSON.stringify(cf.ranges) === JSON.stringify(ranges);
  sheet.condFormats = (sheet.condFormats ?? []).filter((cf) => !sameRanges(cf));
  const spreadsheet = doc?.getElementsByTagName("office:spreadsheet")[0];
  // Remove any calcext target we previously wrote for these ranges.
  if (doc) for (const cf of Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "conditional-format"))
    if ((cf.getAttribute("calcext:target-range-address") ?? cf.getAttribute("target-range-address")) === target) cf.parentNode?.removeChild(cf);
  // Revert any style:map-derived cell styles we previously applied over these ranges.
  if (doc) for (const cell of sheet.cells.values())
    if (inRange(cell.row, cell.col) && cell.style?.startsWith("ceCond")) {
      const parent = findOdsStyleByName(doc, cell.style)?.getAttribute("style:parent-style-name");
      cell.style = parent && parent !== "Default" ? parent : undefined;
      cell.edited = true;
    }

  if (spec && doc && spreadsheet) {
    const priority = 1 + Math.max(0, ...sheet.condFormats.flatMap((cf) => cf.rules.map((r) => r.priority)));
    const rule: import("../../core/model").CfRule = { type: spec.kind, priority };
    if (spec.kind === "cellIs") {
      // Standard ODF: an applied (fill) style referenced by a <style:map> on each cell's base style.
      const autoStyles = ensureOdsAutoStyles(doc);
      const fillSt = doc.createElementNS(ODS.style, "style:style");
      const cp = doc.createElementNS(ODS.style, "style:table-cell-properties");
      cp.setAttributeNS(ODS.fo, "fo:background-color", spec.fill);
      fillSt.appendChild(cp);
      const fillName = internOdsStyle(doc, autoStyles, "table-cell", "ceCFfill", fillSt);
      const cond = styleMapCondition(spec.operator, spec.value);
      const base = a1RangeToOdfTarget(sheet.name, ranges[0]!).split(":")[0]!;
      // One derived base style per distinct original style, inheriting it (parent) plus the map.
      const derived = new Map<string, string>();
      const mkDerived = (orig: string): string => {
        const existing = derived.get(orig);
        if (existing) return existing;
        const st = doc.createElementNS(ODS.style, "style:style");
        st.setAttributeNS(ODS.style, "style:family", "table-cell");
        st.setAttributeNS(ODS.style, "style:parent-style-name", orig || "Default");
        const map = doc.createElementNS(ODS.style, "style:map");
        map.setAttributeNS(ODS.style, "style:condition", cond);
        map.setAttributeNS(ODS.style, "style:apply-style-name", fillName);
        map.setAttributeNS(ODS.style, "style:base-cell-address", base);
        st.appendChild(map);
        const used = new Set(Array.from(doc.getElementsByTagName("style:style")).map((s) => s.getAttribute("style:name")));
        let n = 1; while (used.has(`ceCond${n}`)) n++;
        const name = `ceCond${n}`;
        st.setAttributeNS(ODS.style, "style:name", name);
        autoStyles.appendChild(st);
        derived.set(orig, name);
        return name;
      };
      for (const g of ranges)
        for (let r = g.r1; r <= g.r2; r++)
          for (let c = g.c1; c <= g.c2; c++) { const cell = ensureCell(sheet, r, c); cell.style = mkDerived(cell.style ?? ""); cell.edited = true; }
      rule.operator = spec.operator; rule.formulas = [spec.value]; rule.dxf = { bg: spec.fill };
    } else {
      // Colour scales / data bars have no standard ODF form; write LibreOffice's calcext.
      let container = Array.from(spreadsheet.children).find((e) => e.localName === "conditional-formats");
      if (!container) { container = doc.createElementNS(ODS.calcext, "calcext:conditional-formats"); spreadsheet.appendChild(container); }
      const cfEl = doc.createElementNS(ODS.calcext, "calcext:conditional-format");
      cfEl.setAttributeNS(ODS.calcext, "calcext:target-range-address", target);
      if (spec.kind === "colorScale") {
        const cs = doc.createElementNS(ODS.calcext, "calcext:color-scale");
        const types = spec.colors.length >= 3 ? ["minimum", "percentile", "maximum"] : ["minimum", "maximum"];
        types.forEach((ty, i) => { const e = doc.createElementNS(ODS.calcext, "calcext:color-scale-entry"); e.setAttributeNS(ODS.calcext, "calcext:type", ty); if (ty === "percentile") e.setAttributeNS(ODS.calcext, "calcext:value", "50"); e.setAttributeNS(ODS.calcext, "calcext:color", spec.colors[i]!); cs.appendChild(e); });
        cfEl.appendChild(cs);
        rule.colorScale = { cfvo: types.map((ty) => ({ type: ty === "minimum" ? "min" : ty === "maximum" ? "max" : "percentile", val: ty === "percentile" ? 50 : undefined })), colors: spec.colors };
      } else {
        const db = doc.createElementNS(ODS.calcext, "calcext:data-bar");
        db.setAttributeNS(ODS.calcext, "calcext:positive-color", spec.color);
        for (const ty of ["minimum", "maximum"]) { const e = doc.createElementNS(ODS.calcext, "calcext:formatting-entry"); e.setAttributeNS(ODS.calcext, "calcext:type", ty); db.appendChild(e); }
        cfEl.appendChild(db);
        rule.dataBar = { color: spec.color, min: { type: "min" }, max: { type: "max" } };
      }
      container.appendChild(cfEl);
    }
    sheet.condFormats.push({ ranges, rules: [rule] });
  }
  if (!sheet.condFormats?.length) sheet.condFormats = undefined;
  sheet.odsDirty = true;
}

// --- ods style write-back -------------------------------------------------


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
      if (odsCellDirty(cell)) {
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
