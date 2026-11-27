// An XlsBook as the xlsx parts sheetedit's xlsx reader understands, so a .xls opens with every
// xlsx editing path. Style indexes are kept 1:1 (BIFF XF n is cellXfs n), and palette colours
// carry their index so the BIFF writer can map them back exactly.
import { strToU8 } from "fflate";
import { colToLetters } from "../../core/model";
import { isDateFmt } from "../../core/dates";
import { colWidthToPx } from "../xlsx/read";
import { rangeA1 } from "./read";
import { icvCss } from "./palette";
import type { XlsBook, XlsCell, XlsFont, XlsRange, XlsSheet, XlsXf } from "./types";

const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export const esc = (s: string): string =>
  s.replace(/[&<>"]/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : "&quot;"))
    // XML 1.0 cannot carry most control characters; encode them the OOXML way.
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, (c) => `_x${c.charCodeAt(0).toString(16).padStart(4, "0").toUpperCase()}_`);

const PATTERNS = ["none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"];
export const BORDER_STYLES = ["", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"];
const H_ALIGN = ["", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"];
const V_ALIGN = ["top", "center", "", "justify", "distributed"];
const CF_OPS = ["", "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"];
const DV_TYPES = ["none", "whole", "decimal", "list", "date", "time", "textLength", "custom"];
const DV_OPS = ["between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"];
const DV_ERR = ["stop", "warning", "information"];

/** A palette colour element: rgb for the xlsx reader plus the index for the BIFF writer. */
function colorXml(tag: string, book: XlsBook, icv: number, auto: "omit" | "black" = "omit"): string {
  const css = icvCss(book.palette, icv);
  if (!css) return auto === "black" ? `<${tag} rgb="FF000000" auto="1"/>` : "";
  return `<${tag} rgb="FF${css.slice(1).toUpperCase()}" indexed="${icv}"/>`;
}

function fontXml(f: XlsFont, book: XlsBook): string {
  let s = "<font>";
  if (f.weight >= 700) s += "<b/>";
  if (f.italic) s += "<i/>";
  if (f.strike) s += "<strike/>";
  if (f.underline) s += f.underline === 2 ? '<u val="double"/>' : f.underline === 0x21 ? '<u val="singleAccounting"/>' : f.underline === 0x22 ? '<u val="doubleAccounting"/>' : "<u/>";
  if (f.escapement === 1) s += '<vertAlign val="superscript"/>';
  else if (f.escapement === 2) s += '<vertAlign val="subscript"/>';
  s += `<sz val="${f.height / 20}"/>`;
  s += colorXml("color", book, f.color);
  s += `<name val="${esc(f.name || "Arial")}"/></font>`;
  return s;
}

function runPrXml(f: XlsFont | undefined, book: XlsBook): string {
  if (!f) return "";
  let s = "<rPr>";
  if (f.weight >= 700) s += "<b/>";
  if (f.italic) s += "<i/>";
  if (f.strike) s += "<strike/>";
  if (f.underline) s += f.underline === 2 ? '<u val="double"/>' : "<u/>";
  if (f.escapement === 1) s += '<vertAlign val="superscript"/>';
  else if (f.escapement === 2) s += '<vertAlign val="subscript"/>';
  s += `<sz val="${f.height / 20}"/>`;
  s += colorXml("color", book, f.color);
  s += `<rFont val="${esc(f.name || "Arial")}"/></rPr>`;
  return s;
}

/** BIFF font index (index 4 does not exist) to a position in the FONT list. */
export const fontPos = (ifnt: number): number => (ifnt < 4 ? ifnt : ifnt - 1);

interface StylesOut { xml: string; fmtOf: (xf: number) => string | number | undefined }

function stylesXml(book: XlsBook, dxfs: string[]): StylesOut {
  const fmtId = (id: number): number => (book.biff < 5 ? 164 + id : id);
  const numFmts = [...book.formats].map(([id, code]) => `<numFmt numFmtId="${fmtId(id)}" formatCode="${esc(code)}"/>`);
  const fonts = (book.fonts.length ? book.fonts : [{ height: 200, italic: false, strike: false, weight: 400, underline: 0, escapement: 0, color: 0x7fff, name: "Arial" }]).map((f) => fontXml(f, book));
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders = ["<border><left/><right/><top/><bottom/><diagonal/></border>"];
  const intern = (pool: string[], xml: string): number => {
    const i = pool.indexOf(xml);
    if (i >= 0) return i;
    pool.push(xml);
    return pool.length - 1;
  };
  const xfs = (book.xfs.length ? book.xfs : []).map((xf) => xfXml(xf));
  function xfXml(xf: XlsXf): string {
    let fill = 0;
    if (xf.pattern) {
      const type = PATTERNS[xf.pattern] ?? "solid";
      fill = intern(fills, `<fill><patternFill patternType="${type}">${colorXml("fgColor", book, xf.fore, "black")}${xf.pattern > 1 ? colorXml("bgColor", book, xf.back) : ""}</patternFill></fill>`);
    }
    const side = (tag: string, b: { style: number; color: number }): string =>
      b.style ? `<${tag} style="${BORDER_STYLES[b.style] ?? "thin"}">${colorXml("color", book, b.color, "black")}</${tag}>` : `<${tag}/>`;
    let border = 0;
    if (xf.left.style || xf.right.style || xf.top.style || xf.bottom.style || xf.diag) {
      const d = xf.diag;
      border = intern(borders, `<border${d?.up ? ' diagonalUp="1"' : ""}${d?.down ? ' diagonalDown="1"' : ""}>${side("left", xf.left)}${side("right", xf.right)}${side("top", xf.top)}${side("bottom", xf.bottom)}${d ? side("diagonal", d) : "<diagonal/>"}</border>`);
    }
    const font = Math.min(fontPos(xf.font), fonts.length - 1);
    const al: string[] = [];
    if (H_ALIGN[xf.hAlign]) al.push(`horizontal="${H_ALIGN[xf.hAlign]}"`);
    if (V_ALIGN[xf.vAlign]) al.push(`vertical="${V_ALIGN[xf.vAlign]}"`);
    if (xf.wrap) al.push('wrapText="1"');
    if (xf.rot) al.push(`textRotation="${xf.rot}"`);
    if (xf.indent) al.push(`indent="${xf.indent}"`);
    if (xf.shrink) al.push('shrinkToFit="1"');
    const prot = !xf.locked || xf.hidden ? `<protection${xf.locked ? "" : ' locked="0"'}${xf.hidden ? ' hidden="1"' : ""}/>` : "";
    return `<xf numFmtId="${xf.fmt ? fmtId(xf.fmt) : 0}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${xf.fmt ? ' applyNumberFormat="1"' : ""}${al.length ? ' applyAlignment="1"' : ""}${prot ? ' applyProtection="1"' : ""}>${al.length ? `<alignment ${al.join(" ")}/>` : ""}${prot}</xf>`;
  }
  if (!xfs.length) xfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
  const xml = HEAD + `<styleSheet xmlns="${SS}">` +
    (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join("")}</numFmts>` : "") +
    `<fonts count="${fonts.length}">${fonts.join("")}</fonts>` +
    `<fills count="${fills.length}">${fills.join("")}</fills>` +
    `<borders count="${borders.length}">${borders.join("")}</borders>` +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    `<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>` +
    "</styleSheet>";
  const fmtOf = (xf: number): string | number | undefined => {
    const id = book.xfs[xf]?.fmt;
    if (!id) return undefined;
    return book.formats.get(id) ?? (book.biff < 5 ? undefined : id);
  };
  return { xml, fmtOf };
}

const num = (v: number): string => (Number.isFinite(v) ? String(v) : "0");
const a1 = (r: number, c: number): string => colToLetters(c + 1) + String(r + 1);
const sqref = (rs: XlsRange[]): string => rs.map(rangeA1).join(" ");
const twipsPt = (t: number): number => Math.round((t / 20) * 100) / 100;

/** A sheet's column widths and row heights in px (0-based), as the drawing anchors measure them. */
export function xlsSheetPx(s: XlsSheet): { colPx: (c: number) => number; rowPx: (r: number) => number } {
  const colPx = (c: number): number => {
    const ci = s.cols.find((x) => c >= x.first && c <= x.last);
    if (ci?.hidden) return 0;
    if (ci) return colWidthToPx(ci.width / 256, 7);
    const defW = s.stdWidth != null ? s.stdWidth / 256 : (s.defColWidth ?? 8) + 0.71;
    return colWidthToPx(defW, 7);
  };
  const rowPx = (r: number): number => {
    const info = s.rows.get(r);
    if (info?.hidden) return 0;
    return Math.round((info?.height ?? s.defRowHeight ?? 255) / 15);
  };
  return { colPx, rowPx };
}

/** Convert the book into xlsx parts. */
export function xlsBookToXlsx(book: XlsBook): Record<string, Uint8Array> {
  const files: Record<string, Uint8Array> = {};
  const ct: string[] = [];
  const shared: string[] = [];
  const sharedIdx = new Map<string, number>();
  const dxfs: string[] = [];
  const addShared = (xml: string): number => {
    let i = sharedIdx.get(xml);
    if (i == null) { i = shared.length; shared.push(xml); sharedIdx.set(xml, i); }
    return i;
  };
  const fontAt = (ifnt: number): XlsFont | undefined => book.fonts[fontPos(ifnt)];
  const fmtOfXf = (xf: number): string | number | undefined => {
    const id = book.xfs[xf]?.fmt;
    if (!id) return undefined;
    return book.formats.get(id) ?? (book.biff < 5 ? undefined : id);
  };
  const dateXf = new Map<number, boolean>();
  const isDate = (xf: number): boolean => {
    let v = dateXf.get(xf);
    if (v == null) { v = isDateFmt(fmtOfXf(xf)); dateXf.set(xf, v); }
    return v;
  };
  const sheetRels: string[][] = [];
  let drawingN = 0, mediaN = 0, commentsN = 0;
  book.sheets.forEach((sh, si) => {
    const rels: string[] = [];
    const relId = (type: string, target: string, external = false): string => {
      const id = `rId${rels.length + 1}`;
      rels.push(`<Relationship Id="${id}" Type="${R}/${type}" Target="${esc(target)}"${external ? ' TargetMode="External"' : ""}/>`);
      return id;
    };
    files[`xl/worksheets/sheet${si + 1}.xml`] = strToU8(sheetXml(sh, si));
    sheetRels.push(rels);
    if (rels.length) files[`xl/worksheets/_rels/sheet${si + 1}.xml.rels`] = strToU8(HEAD + `<Relationships xmlns="${PKG}">${rels.join("")}</Relationships>`);

    function sheetXml(s: XlsSheet, index: number): string {
      const w2 = s.window2 ?? 0x6b6;
      let maxR = 0, maxC = 0;
      const rowsCells = new Map<number, XlsCell[]>();
      for (const cell of s.cells.values()) {
        let list = rowsCells.get(cell.r);
        if (!list) rowsCells.set(cell.r, (list = []));
        list.push(cell);
        if (cell.r > maxR) maxR = cell.r;
        if (cell.c > maxC) maxC = cell.c;
      }
      const defH = s.defRowHeight ?? 255;
      const rowNums = new Set<number>([...rowsCells.keys()]);
      for (const [r, info] of s.rows) if (info.hidden || info.level || info.collapsed || (info.custom && info.height != null) || (info.height != null && info.height !== defH)) rowNums.add(r);
      let rowsXml = "";
      for (const r of [...rowNums].sort((a, b) => a - b)) {
        if (r > 0xffff) continue;
        const info = s.rows.get(r);
        let attrs = ` r="${r + 1}"`;
        if (info?.height != null && (info.custom || info.height !== defH)) attrs += ` ht="${twipsPt(info.height)}" customHeight="1"`;
        if (info?.hidden) attrs += ' hidden="1"';
        if (info?.level) attrs += ` outlineLevel="${info.level}"`;
        if (info?.collapsed) attrs += ' collapsed="1"';
        if (info?.xf != null && info.xf !== 15 && info.xf < book.xfs.length) attrs += ` s="${info.xf}" customFormat="1"`;
        const cells = (rowsCells.get(r) ?? []).sort((a, b) => a.c - b.c).map(cellXml).join("");
        rowsXml += `<row${attrs}>${cells}</row>`;
      }
      const cols = s.cols
        .filter((c) => c.first <= c.last)
        .sort((a, b) => a.first - b.first)
        .map((c) => `<col min="${c.first + 1}" max="${c.last + 1}" width="${num(c.width / 256)}"${c.width ? ' customWidth="1"' : ""}${c.hidden ? ' hidden="1"' : ""}${c.level ? ` outlineLevel="${c.level}"` : ""}${c.collapsed ? ' collapsed="1"' : ""}${c.xf !== 15 && c.xf < book.xfs.length ? ` style="${c.xf}"` : ""}/>`)
        .join("");
      const sheetPr: string[] = [];
      if (s.tabColor != null) { const css = icvCss(book.palette, s.tabColor); if (css) sheetPr.push(`<tabColor rgb="FF${css.slice(1).toUpperCase()}" indexed="${s.tabColor}"/>`); }
      if (s.summaryBelow === false || s.summaryRight === false) sheetPr.push(`<outlinePr summaryBelow="${s.summaryBelow === false ? 0 : 1}" summaryRight="${s.summaryRight === false ? 0 : 1}"/>`);
      if (s.print.fitToPage) sheetPr.push('<pageSetUpPr fitToPage="1"/>');
      let pane = "";
      if (s.pane && (s.pane.x || s.pane.y)) {
        const frozen = (w2 & 0x8) !== 0;
        const p = s.pane;
        const tl = a1(p.top, p.left);
        const active = p.x && p.y ? "bottomRight" : p.y ? "bottomLeft" : "topRight";
        pane = frozen
          ? `<pane${p.x ? ` xSplit="${p.x}"` : ""}${p.y ? ` ySplit="${p.y}"` : ""} topLeftCell="${tl}" activePane="${active}" state="frozen"/>`
          : `<pane${p.x ? ` xSplit="${p.x}"` : ""}${p.y ? ` ySplit="${p.y}"` : ""} topLeftCell="${tl}" activePane="${active}"/>`;
      }
      const view = `<sheetViews><sheetView${w2 & 0x2 ? "" : ' showGridLines="0"'}${w2 & 0x10 ? "" : ' showZeros="0"'}${w2 & 0x40 ? ' rightToLeft="1"' : ""}${index === book.activeTab ? ' tabSelected="1"' : ""} workbookViewId="0">${pane}</sheetView></sheetViews>`;
      const defW = s.stdWidth != null ? s.stdWidth / 256 : s.defColWidth != null ? s.defColWidth + 0.71 : undefined;
      let out = HEAD + `<worksheet xmlns="${SS}" xmlns:r="${R}">`;
      const cn = s.codeName ? ` codeName="${esc(s.codeName)}"` : "";
      if (sheetPr.length || cn) out += `<sheetPr${cn}>${sheetPr.join("")}</sheetPr>`;
      out += `<dimension ref="${maxR || maxC ? "A1:" + a1(maxR, maxC) : "A1"}"/>` + view;
      out += `<sheetFormatPr${defW != null ? ` defaultColWidth="${num(Math.round(defW * 100) / 100)}"` : ""} defaultRowHeight="${twipsPt(defH)}"/>`;
      if (cols) out += `<cols>${cols}</cols>`;
      out += `<sheetData>${rowsXml}</sheetData>`;
      if (s.protect) {
        const allow = s.allow;
        const locks: Record<string, boolean> = { objects: !!s.objProtect, scenarios: !!s.scenProtect };
        if (allow != null) {
          const names = ["objects", "scenarios", "formatCells", "formatColumns", "formatRows", "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows", "selectLockedCells", "sort", "autoFilter", "pivotTables", "selectUnlockedCells"];
          names.forEach((n, i) => { if (i >= 2) locks[n] = !(allow & (1 << i)); });
        }
        const attrs = Object.entries(locks).map(([k, v]) => ` ${k}="${v ? 1 : 0}"`).join("");
        out += `<sheetProtection sheet="1"${attrs}${s.password ? ` password="${s.password.toString(16).toUpperCase().padStart(4, "0")}"` : ""}/>`;
      }
      if (s.autoFilter) out += `<autoFilter ref="${rangeA1(s.autoFilter)}"/>`;
      if (s.merges.length) out += `<mergeCells count="${s.merges.length}">${s.merges.map((m) => `<mergeCell ref="${rangeA1(m)}"/>`).join("")}</mergeCells>`;
      let prio = 1;
      for (const cf of s.condFmts) {
        const rules = cf.rules.map((rule) => {
          let dxf = "<dxf>";
          if (rule.bold != null || rule.italic || rule.strike || rule.underline || rule.color != null)
            dxf += `<font>${rule.bold ? "<b/>" : ""}${rule.italic ? "<i/>" : ""}${rule.strike ? "<strike/>" : ""}${rule.underline ? "<u/>" : ""}${rule.color != null ? colorXml("color", book, rule.color) : ""}</font>`;
          if (rule.fill != null) dxf += `<fill><patternFill patternType="solid">${colorXml("fgColor", book, rule.fill)}${colorXml("bgColor", book, rule.fill)}</patternFill></fill>`;
          dxf += "</dxf>";
          dxfs.push(dxf);
          const id = dxfs.length - 1;
          if (rule.type === 2) return rule.formula1 ? `<cfRule type="expression" dxfId="${id}" priority="${prio++}"><formula>${esc(rule.formula1)}</formula></cfRule>` : "";
          const op = CF_OPS[rule.op];
          if (!op || !rule.formula1) return "";
          return `<cfRule type="cellIs" dxfId="${id}" priority="${prio++}" operator="${op}"><formula>${esc(rule.formula1)}</formula>${rule.formula2 && (rule.op === 1 || rule.op === 2) ? `<formula>${esc(rule.formula2)}</formula>` : ""}</cfRule>`;
        }).join("");
        if (rules) out += `<conditionalFormatting sqref="${sqref(cf.ranges)}">${rules}</conditionalFormatting>`;
      }
      const dvs = s.dvs.filter((d) => d.type > 0 && d.type < DV_TYPES.length).map((d) => {
        const attrs = [`type="${DV_TYPES[d.type]}"`];
        if (d.errStyle) attrs.push(`errorStyle="${DV_ERR[d.errStyle] ?? "stop"}"`);
        if (d.type !== 3 && d.type !== 7 && d.op) attrs.push(`operator="${DV_OPS[d.op] ?? "between"}"`);
        if (d.allowBlank) attrs.push('allowBlank="1"');
        attrs.push('showInputMessage="1"', 'showErrorMessage="1"');
        if (d.errorTitle) attrs.push(`errorTitle="${esc(d.errorTitle)}"`);
        if (d.error) attrs.push(`error="${esc(d.error)}"`);
        if (d.promptTitle) attrs.push(`promptTitle="${esc(d.promptTitle)}"`);
        if (d.prompt) attrs.push(`prompt="${esc(d.prompt)}"`);
        attrs.push(`sqref="${sqref(d.ranges)}"`);
        return `<dataValidation ${attrs.join(" ")}>${d.formula1 != null ? `<formula1>${esc(d.formula1)}</formula1>` : ""}${d.formula2 != null ? `<formula2>${esc(d.formula2)}</formula2>` : ""}</dataValidation>`;
      });
      if (dvs.length) out += `<dataValidations count="${dvs.length}">${dvs.join("")}</dataValidations>`;
      if (s.links.length) {
        out += "<hyperlinks>" + s.links.map((l) => {
          const rid = l.url ? relId("hyperlink", l.url, true) : "";
          return `<hyperlink ref="${rangeA1(l.range)}"${rid ? ` r:id="${rid}"` : ""}${l.location ? ` location="${esc(l.location)}"` : ""}${l.tip ? ` tooltip="${esc(l.tip)}"` : ""}/>`;
        }).join("") + "</hyperlinks>";
      }
      const p = s.print;
      if (p.gridLines || p.headings || p.hCenter || p.vCenter)
        out += `<printOptions${p.hCenter ? ' horizontalCentered="1"' : ""}${p.vCenter ? ' verticalCentered="1"' : ""}${p.headings ? ' headings="1"' : ""}${p.gridLines ? ' gridLines="1"' : ""}/>`;
      if (p.left != null || p.right != null || p.top != null || p.bottom != null || p.headerMargin != null)
        out += `<pageMargins left="${num(p.left ?? 0.75)}" right="${num(p.right ?? 0.75)}" top="${num(p.top ?? 1)}" bottom="${num(p.bottom ?? 1)}" header="${num(p.headerMargin ?? 0.5)}" footer="${num(p.footerMargin ?? 0.5)}"/>`;
      if (p.paperSize != null || p.fitWidth != null) {
        const ps: string[] = [];
        if (!p.noPls) {
          if (p.paperSize) ps.push(`paperSize="${p.paperSize}"`);
          if (p.scale && p.scale !== 100) ps.push(`scale="${p.scale}"`);
          if (!p.noOrient) ps.push(`orientation="${p.portrait ? "portrait" : "landscape"}"`);
        }
        if (p.usePage && p.pageStart != null) ps.push(`firstPageNumber="${p.pageStart}"`, 'useFirstPageNumber="1"');
        if (p.fitWidth != null && p.fitWidth !== 1) ps.push(`fitToWidth="${p.fitWidth}"`);
        if (p.fitHeight != null && p.fitHeight !== 1) ps.push(`fitToHeight="${p.fitHeight}"`);
        if (p.overThenDown) ps.push('pageOrder="overThenDown"');
        if (ps.length) out += `<pageSetup ${ps.join(" ")}/>`;
      }
      if (p.header || p.footer) out += `<headerFooter>${p.header ? `<oddHeader>${esc(p.header)}</oddHeader>` : ""}${p.footer ? `<oddFooter>${esc(p.footer)}</oddFooter>` : ""}</headerFooter>`;
      if (p.rowBreaks?.length) out += `<rowBreaks count="${p.rowBreaks.length}" manualBreakCount="${p.rowBreaks.length}">${p.rowBreaks.map((b) => `<brk id="${b}" max="16383" man="1"/>`).join("")}</rowBreaks>`;
      if (p.colBreaks?.length) out += `<colBreaks count="${p.colBreaks.length}" manualBreakCount="${p.colBreaks.length}">${p.colBreaks.map((b) => `<brk id="${b}" max="1048575" man="1"/>`).join("")}</colBreaks>`;
      if (s.images.length) {
        drawingN++;
        const rid = relId("drawing", `../drawings/drawing${drawingN}.xml`);
        out += `<drawing r:id="${rid}"/>`;
        writeDrawing(s, drawingN);
      }
      if (s.notes.length) {
        commentsN++;
        relId("comments", `../comments${commentsN}.xml`);
        writeComments(s, commentsN);
      }
      return out + "</worksheet>";
    }

    function cellXml(cell: XlsCell): string {
      const xf = cell.xf < book.xfs.length ? cell.xf : 15;
      let attrs = ` r="${a1(cell.r, cell.c)}"${xf ? ` s="${xf}"` : ""}`;
      const f = cell.formula != null ? (cell.arrayRef ? `<f t="array" ref="${cell.arrayRef}">${esc(cell.formula)}</f>` : `<f>${esc(cell.formula)}</f>`) : "";
      const v = cell.val;
      switch (v.k) {
        case "n": {
          const n = book.date1904 && isDate(xf) ? v.v + 1462 : v.v;
          return `<c${attrs}>${f}<v>${num(n)}</v></c>`;
        }
        case "s":
          if (cell.formula != null) return `<c${attrs} t="str">${f}<v>${esc(v.v)}</v></c>`;
          return `<c${attrs} t="s"><v>${addShared(stringXml(v.v, v.runs))}</v></c>`;
        case "b": attrs += ' t="b"'; return `<c${attrs}>${f}<v>${v.v ? 1 : 0}</v></c>`;
        case "e": attrs += ' t="e"'; return `<c${attrs}>${f}<v>${esc(v.v)}</v></c>`;
        default: return f ? `<c${attrs}>${f}</c>` : `<c${attrs}/>`;
      }
    }
  });

  function stringXml(text: string, runs?: { ich: number; font: number }[]): string {
    const t = (s: string): string => `<t${/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ""}>${esc(s)}</t>`;
    if (!runs?.length) return `<si>${t(text)}</si>`;
    let out = "<si>";
    const sorted = runs.filter((r) => r.ich < text.length).sort((a, b) => a.ich - b.ich);
    if (!sorted.length || sorted[0]!.ich > 0) out += `<r>${t(text.slice(0, sorted[0]?.ich ?? text.length))}</r>`;
    sorted.forEach((r, i) => {
      const end = sorted[i + 1]?.ich ?? text.length;
      if (end > r.ich) out += `<r>${runPrXml(fontAt(r.font), book)}${t(text.slice(r.ich, end))}</r>`;
    });
    return out + "</si>";
  }

  function writeComments(s: XlsSheet, n: number): void {
    const authors: string[] = [];
    const list = s.notes.filter((x) => x.text || x.author).map((x) => {
      let ai = authors.indexOf(x.author);
      if (ai < 0) { ai = authors.length; authors.push(x.author); }
      return `<comment ref="${a1(x.r, x.c)}" authorId="${ai}"><text><t xml:space="preserve">${esc(x.text)}</t></text></comment>`;
    });
    files[`xl/comments${n}.xml`] = strToU8(HEAD + `<comments xmlns="${SS}"><authors>${authors.map((a) => `<author>${esc(a)}</author>`).join("")}</authors><commentList>${list.join("")}</commentList></comments>`);
    ct.push(`<Override PartName="/xl/comments${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`);
  }

  function writeDrawing(s: XlsSheet, n: number): void {
    const { colPx, rowPx } = xlsSheetPx(s);
    const EMU = 9525;
    const rels: string[] = [];
    const anchors = s.images.map((im, i) => {
      mediaN++;
      const ext = im.blip.ext === "jpeg" ? "jpeg" : im.blip.ext;
      const media = `image${mediaN}.${ext}`;
      files[`xl/media/${media}`] = im.blip.bytes;
      rels.push(`<Relationship Id="rId${i + 1}" Type="${R}/image" Target="../media/${media}"/>`);
      const a = im.anchor;
      const off = (frac: number, size: number): number => Math.round(Math.min(1, frac) * size * EMU);
      return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${a.col1}</xdr:col><xdr:colOff>${off(a.dx1 / 1024, colPx(a.col1))}</xdr:colOff><xdr:row>${a.row1}</xdr:row><xdr:rowOff>${off(a.dy1 / 256, rowPx(a.row1))}</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>${a.col2}</xdr:col><xdr:colOff>${off(a.dx2 / 1024, colPx(a.col2))}</xdr:colOff><xdr:row>${a.row2}</xdr:row><xdr:rowOff>${off(a.dy2 / 256, rowPx(a.row2))}</xdr:rowOff></xdr:to>` +
        `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 2}" name="${esc(im.name || `Picture ${i + 1}`)}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId${i + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        `<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
    });
    files[`xl/drawings/drawing${n}.xml`] = strToU8(HEAD + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}">${anchors.join("")}</xdr:wsDr>`);
    files[`xl/drawings/_rels/drawing${n}.xml.rels`] = strToU8(HEAD + `<Relationships xmlns="${PKG}">${rels.join("")}</Relationships>`);
    ct.push(`<Override PartName="/xl/drawings/drawing${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
  }

  // Workbook part, defined names and the package plumbing.
  const loadedIndex = new Map<number, number>();
  book.sheets.forEach((s, i) => loadedIndex.set(s.boundIndex ?? i, i));
  const names: string[] = [];
  for (const n of book.names) {
    if (!n.formula || n.func || n.name.startsWith("_xlfn.")) continue;
    const local = n.itab ? loadedIndex.get(n.itab - 1) : undefined;
    if (n.itab && local == null) continue;
    // Print ranges and the filter range are xlsx built-ins; the other built-ins keep their names.
    const xlnm = n.builtin === 6 || n.builtin === 7 || n.builtin === 13;
    const name = xlnm ? "_xlnm." + n.name : n.name;
    let formula = n.formula;
    if (n.builtin === 6) formula = formula.replace(/(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?=,|$)/g, (_m, a, c1, b, c2) => `${a}${c1}$1:${b}${c2}$65536`);
    if (!/^[A-Za-z_\\¡-￿][^\s]*$/.test(name)) continue;
    names.push(`<definedName name="${esc(name)}"${local != null ? ` localSheetId="${local}"` : ""}${n.hidden ? ' hidden="1"' : ""}>${esc(formula)}</definedName>`);
  }
  const sheetsXml = book.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}"${s.state === 1 ? ' state="hidden"' : s.state === 2 ? ' state="veryHidden"' : ""} r:id="rId${i + 1}"/>`).join("");
  const wbProt = book.protect || book.windowProtect ? `<workbookProtection${book.protect ? ' lockStructure="1"' : ""}${book.windowProtect ? ' lockWindows="1"' : ""}${book.password ? ` workbookPassword="${book.password.toString(16).toUpperCase().padStart(4, "0")}"` : ""}/>` : "";
  files["xl/workbook.xml"] = strToU8(HEAD + `<workbook xmlns="${SS}" xmlns:r="${R}"><workbookPr/>${wbProt}<bookViews><workbookView activeTab="${book.activeTab}"/></bookViews><sheets>${sheetsXml}</sheets>${names.length ? `<definedNames>${names.join("")}</definedNames>` : ""}</workbook>`);
  const n = book.sheets.length;
  const wbRels = book.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`);
  wbRels.push(`<Relationship Id="rId${n + 1}" Type="${R}/styles" Target="styles.xml"/>`, `<Relationship Id="rId${n + 2}" Type="${R}/sharedStrings" Target="sharedStrings.xml"/>`);
  files["xl/_rels/workbook.xml.rels"] = strToU8(HEAD + `<Relationships xmlns="${PKG}">${wbRels.join("")}</Relationships>`);
  files["xl/styles.xml"] = strToU8(stylesXml(book, dxfs).xml);
  files["xl/sharedStrings.xml"] = strToU8(HEAD + `<sst xmlns="${SS}" count="${shared.length}" uniqueCount="${shared.length}">${shared.join("")}</sst>`);
  files["_rels/.rels"] = strToU8(HEAD + `<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  files["[Content_Types].xml"] = strToU8(HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="bmp" ContentType="image/bmp"/><Default Extension="tiff" ContentType="image/tiff"/><Default Extension="pict" ContentType="image/pict"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    book.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    ct.join("") + "</Types>");
  return files;
}
