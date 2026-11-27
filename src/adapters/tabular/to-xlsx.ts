// A TBook as the xlsx parts sheetedit's xlsx reader understands, so HTML and Excel 2003 XML
// exports open with every xlsx editing path.
import { strToU8 } from "fflate";
import { colToLetters } from "../../core/model";
import type { TBook, TBorder, TCell, TFont, TRange, TRun, TSheet, TStyle } from "./model";

const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const esc = (s: string): string =>
  s.replace(/[&<>"]/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : "&quot;"))
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, (c) => `_x${c.charCodeAt(0).toString(16).padStart(4, "0").toUpperCase()}_`);

const num = (v: number): string => (Number.isFinite(v) ? String(v) : "0");
const a1 = (r: number, c: number): string => colToLetters(c + 1) + String(r + 1);
export const rangeA1 = (g: TRange): string => (g.r1 === g.r2 && g.c1 === g.c2 ? a1(g.r1, g.c1) : a1(g.r1, g.c1) + ":" + a1(g.r2, g.c2));
const sqref = (rs: TRange[]): string => rs.map(rangeA1).join(" ");
const argb = (css: string): string => "FF" + css.slice(1).toUpperCase();

/** Pixels to xlsx character units for a normal font whose digits are `mdw` px wide. */
const pxToChars = (px: number, mdw: number): number => Math.trunc((px / mdw) * 256) / 256;

const BUILTIN_FMT: Record<string, number> = {
  "General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "0.00E+00": 11, "# ?/?": 12, "# ??/??": 13,
  "m/d/yyyy": 14, "d-mmm-yy": 15, "d-mmm": 16, "mmm-yy": 17, "h:mm AM/PM": 18, "h:mm:ss AM/PM": 19, "h:mm": 20, "h:mm:ss": 21,
  "m/d/yyyy h:mm": 22, "mm:ss": 45, "[h]:mm:ss": 46, "mmss.0": 47, "##0.0E+0": 48, "@": 49,
};

function fontXml(f: TFont, tag: "font" | "rPr", dflt: { name: string; size: number }): string {
  let s = `<${tag}>`;
  if (f.bold) s += "<b/>";
  if (f.italic) s += "<i/>";
  if (f.strike) s += "<strike/>";
  if (f.underline) s += f.underline === "single" ? "<u/>" : `<u val="${f.underline}"/>`;
  if (f.vertAlign) s += `<vertAlign val="${f.vertAlign}"/>`;
  s += `<sz val="${num(f.size ?? dflt.size)}"/>`;
  if (f.color) s += `<color rgb="${argb(f.color)}"/>`;
  s += `<${tag === "font" ? "name" : "rFont"} val="${esc(f.name ?? dflt.name)}"/>`;
  return s + `</${tag}>`;
}

function borderSide(tag: string, b: TBorder | undefined): string {
  if (!b || !b.style || b.style === "none") return `<${tag}/>`;
  return `<${tag} style="${b.style}"><color ${b.color ? `rgb="${argb(b.color)}"` : 'auto="1"'}/></${tag}>`;
}

class StylePool {
  fonts: string[] = [];
  fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  borders = ["<border><left/><right/><top/><bottom/><diagonal/></border>"];
  xfs: string[] = [];
  numFmts = new Map<string, number>();
  dxfs: string[] = [];
  private xfIdx = new Map<string, number>();
  private styleIdx = new WeakMap<TStyle, number>();
  constructor(readonly dflt: { name: string; size: number }) {
    this.fonts.push(fontXml({}, "font", dflt));
    this.xfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    this.xfIdx.set(this.xfs[0]!, 0);
  }
  private intern(pool: string[], xml: string): number {
    const i = pool.indexOf(xml);
    if (i >= 0) return i;
    pool.push(xml);
    return pool.length - 1;
  }
  fmtId(code: string | undefined): number {
    if (!code || code === "General") return 0;
    const b = BUILTIN_FMT[code];
    if (b != null) return b;
    let id = this.numFmts.get(code);
    if (id == null) { id = 164 + this.numFmts.size; this.numFmts.set(code, id); }
    return id;
  }
  fillXml(fill: NonNullable<TStyle["fill"]>): string {
    const fg = fill.fg ? `<fgColor rgb="${argb(fill.fg)}"/>` : "";
    const bg = fill.bg ? `<bgColor rgb="${argb(fill.bg)}"/>` : fill.pattern === "solid" ? '<bgColor indexed="64"/>' : "";
    return `<fill><patternFill patternType="${fill.pattern}">${fg}${bg}</patternFill></fill>`;
  }
  borderXml(b: NonNullable<TStyle["border"]>): string {
    return `<border${b.diagUp ? ' diagonalUp="1"' : ""}${b.diagDown ? ' diagonalDown="1"' : ""}>${borderSide("left", b.left)}${borderSide("right", b.right)}${borderSide("top", b.top)}${borderSide("bottom", b.bottom)}${borderSide("diagonal", b.diagonal)}</border>`;
  }
  /** The cellXfs index of a style (0 for none). */
  xf(st: TStyle | undefined): number {
    if (!st) return 0;
    const cached = this.styleIdx.get(st);
    if (cached != null) return cached;
    const font = st.font ? this.intern(this.fonts, fontXml(st.font, "font", this.dflt)) : 0;
    const fill = st.fill && st.fill.pattern !== "none" ? this.intern(this.fills, this.fillXml(st.fill)) : 0;
    const border = st.border ? this.intern(this.borders, this.borderXml(st.border)) : 0;
    const fmt = this.fmtId(st.numFmt);
    const al: string[] = [];
    const a = st.align;
    if (a?.h) al.push(`horizontal="${a.h}"`);
    if (a?.v) al.push(`vertical="${a.v}"`);
    if (a?.wrap) al.push('wrapText="1"');
    if (a?.rotate) al.push(`textRotation="${a.rotate}"`);
    if (a?.indent) al.push(`indent="${a.indent}"`);
    if (a?.shrink) al.push('shrinkToFit="1"');
    const prot = st.locked === false || st.hidden ? `<protection${st.locked === false ? ' locked="0"' : ""}${st.hidden ? ' hidden="1"' : ""}/>` : "";
    const xml = `<xf numFmtId="${fmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${fmt ? ' applyNumberFormat="1"' : ""}${font ? ' applyFont="1"' : ""}${fill ? ' applyFill="1"' : ""}${border ? ' applyBorder="1"' : ""}${al.length ? ' applyAlignment="1"' : ""}${prot ? ' applyProtection="1"' : ""}>${al.length ? `<alignment ${al.join(" ")}/>` : ""}${prot}</xf>`;
    let i = this.xfIdx.get(xml);
    if (i == null) { i = this.xfs.length; this.xfs.push(xml); this.xfIdx.set(xml, i); }
    this.styleIdx.set(st, i);
    return i;
  }
  dxf(st: TStyle): number {
    let s = "<dxf>";
    const f = st.font;
    if (f) s += `<font>${f.bold ? "<b/>" : ""}${f.italic ? "<i/>" : ""}${f.strike ? "<strike/>" : ""}${f.underline ? "<u/>" : ""}${f.color ? `<color rgb="${argb(f.color)}"/>` : ""}</font>`;
    if (st.numFmt) s += `<numFmt numFmtId="${this.fmtId(st.numFmt)}" formatCode="${esc(st.numFmt)}"/>`;
    if (st.fill?.fg) s += `<fill><patternFill patternType="solid"><fgColor rgb="${argb(st.fill.fg)}"/><bgColor rgb="${argb(st.fill.fg)}"/></patternFill></fill>`;
    if (st.border) s += this.borderXml(st.border);
    s += "</dxf>";
    this.dxfs.push(s);
    return this.dxfs.length - 1;
  }
  xml(): string {
    const fmts = [...this.numFmts].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`);
    return HEAD + `<styleSheet xmlns="${SS}">` +
      (fmts.length ? `<numFmts count="${fmts.length}">${fmts.join("")}</numFmts>` : "") +
      `<fonts count="${this.fonts.length}">${this.fonts.join("")}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join("")}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join("")}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join("")}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      `<dxfs count="${this.dxfs.length}">${this.dxfs.join("")}</dxfs>` +
      "</styleSheet>";
  }
}

const tEl = (s: string): string => `<t${/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ""}>${esc(s)}</t>`;

/** Convert the book into xlsx parts. */
export function tBookToXlsx(book: TBook): Record<string, Uint8Array> {
  const dflt = book.defaultFont ?? { name: "Calibri", size: 11 };
  const mdw = Math.max(1, Math.round((dflt.size * 4) / 3 / 2));
  const styles = new StylePool(dflt);
  const files: Record<string, Uint8Array> = {};
  const ct: string[] = [];
  const shared: string[] = [];
  const sharedIdx = new Map<string, number>();
  const addShared = (xml: string): number => {
    let i = sharedIdx.get(xml);
    if (i == null) { i = shared.length; shared.push(xml); sharedIdx.set(xml, i); }
    return i;
  };
  const stringXml = (text: string, runs?: TRun[]): string => {
    if (!runs || runs.length < 2) return `<si>${tEl(text)}</si>`;
    return "<si>" + runs.map((r) => `<r>${r.font ? fontXml(r.font, "rPr", dflt) : ""}${tEl(r.text)}</r>`).join("") + "</si>";
  };
  let commentsN = 0;
  book.sheets.forEach((sh, si) => {
    const rels: string[] = [];
    const relId = (type: string, target: string, external = false): string => {
      const id = `rId${rels.length + 1}`;
      rels.push(`<Relationship Id="${id}" Type="${R}/${type}" Target="${esc(target)}"${external ? ' TargetMode="External"' : ""}/>`);
      return id;
    };
    files[`xl/worksheets/sheet${si + 1}.xml`] = strToU8(sheetXml(sh, si));
    if (rels.length) files[`xl/worksheets/_rels/sheet${si + 1}.xml.rels`] = strToU8(HEAD + `<Relationships xmlns="${PKG}">${rels.join("")}</Relationships>`);

    function cellXml(cell: TCell): string {
      const xf = styles.xf(cell.style);
      const attrs = ` r="${a1(cell.r, cell.c)}"${xf ? ` s="${xf}"` : ""}`;
      const f = cell.f != null ? (cell.array ? `<f t="array" ref="${cell.array}">${esc(cell.f)}</f>` : `<f>${esc(cell.f)}</f>`) : "";
      const v = cell.v;
      if (!v) return f ? `<c${attrs}>${f}</c>` : `<c${attrs}/>`;
      switch (v.t) {
        case "n": return `<c${attrs}>${f}<v>${num(v.v)}</v></c>`;
        case "s":
          if (f) return `<c${attrs} t="str">${f}<v>${esc(v.v)}</v></c>`;
          return `<c${attrs} t="s"><v>${addShared(stringXml(v.v, v.runs))}</v></c>`;
        case "b": return `<c${attrs} t="b">${f}<v>${v.v ? 1 : 0}</v></c>`;
        case "e": return `<c${attrs} t="e">${f}<v>${esc(v.v)}</v></c>`;
      }
    }

    function sheetXml(s: TSheet, index: number): string {
      let maxR = 0, maxC = 0;
      const byRow = new Map<number, TCell[]>();
      for (const cell of s.cells.values()) {
        if (cell.r > 1048575 || cell.c > 16383) continue;
        let list = byRow.get(cell.r);
        if (!list) byRow.set(cell.r, (list = []));
        list.push(cell);
        if (cell.r > maxR) maxR = cell.r;
        if (cell.c > maxC) maxC = cell.c;
      }
      const rowNums = new Set<number>([...byRow.keys(), ...s.rows.keys()]);
      let rowsXml = "";
      for (const r of [...rowNums].sort((a, b) => a - b)) {
        if (r > 1048575) continue;
        const info = s.rows.get(r);
        let attrs = ` r="${r + 1}"`;
        if (info?.height != null) attrs += ` ht="${num(info.height)}" customHeight="1"`;
        if (info?.hidden) attrs += ' hidden="1"';
        if (info?.style) attrs += ` s="${styles.xf(info.style)}" customFormat="1"`;
        const cells = (byRow.get(r) ?? []).sort((a, b) => a.c - b.c).map(cellXml).join("");
        rowsXml += cells ? `<row${attrs}>${cells}</row>` : `<row${attrs}/>`;
      }
      const cols = [...s.cols].filter(([c]) => c <= 16383).sort((a, b) => a[0] - b[0]).map(([c, info]) => {
        const w = info.width != null ? pxToChars(info.width, mdw) : s.defaultColWidth != null ? pxToChars(s.defaultColWidth, mdw) : 8.43;
        return `<col min="${c + 1}" max="${c + 1}" width="${num(w)}"${info.width != null ? ' customWidth="1"' : ""}${info.hidden ? ' hidden="1"' : ""}${info.style ? ` style="${styles.xf(info.style)}"` : ""}/>`;
      }).join("");
      const sheetPr: string[] = [];
      if (s.tabColor) sheetPr.push(`<tabColor rgb="${argb(s.tabColor)}"/>`);
      if (s.print?.fitToPage) sheetPr.push('<pageSetUpPr fitToPage="1"/>');
      let pane = "";
      const fz = s.freeze;
      if (fz && (fz.rows || fz.cols)) {
        const tl = a1(fz.top ?? fz.rows, fz.left ?? fz.cols);
        const active = fz.rows && fz.cols ? "bottomRight" : fz.rows ? "bottomLeft" : "topRight";
        pane = `<pane${fz.cols ? ` xSplit="${fz.cols}"` : ""}${fz.rows ? ` ySplit="${fz.rows}"` : ""} topLeftCell="${tl}" activePane="${active}" state="frozen"/>`;
      }
      const view = `<sheetViews><sheetView${s.hideGrid ? ' showGridLines="0"' : ""}${s.hideZeros ? ' showZeros="0"' : ""}${s.rtl ? ' rightToLeft="1"' : ""}${index === (book.activeSheet ?? 0) ? ' tabSelected="1"' : ""} workbookViewId="0">${pane}</sheetView></sheetViews>`;
      let out = HEAD + `<worksheet xmlns="${SS}" xmlns:r="${R}">`;
      if (sheetPr.length) out += `<sheetPr>${sheetPr.join("")}</sheetPr>`;
      out += `<dimension ref="${maxR || maxC ? "A1:" + a1(maxR, maxC) : "A1"}"/>` + view;
      const defW = s.defaultColWidth != null ? ` defaultColWidth="${num(pxToChars(s.defaultColWidth, mdw))}"` : "";
      out += `<sheetFormatPr${defW} defaultRowHeight="${num(s.defaultRowHeight ?? 15)}"${s.defaultRowHeight != null ? ' customHeight="1"' : ""}/>`;
      if (cols) out += `<cols>${cols}</cols>`;
      out += `<sheetData>${rowsXml}</sheetData>`;
      if (s.autoFilter) out += `<autoFilter ref="${rangeA1(s.autoFilter)}"/>`;
      if (s.merges.length) out += `<mergeCells count="${s.merges.length}">${s.merges.map((m) => `<mergeCell ref="${rangeA1(m)}"/>`).join("")}</mergeCells>`;
      let prio = 1;
      for (const cf of s.condFormats) {
        const rules = cf.rules.map((rule) => {
          const id = styles.dxf(rule.style);
          if (rule.kind === "expression") return `<cfRule type="expression" dxfId="${id}" priority="${prio++}"><formula>${esc(rule.f1)}</formula></cfRule>`;
          const two = rule.op === "between" || rule.op === "notBetween";
          return `<cfRule type="cellIs" dxfId="${id}" priority="${prio++}" operator="${rule.op}"><formula>${esc(rule.f1)}</formula>${two && rule.f2 != null ? `<formula>${esc(rule.f2)}</formula>` : ""}</cfRule>`;
        }).join("");
        if (rules && cf.ranges.length) out += `<conditionalFormatting sqref="${sqref(cf.ranges)}">${rules}</conditionalFormatting>`;
      }
      const dvs = s.validations.filter((d) => d.ranges.length).map((d) => {
        const attrs = [`type="${d.type}"`];
        if (d.errorStyle && d.errorStyle !== "stop") attrs.push(`errorStyle="${d.errorStyle}"`);
        if (d.op && d.type !== "list" && d.type !== "custom") attrs.push(`operator="${d.op}"`);
        if (d.allowBlank) attrs.push('allowBlank="1"');
        if (d.noDropdown) attrs.push('showDropDown="1"');
        attrs.push('showInputMessage="1"', 'showErrorMessage="1"');
        if (d.errorTitle) attrs.push(`errorTitle="${esc(d.errorTitle)}"`);
        if (d.error) attrs.push(`error="${esc(d.error)}"`);
        if (d.promptTitle) attrs.push(`promptTitle="${esc(d.promptTitle)}"`);
        if (d.prompt) attrs.push(`prompt="${esc(d.prompt)}"`);
        attrs.push(`sqref="${sqref(d.ranges)}"`);
        return `<dataValidation ${attrs.join(" ")}>${d.f1 != null ? `<formula1>${esc(d.f1)}</formula1>` : ""}${d.f2 != null ? `<formula2>${esc(d.f2)}</formula2>` : ""}</dataValidation>`;
      });
      if (dvs.length) out += `<dataValidations count="${dvs.length}">${dvs.join("")}</dataValidations>`;
      const links = [...s.cells.values()].filter((c) => c.link && (c.link.href || c.link.location));
      if (links.length) {
        out += "<hyperlinks>" + links.map((c) => {
          const l = c.link!;
          const rid = l.href ? relId("hyperlink", l.href, true) : "";
          return `<hyperlink ref="${a1(c.r, c.c)}"${rid ? ` r:id="${rid}"` : ""}${l.location ? ` location="${esc(l.location)}"` : ""}${l.tip ? ` tooltip="${esc(l.tip)}"` : ""}/>`;
        }).join("") + "</hyperlinks>";
      }
      const p = s.print;
      if (p) {
        if (p.gridLines || p.hCenter || p.vCenter)
          out += `<printOptions${p.hCenter ? ' horizontalCentered="1"' : ""}${p.vCenter ? ' verticalCentered="1"' : ""}${p.gridLines ? ' gridLines="1"' : ""}/>`;
        const m = p.margins;
        if (m) out += `<pageMargins left="${num(m.left ?? 0.7)}" right="${num(m.right ?? 0.7)}" top="${num(m.top ?? 0.75)}" bottom="${num(m.bottom ?? 0.75)}" header="${num(m.header ?? 0.3)}" footer="${num(m.footer ?? 0.3)}"/>`;
        const ps: string[] = [];
        if (p.paperSize) ps.push(`paperSize="${p.paperSize}"`);
        if (p.scale && p.scale !== 100) ps.push(`scale="${p.scale}"`);
        if (p.fitWidth != null && p.fitWidth !== 1) ps.push(`fitToWidth="${p.fitWidth}"`);
        if (p.fitHeight != null && p.fitHeight !== 1) ps.push(`fitToHeight="${p.fitHeight}"`);
        if (p.landscape) ps.push('orientation="landscape"');
        if (ps.length) out += `<pageSetup ${ps.join(" ")}/>`;
        if (p.header || p.footer) out += `<headerFooter>${p.header ? `<oddHeader>${esc(p.header)}</oddHeader>` : ""}${p.footer ? `<oddFooter>${esc(p.footer)}</oddFooter>` : ""}</headerFooter>`;
      }
      const notes = [...s.cells.values()].filter((c) => c.comment);
      if (notes.length) {
        commentsN++;
        relId("comments", `../comments${commentsN}.xml`);
        const authors: string[] = [];
        const list = notes.map((c) => {
          const au = c.comment!.author ?? "";
          let ai = authors.indexOf(au);
          if (ai < 0) { ai = authors.length; authors.push(au); }
          return `<comment ref="${a1(c.r, c.c)}" authorId="${ai}"><text><t xml:space="preserve">${esc(c.comment!.text)}</t></text></comment>`;
        });
        files[`xl/comments${commentsN}.xml`] = strToU8(HEAD + `<comments xmlns="${SS}"><authors>${authors.map((a) => `<author>${esc(a)}</author>`).join("")}</authors><commentList>${list.join("")}</commentList></comments>`);
        ct.push(`<Override PartName="/xl/comments${commentsN}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`);
      }
      return out + "</worksheet>";
    }
  });

  const names = book.names.filter((n) => n.ref && /^[A-Za-z_\\¡-￿][^\s]*$/.test(n.name)).map((n) =>
    `<definedName name="${esc(n.name)}"${n.sheet != null ? ` localSheetId="${n.sheet}"` : ""}${n.hidden ? ' hidden="1"' : ""}>${esc(n.ref)}</definedName>`);
  const sheetsXml = book.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ""} r:id="rId${i + 1}"/>`).join("");
  files["xl/workbook.xml"] = strToU8(HEAD + `<workbook xmlns="${SS}" xmlns:r="${R}"><workbookPr${book.date1904 ? ' date1904="1"' : ""}/><bookViews><workbookView activeTab="${book.activeSheet ?? 0}"/></bookViews><sheets>${sheetsXml}</sheets>${names.length ? `<definedNames>${names.join("")}</definedNames>` : ""}<calcPr calcId="191029"/></workbook>`);
  const n = book.sheets.length;
  const wbRels = book.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`);
  wbRels.push(`<Relationship Id="rId${n + 1}" Type="${R}/styles" Target="styles.xml"/>`, `<Relationship Id="rId${n + 2}" Type="${R}/sharedStrings" Target="sharedStrings.xml"/>`);
  files["xl/_rels/workbook.xml.rels"] = strToU8(HEAD + `<Relationships xmlns="${PKG}">${wbRels.join("")}</Relationships>`);
  files["xl/styles.xml"] = strToU8(styles.xml());
  files["xl/sharedStrings.xml"] = strToU8(HEAD + `<sst xmlns="${SS}" count="${shared.length}" uniqueCount="${shared.length}">${shared.join("")}</sst>`);
  files["_rels/.rels"] = strToU8(HEAD + `<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  files["[Content_Types].xml"] = strToU8(HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    book.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    ct.join("") + "</Types>");
  return files;
}
