// HTML tables as a workbook: Excel's "Save as Web Page" files (x:num / x:str / x:fmla, class styles,
// mso-number-format, <col> widths, its XML island of sheet names and panes), the single-file MHTML
// variant, and any page with tables, each table a sheet.
import { regional } from "../../core/regional";
import { DEFAULT_PALETTE } from "../xls/palette";
import { cssPt, declsToStyle, htmlFontSize, parseCss, parseDecls, type Decls, type Rule } from "../tabular/css";
import { cellKey, newSheet, uniqueSheetNames, type TBook, type TCell, type TFont, type TName, type TRun, type TSheet, type TStyle, type TValue } from "../tabular/model";
import { childEls, decodeEntities, findAll, parseHtmlTree, textOf, type HNode } from "../tabular/tree";
import { sniffCommaDecimal, typeCellText, parseLooseNumber } from "../tabular/values";

export interface HtmlReadOptions {
  /** Parse with DOMParser when available (default true); false forces the built-in tokenizer. */
  useDom?: boolean;
  /** The name of the n-th sheet (1-based) when the file names none. */
  sheetName?: (n: number) => string;
}

const INHERITED = ["color", "font-family", "font-size", "font-weight", "font-style", "text-align", "white-space", "text-decoration", "text-underline-style", "mso-spacerun"];
const BLOCK = new Set(["p", "div", "li", "h1", "h2", "h3", "h4", "h5", "h6", "pre", "blockquote", "table", "tr", "ul", "ol", "dl", "dt", "dd", "address", "section", "article", "header", "footer"]);
const SKIP = new Set(["script", "style", "noscript", "head", "title", "template", "select", "button", "iframe", "object", "svg", "math", "xml"]);
const MAX_LINES = 20;
const MAX_SPAN = 1000;

interface Ctx { id: number; decls: Decls; kids: Map<string, Ctx> }

// ---- Excel's XML island: sheet names, panes, names --------------------------------------------

interface IslandSheet { name: string; source?: string; freeze?: TSheet["freeze"]; state?: TSheet["state"]; tab?: string; hideGrid?: boolean; hideZeros?: boolean; selected?: boolean }
interface Island { sheets: IslandSheet[]; active?: number; names: { name: string; sheet?: number; formula: string; hidden: boolean }[] }

const tagText = (src: string, tag: string): string | undefined => {
  const m = new RegExp(`<x:${tag}>([^<]*)</x:${tag}>`, "i").exec(src);
  return m ? decodeEntities(m[1]!).trim() : undefined;
};
const hasTag = (src: string, tag: string): boolean => new RegExp(`<x:${tag}\\s*/?>`, "i").test(src);

function readIsland(raw: string): Island {
  const out: Island = { sheets: [], names: [] };
  for (const m of raw.matchAll(/<x:ExcelWorksheet>([\s\S]*?)<\/x:ExcelWorksheet>/gi)) {
    const b = m[1]!;
    const s: IslandSheet = { name: tagText(b, "Name") ?? "" };
    const src = /<x:WorksheetSource\s+HRef\s*=\s*["']?([^"'\s>]+)/i.exec(b);
    if (src) s.source = decodeEntities(src[1]!);
    if (hasTag(b, "FreezePanes") || hasTag(b, "FrozenNoSplit")) {
      const rows = Number(tagText(b, "SplitHorizontal") ?? 0), cols = Number(tagText(b, "SplitVertical") ?? 0);
      if (rows || cols) s.freeze = { rows, cols, top: Number(tagText(b, "TopRowBottomPane") ?? rows) || rows, left: Number(tagText(b, "LeftColumnRightPane") ?? cols) || cols };
    }
    const vis = tagText(b, "Visible");
    if (vis === "SheetHidden") s.state = "hidden";
    else if (vis === "SheetVeryHidden") s.state = "veryHidden";
    const tab = Number(tagText(b, "TabColorIndex") ?? NaN);
    if (tab >= 0 && tab < 64) s.tab = DEFAULT_PALETTE[tab < 8 ? tab : tab - 8];
    if (hasTag(b, "DoNotDisplayGridlines")) s.hideGrid = true;
    if (hasTag(b, "DoNotDisplayZeros")) s.hideZeros = true;
    if (hasTag(b, "Selected")) s.selected = true;
    out.sheets.push(s);
  }
  const active = tagText(raw, "ActiveSheet");
  if (active != null) out.active = Number(active);
  for (const m of raw.matchAll(/<x:ExcelName>([\s\S]*?)<\/x:ExcelName>/gi)) {
    const b = m[1]!;
    const name = tagText(b, "Name"), formula = tagText(b, "Formula");
    const idx = tagText(b, "SheetIndex");
    if (name && formula) out.names.push({ name, formula: formula.replace(/^=/, ""), sheet: idx != null ? Number(idx) - 1 : undefined, hidden: hasTag(b, "Hidden") });
  }
  return out;
}

// ---- MHTML ------------------------------------------------------------------------------------

interface MimePart { location: string; type: string; text: string }

function decodeQp(s: string): Uint8Array {
  const src = s.replace(/=\r?\n/g, "");
  const out: number[] = [];
  for (let i = 0; i < src.length; i++) {
    const ch = src.charCodeAt(i);
    if (ch === 0x3d && /^[0-9A-Fa-f]{2}$/.test(src.slice(i + 1, i + 3))) { out.push(parseInt(src.slice(i + 1, i + 3), 16)); i += 2; }
    else out.push(ch & 0xff);
  }
  return new Uint8Array(out);
}

function decodeBase64(s: string): Uint8Array {
  const bin = atob(s.replace(/[^A-Za-z0-9+/=]/g, ""));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** Split a multipart/related (MHTML) file into its decoded parts. */
export function parseMhtml(text: string): MimePart[] {
  const bm = /boundary\s*=\s*"?([^";\r\n]+)"?/i.exec(text);
  if (!bm) return [];
  const parts: MimePart[] = [];
  for (const chunk of text.split("--" + bm[1]!).slice(1)) {
    const sep = /\r?\n\r?\n/.exec(chunk);
    if (!sep) continue;
    const head = chunk.slice(0, sep.index), body = chunk.slice(sep.index + sep[0].length);
    const header = (name: string): string => new RegExp(`^${name}:\\s*([^\\r\\n]*)`, "im").exec(head)?.[1]?.trim() ?? "";
    const cte = header("Content-Transfer-Encoding").toLowerCase();
    const type = header("Content-Type");
    const charset = /charset\s*=\s*"?([\w.:-]+)/i.exec(type)?.[1] ?? "utf-8";
    const bytes = cte === "base64" ? decodeBase64(body) : cte === "quoted-printable" ? decodeQp(body) : Uint8Array.from(body, (c) => c.charCodeAt(0) & 0xff);
    let decoded: string;
    try { decoded = new TextDecoder(/^(us-ascii|iso-8859-1)$/i.test(charset) ? "windows-1252" : charset).decode(bytes); }
    catch { decoded = new TextDecoder("windows-1252").decode(bytes); }
    parts.push({ location: header("Content-Location"), type: type.split(";")[0]!.toLowerCase(), text: decoded });
  }
  return parts;
}

const baseName = (p: string): string => decodeURIComponent(p.replace(/\\/g, "/").split(/[?#]/)[0]!.split("/").pop() ?? "").toLowerCase();

// ---- the reader -------------------------------------------------------------------------------

interface DocCtx { rules: Rule[]; excel: boolean; excelHeights: boolean; ctxId: number; root: Ctx }

function newDoc(rules: Rule[], excel: boolean, excelHeights: boolean): DocCtx {
  return { rules, excel, excelHeights, ctxId: 1, root: { id: 0, decls: {}, kids: new Map() } };
}

function matches(rule: Rule, node: HNode): boolean {
  const ok = (sel: Rule["chain"][number]["sel"], n: HNode): boolean => {
    if (sel.tag && sel.tag !== n.tag) return false;
    if (sel.id && (n.attrs.id ?? "").toLowerCase() !== sel.id) return false;
    if (sel.classes.length) {
      const cls = (n.attrs.class ?? "").toLowerCase().split(/\s+/);
      if (!sel.classes.every((c) => cls.includes(c))) return false;
    }
    return true;
  };
  const chain = rule.chain;
  if (!ok(chain[chain.length - 1]!.sel, node)) return false;
  let n: HNode | undefined = node;
  for (let i = chain.length - 2; i >= 0; i--) {
    const direct = chain[i + 1]!.comb === ">";
    n = n?.parent;
    while (n && !ok(chain[i]!.sel, n)) { if (direct) return false; n = n.parent; }
    if (!n) return false;
  }
  return true;
}

/** Presentational attributes and tags as the declarations a browser's defaults give them. */
function hints(n: HNode): Decls {
  const d: Decls = {};
  const a = n.attrs;
  if (a.bgcolor) d["background-color"] = a.bgcolor;
  if (a.align && n.tag !== "table" && n.tag !== "img") d["text-align"] = a.align;
  if (a.valign) d["vertical-align"] = a.valign;
  switch (n.tag) {
    case "b": case "strong": d["font-weight"] = "bold"; break;
    case "i": case "em": case "cite": case "var": d["font-style"] = "italic"; break;
    case "u": case "ins": d["text-decoration"] = "underline"; break;
    case "s": case "strike": case "del": d["text-decoration"] = "line-through"; break;
    case "sup": d["vertical-align"] = "super"; break;
    case "sub": d["vertical-align"] = "sub"; break;
    case "center": d["text-align"] = "center"; break;
    case "th": d["font-weight"] = "bold"; if (!a.align) d["text-align"] = "center"; break;
    case "h1": case "h2": case "h3": case "h4": case "h5": case "h6": d["font-weight"] = "bold"; break;
    case "font":
      if (a.color) d.color = a.color;
      if (a.face) d["font-family"] = a.face;
      if (a.size) { const pt = htmlFontSize(a.size); if (pt) d["font-size"] = pt + "pt"; }
      break;
  }
  return d;
}

/** The effective declarations of a node, given its parent's context. */
function ctxOf(doc: DocCtx, parent: Ctx, n: HNode): Ctx {
  const sig = `${n.tag}|${n.attrs.class ?? ""}|${n.attrs.id ?? ""}|${n.attrs.style ?? ""}|${n.attrs.bgcolor ?? ""}|${n.attrs.align ?? ""}|${n.attrs.valign ?? ""}|${n.attrs.color ?? ""}|${n.attrs.face ?? ""}|${n.attrs.size ?? ""}`;
  const hit = parent.kids.get(sig);
  if (hit) return hit;
  const d: Decls = {};
  for (const k of INHERITED) if (parent.decls[k] != null) d[k] = parent.decls[k]!;
  // Table backgrounds, row alignment and table borders reach the cells as a browser paints them.
  if (n.tag === "td" || n.tag === "th" || n.tag === "tr" || n.tag === "tbody" || n.tag === "thead" || n.tag === "tfoot") {
    for (const k of ["background-color", "background", "vertical-align", "mso-pattern", "x-cell-border"]) if (parent.decls[k] != null && d[k] == null) d[k] = parent.decls[k]!;
  }
  Object.assign(d, hints(n));
  const own: Decls = {};
  const matched = doc.rules.filter((r) => matches(r, n)).sort((x, y) => x.spec - y.spec || x.order - y.order);
  for (const r of matched) Object.assign(own, r.decls);
  if (n.attrs.style) parseDecls(n.attrs.style, own);
  if (own.background && own["background-color"] == null) delete d["background-color"];
  Object.assign(d, own);
  if (n.tag === "table" && Number(n.attrs.border) > 0) d["x-cell-border"] = "0.75pt solid " + (n.attrs.bordercolor ?? "#000000");
  // Relative font sizes resolve against the parent's size here, so children inherit points.
  if (d["font-size"] && own["font-size"]) {
    const pt = cssPt(d["font-size"], cssPt(parent.decls["font-size"] ?? "") ?? 11);
    if (pt) d["font-size"] = pt + "pt";
  }
  const ctx: Ctx = { id: doc.ctxId++, decls: d, kids: new Map() };
  parent.kids.set(sig, ctx);
  return ctx;
}

const isHidden = (d: Decls): boolean => /^none$/i.test(d.display ?? "") || /^hidden$/i.test(d.visibility ?? "");

interface Collected { text: string; runs: { text: string; decls: Decls }[]; link?: { href?: string; location?: string } }

/** A cell's text with its styled runs: whitespace collapsed, <br> and blocks as line breaks. */
function collectCell(doc: DocCtx, cellCtx: Ctx, cell: HNode): Collected {
  const runs: { text: string; decls: Decls }[] = [];
  let link: Collected["link"];
  let lastSpace = true;
  const push = (s: string, decls: Decls): void => {
    if (!s) return;
    const prev = runs[runs.length - 1];
    if (prev && prev.decls === decls) prev.text += s;
    else runs.push({ text: s, decls });
  };
  const newline = (decls: Decls): void => {
    const prev = runs[runs.length - 1];
    if (!prev) return;
    prev.text = prev.text.replace(/ +$/, "");
    if (!prev.text.endsWith("\n")) push("\n", decls);
    lastSpace = true;
  };
  const walk = (n: HNode, ctx: Ctx, pre: boolean): void => {
    for (const k of n.kids) {
      if (typeof k === "string") {
        if (pre) { push(k, ctx.decls); lastSpace = /\s$/.test(k); continue; }
        let s = k.replace(/[ \t\n\r\f]+/g, " ");
        if (lastSpace) s = s.replace(/^ /, "");
        if (s) { push(s, ctx.decls); lastSpace = s.endsWith(" "); }
        continue;
      }
      if (SKIP.has(k.tag)) continue;
      if (k.tag === "br") { push("\n", ctx.decls); lastSpace = true; continue; }
      const kc = ctxOf(doc, ctx, k);
      // Excel keeps the clipped part of a cell's text in a hidden span; it is still the cell's text.
      if (isHidden(kc.decls) && !(doc.excel && k.tag === "span")) continue;
      if (k.tag === "a" && !link) {
        const href = (k.attrs.href ?? "").trim();
        if (href && !/^javascript:|^#_msocom/i.test(href)) link = href.startsWith("#") ? { location: decodeURIComponent(href.slice(1)) } : { href };
      }
      if (k.tag === "img") { if (k.attrs.alt) push(k.attrs.alt, kc.decls); continue; }
      const block = BLOCK.has(k.tag);
      if (block) newline(kc.decls);
      const spacerun = /yes/i.test(kc.decls["mso-spacerun"] ?? "");
      walk(k, kc, pre || spacerun || k.tag === "pre" || /^pre/i.test(kc.decls["white-space"] ?? ""));
      if (block) newline(kc.decls);
      if (k.tag === "td" || k.tag === "th") push("\t", kc.decls);
    }
  };
  walk(cell, cellCtx, false);
  // Trim the ends and the spaces around line breaks.
  let text = runs.map((r) => r.text).join("");
  const lead = text.length - text.replace(/^[\s ]+/, "").length;
  const trail = text.length - text.replace(/[\s ]+$/, "").length;
  if (lead || trail) {
    let skip = lead, keep = text.length - lead - trail;
    for (const r of runs) {
      const cut = Math.min(skip, r.text.length);
      r.text = r.text.slice(cut);
      skip -= cut;
      r.text = r.text.slice(0, Math.max(0, keep));
      keep -= r.text.length;
    }
  }
  for (const r of runs) r.text = r.text.replace(/ *\n */g, "\n");
  const kept = runs.filter((r) => r.text);
  text = kept.map((r) => r.text).join("");
  return { text, runs: kept, link };
}

const sameFont = (a: TFont | undefined, b: TFont | undefined): boolean => JSON.stringify(a ?? {}) === JSON.stringify(b ?? {});

/** Parse a number Excel wrote with the table's displayed separators. */
function excelNumber(text: string, dec: string, group: string): number | null {
  let t = text.trim().replace(/[\s ]/g, "");
  if (group) t = t.split(group).join("");
  if (dec !== ".") t = t.replace(dec, ".");
  t = t.replace(/[$€£¥%]/g, "");
  if (/^\(.*\)$/.test(t)) t = "-" + t.slice(1, -1);
  const n = Number(t);
  return t !== "" && Number.isFinite(n) ? n : null;
}

const ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"]);

interface TableResult { sheet: TSheet; cells: number }

function readTable(doc: DocCtx, table: HNode, tableCtx: Ctx, name: string): TableResult {
  const sheet = newSheet(name);
  const excel = doc.excel;
  const td = tableCtx.decls;
  const dispDec = td["mso-displayed-decimal-separator"] ? td["mso-displayed-decimal-separator"].replace(/^["']|["']$/g, "").replace(/\\/g, "") : ".";
  const dispGroup = td["mso-displayed-thousand-separator"] ? td["mso-displayed-thousand-separator"].replace(/^["']|["']$/g, "").replace(/\\/g, "") : ",";
  // Column widths from <col> / <colgroup>.
  let col = 0;
  const cols = (parent: HNode, pctx: Ctx): void => {
    for (const k of childEls(parent)) {
      if (k.tag === "colgroup") {
        const gctx = ctxOf(doc, pctx, k);
        if (childEls(k, "col").length) cols(k, gctx);
        else col += Math.min(MAX_SPAN, Math.max(1, Number(k.attrs.span) || 1));
      } else if (k.tag === "col") {
        const cctx = ctxOf(doc, pctx, k);
        const span = Math.min(MAX_SPAN, Math.max(1, Number(k.attrs.span) || 1));
        const wPt = cssPt(cctx.decls.width);
        const w = wPt != null ? wPt / 0.75 : /^\d+$/.test(k.attrs.width ?? "") ? Number(k.attrs.width) : undefined;
        const hidden = isHidden(cctx.decls);
        for (let i = 0; i < span; i++, col++) if (w != null || hidden) sheet.cols.set(col, { width: hidden ? undefined : w != null ? Math.round(w) : undefined, hidden: hidden || undefined });
      }
    }
  };
  cols(table, tableCtx);
  const hasCols = sheet.cols.size > 0;
  // Rows: direct children and those inside row groups, never a nested table's.
  const rows: { tr: HNode; ctx: Ctx }[] = [];
  for (const k of childEls(table)) {
    if (k.tag === "tr") rows.push({ tr: k, ctx: ctxOf(doc, tableCtx, k) });
    else if (k.tag === "thead" || k.tag === "tbody" || k.tag === "tfoot") {
      const gctx = ctxOf(doc, tableCtx, k);
      for (const tr of childEls(k, "tr")) rows.push({ tr, ctx: ctxOf(doc, gctx, tr) });
    }
  }
  // A caption names the sheet and heads it as a first, merged row.
  const caption = childEls(table, "caption")[0];
  const capText = caption ? textOf(caption).replace(/\s+/g, " ").trim() : "";
  if (caption && capText) {
    sheet.name = capText;
    const width = Math.max(1, ...rows.map(({ tr }) => childEls(tr).filter((k) => k.tag === "td" || k.tag === "th").reduce((n, k) => n + Math.min(MAX_SPAN, Math.max(1, Number(k.attrs.colspan) || 1)), 0)));
    const tr: HNode = { tag: "tr", attrs: {}, kids: [], parent: table };
    tr.kids.push({ tag: "td", attrs: { colspan: String(width), style: "text-align:center" }, kids: caption.kids, parent: tr });
    rows.unshift({ tr, ctx: ctxOf(doc, tableCtx, tr) });
  }
  const occupied = new Set<number>();
  const pending: { r: number; c: number; cell: HNode; ctx: Ctx; text: Collected; tableStr: boolean }[] = [];
  const rowHidden: boolean[] = [];
  rows.forEach(({ tr, ctx }, r) => {
    const hidden = isHidden(ctx.decls) || tr.attrs.height === "0";
    rowHidden[r] = hidden;
    const hpt = cssPt(ctx.decls.height);
    const userset = /userset/i.test(ctx.decls["mso-height-source"] ?? "");
    // Excel states every row's height; only the ones set by hand are kept. Other exporters mean theirs.
    const h = doc.excelHeights ? (userset ? hpt : undefined) : /^\d+$/.test(tr.attrs.height ?? "") ? Number(tr.attrs.height) * 0.75 : hpt;
    if (hidden || h != null) sheet.rows.set(r, { height: hidden ? undefined : h, hidden: hidden || undefined });
    let c = 0;
    for (const cell of childEls(tr)) {
      if (cell.tag !== "td" && cell.tag !== "th") continue;
      while (occupied.has(cellKey(r, c))) c++;
      const cctx = ctxOf(doc, ctx, cell);
      const colspan = Math.min(MAX_SPAN, Math.max(1, Number(cell.attrs.colspan) || 1));
      let rowspan = Math.max(0, Number(cell.attrs.rowspan) || 1);
      if (rowspan === 0 || rowspan > rows.length - r) rowspan = Math.max(1, rows.length - r);
      for (let i = 0; i < rowspan; i++) for (let j = 0; j < colspan; j++) occupied.add(cellKey(r + i, c + j));
      const ignore = (cctx.decls["mso-ignore"] ?? "").toLowerCase();
      if ((colspan > 1 && !ignore.includes("colspan")) || (rowspan > 1 && !ignore.includes("rowspan")))
        sheet.merges.push({ r1: r, c1: c, r2: r + (ignore.includes("rowspan") ? 0 : rowspan - 1), c2: c + (ignore.includes("colspan") ? 0 : colspan - 1) });
      if (!hasCols && colspan === 1 && /^\d+$/.test(cell.attrs.width ?? "") && !sheet.cols.has(c)) sheet.cols.set(c, { width: Number(cell.attrs.width) });
      const tableStr = "x:str" in table.attrs || "x:str" in tr.attrs;
      pending.push({ r, c, cell, ctx: cctx, text: collectCell(doc, cctx, cell), tableStr });
      c += colspan;
    }
  });
  const commaDecimal = sniffCommaDecimal(pending.map((p) => p.text.text));
  const dayOrder = commaDecimal ? "dmy" : regional().dayOrder;
  let count = 0;
  for (const p of pending) {
    const a = p.cell.attrs;
    const style = declsToStyle({ ...p.ctx.decls, ...(p.ctx.decls["x-cell-border"] && !p.ctx.decls.border ? { border: p.ctx.decls["x-cell-border"] } : {}) }, { baseSize: 11, excel });
    const cell: TCell = { r: p.r, c: p.c, style };
    const text = p.text.text;
    let v: TValue | undefined;
    const strFmt = style.numFmt === "@";
    if (a["x:fmla"]) cell.f = a["x:fmla"].replace(/^=/, "");
    if (a["x:err"] != null) v = { t: "e", v: a["x:err"] || text };
    else if (a["x:bool"] != null) v = { t: "b", v: /^(true|1)$/i.test(a["x:bool"] || text) };
    else if (a["x:num"] != null) {
      const n = a["x:num"] !== "" ? Number(a["x:num"]) : excelNumber(text, dispDec, dispGroup);
      v = n != null && Number.isFinite(n) ? { t: "n", v: n } : text ? { t: "s", v: text } : undefined;
    } else if (a["x:str"] != null) v = { t: "s", v: a["x:str"] !== "" ? a["x:str"].replace(/^'/, "") : text };
    else if (text) {
      if (ERRORS.has(text.trim())) v = { t: "e", v: text.trim() };
      else if (strFmt || p.tableStr) v = { t: "s", v: text };
      else {
        const typed = excel ? (parseLooseNumber(text, dispDec === ",") ?? typeCellText(text, dispDec === ",", dayOrder)) : typeCellText(text, commaDecimal, dayOrder);
        v = typed.v;
        if (typed.fmt && !style.numFmt) style.numFmt = typed.fmt;
      }
    }
    if (v?.t === "s") {
      // A uniformly styled cell folds its run font into the cell; mixed runs stay rich text.
      const fonts = p.text.runs.map((r) => declsToStyle(r.decls, { baseSize: 11, excel }).font);
      const visible = p.text.runs.map((r, i) => ({ r, f: fonts[i] })).filter((x) => x.r.text.trim());
      if (visible.length && visible.every((x) => sameFont(x.f, visible[0]!.f))) style.font = visible[0]!.f;
      else if (visible.length > 1 && v.v === text) v.runs = p.text.runs.map((r, i): TRun => ({ text: r.text, font: fonts[i] }));
    }
    if (v) cell.v = v;
    if (p.text.link) cell.link = p.text.link;
    if (v || cell.f || hasLook(style)) { sheet.cells.set(cellKey(p.r, p.c), cell); if (v) count++; }
  }
  // Excel ends its tables with hidden empty sizing rows.
  let last = rows.length - 1;
  const used = new Set<number>([...sheet.cells.values()].map((c) => c.r));
  while (last >= 0 && rowHidden[last] && !used.has(last)) { sheet.rows.delete(last); last--; }
  return { sheet, cells: count };
}

const hasLook = (st: TStyle): boolean => !!(st.fill || st.border);

/** Every table of the document that is not inside another one (Excel's case) or holds none (pages). */
function pickTables(root: HNode, excel: boolean): HNode[] {
  const all = findAll(root, "table");
  const nested = (t: HNode): boolean => { for (let p = t.parent; p; p = p.parent) if (p.tag === "table") return true; return false; };
  if (excel) return all.filter((t) => !nested(t));
  return all.filter((t) => findAll(t, "table").length === 0);
}

function contextFor(doc: DocCtx, node: HNode): Ctx {
  const chain: HNode[] = [];
  for (let p: HNode | undefined = node; p && p.tag !== "#root"; p = p.parent) chain.unshift(p);
  let ctx = doc.root;
  for (const n of chain) ctx = ctxOf(doc, ctx, n);
  return ctx;
}

/** Read the tables of one HTML document into sheets. */
function readDocument(text: string, opts: HtmlReadOptions, extraCss: string[] = []): { sheets: TSheet[]; island: Island; tree: HNode; excel: boolean } {
  const excel = /urn:schemas-microsoft-com:office:excel|mso-number-format|\sx:(?:num|str|fmla)\b|ProgId[^>]*Excel\.Sheet/i.test(text.slice(0, 200000));
  const island = readIsland(text);
  const tree = parseHtmlTree(text, opts.useDom !== false);
  const rules: Rule[] = [];
  for (const css of extraCss) parseCss(css, rules);
  for (const st of findAll(tree, "style")) parseCss(textOf(st), rules);
  const doc = newDoc(rules, excel, /mso-height-source/i.test(text));
  const sheets: TSheet[] = [];
  const tables = pickTables(tree, excel);
  // Exports from other tools often title the table or add notes around it, as plain page text.
  const around = textAround(tree, new Set(tables));
  tables.forEach((t, i) => {
    const res = readTable(doc, t, contextFor(doc, t), "");
    if (!res.cells && !excel) return;
    const before = around.before[i] ?? [], after = i === tables.length - 1 ? around.after : [];
    if (before.length || after.length) addLines(res.sheet, before, after);
    sheets.push(res.sheet);
  });
  return { sheets, island, tree, excel };
}

/** The page's text lines before each table (since the previous one) and after the last one. */
function textAround(tree: HNode, tables: Set<HNode>): { before: string[][]; after: string[] } {
  const before: string[][] = [];
  let lines: string[] = [], cur = "";
  const flush = (): void => { const t = cur.replace(/[\s\u00a0]+/g, " ").trim(); if (/[\p{L}\p{N}]/u.test(t)) lines.push(t); cur = ""; };
  const walk = (n: HNode): void => {
    for (const k of n.kids) {
      if (typeof k === "string") { cur += k; continue; }
      if (SKIP.has(k.tag) || /display\s*:\s*none/i.test(k.attrs.style ?? "")) continue;
      if (tables.has(k)) { flush(); before.push(lines.slice(-MAX_LINES)); lines = []; continue; }
      if (k.tag === "br") { flush(); continue; }
      const block = BLOCK.has(k.tag) || k.tag === "td" || k.tag === "th" || k.tag === "center";
      if (block) flush();
      walk(k);
      if (block) flush();
    }
  };
  walk(tree);
  flush();
  return { before, after: lines.slice(0, MAX_LINES) };
}

/** Put text lines above and below a sheet's table, shifting the table down. */
function addLines(sheet: TSheet, before: string[], after: string[]): void {
  const shift = before.length;
  if (shift) {
    const cells = [...sheet.cells.values()];
    sheet.cells.clear();
    for (const c of cells) { c.r += shift; sheet.cells.set(cellKey(c.r, c.c), c); }
    for (const m of sheet.merges) { m.r1 += shift; m.r2 += shift; }
    const rows = [...sheet.rows];
    sheet.rows.clear();
    for (const [r, info] of rows) sheet.rows.set(r + shift, info);
    if (sheet.autoFilter) { sheet.autoFilter.r1 += shift; sheet.autoFilter.r2 += shift; }
  }
  const put = (r: number, text: string): void => { sheet.cells.set(cellKey(r, 0), { r, c: 0, v: { t: "s", v: text } }); };
  before.forEach((t, i) => put(i, t));
  let last = -1;
  for (const c of sheet.cells.values()) if (c.r > last) last = c.r;
  for (const r of sheet.rows.keys()) if (r > last) last = r;
  after.forEach((t, i) => put(last + 1 + i, t));
}

/** Read an HTML (or MHTML) spreadsheet export into the intermediate book. */
export function readHtmlBook(text: string, opts: HtmlReadOptions = {}): TBook {
  const base = opts.sheetName ?? ((n: number) => `Sheet${n}`);
  const mime = /^\s*(?:MIME-Version:|Content-Type:\s*multipart\/related)/i.test(text) ? parseMhtml(text) : null;
  const mainText = mime ? (mime.find((p) => p.type === "text/html")?.text ?? "") : text;
  const css = mime ? mime.filter((p) => p.type === "text/css").map((p) => p.text) : [];
  const main = readDocument(mainText, opts, css);
  const book: TBook = { sheets: [], names: [] };
  const isl = main.island;
  if (isl.sheets.some((s) => s.source)) {
    // A multi-sheet web page: each sheet lives in its own file, read when the file carries it.
    for (const s of isl.sheets) {
      const part = s.source && mime?.find((p) => p.location && baseName(p.location) === baseName(s.source!));
      const sub = part ? readDocument(part.text, opts, css) : null;
      const sheet = sub?.sheets[0] ?? newSheet("");
      sheet.name = s.name;
      book.sheets.push(sheet);
    }
  } else {
    book.sheets = main.sheets;
    if (main.excel && isl.sheets.length) isl.sheets.forEach((s, i) => { if (book.sheets[i]) book.sheets[i]!.name = s.name; });
  }
  if (!book.sheets.length) book.sheets.push(newSheet(""));
  // Generic pages may name a sheet after its table's caption; the rest take the default names.
  const names = uniqueSheetNames(book.sheets.map((s) => s.name), base);
  book.sheets.forEach((s, i) => {
    s.name = names[i]!;
    const info = isl.sheets[i];
    if (!info) return;
    if (info.freeze) s.freeze = info.freeze;
    if (info.state) s.state = info.state;
    if (info.tab) s.tabColor = info.tab;
    if (info.hideGrid) s.hideGrid = true;
    if (info.hideZeros) s.hideZeros = true;
  });
  if (isl.active != null && isl.active < book.sheets.length) book.activeSheet = isl.active;
  else { const sel = isl.sheets.findIndex((s) => s.selected); if (sel > 0 && sel < book.sheets.length) book.activeSheet = sel; }
  for (const n of isl.names) {
    const builtin = /^(Print_Area|Print_Titles|_FilterDatabase|Criteria|Extract)$/i.test(n.name);
    const name: TName = { name: builtin ? "_xlnm." + n.name : n.name, ref: n.formula, sheet: n.sheet, hidden: n.hidden || undefined };
    if (n.sheet != null && (n.sheet < 0 || n.sheet >= book.sheets.length)) continue;
    book.names.push(name);
    if (/^_FilterDatabase$/i.test(n.name) && n.sheet != null) {
      const m = /!\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$/.exec(n.formula);
      if (m) book.sheets[n.sheet]!.autoFilter = { r1: Number(m[2]) - 1, c1: colIndex(m[1]!), r2: Number(m[4]) - 1, c2: colIndex(m[3]!) };
    }
  }
  book.defaultFont = commonFont(book);
  return book;
}

const colIndex = (letters: string): number => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** The most used font becomes the workbook's normal font; cells then state only what differs. */
function commonFont(book: TBook): { name: string; size: number } {
  const counts = new Map<string, number>();
  for (const s of book.sheets) for (const c of s.cells.values()) {
    if (!c.v) continue;
    const k = `${c.style?.font?.name ?? ""}|${c.style?.font?.size ?? ""}`;
    counts.set(k, (counts.get(k) ?? 0) + 1);
  }
  let best = "|", n = -1;
  for (const [k, v] of counts) if (v > n) { best = k; n = v; }
  const [name, size] = best.split("|");
  const dflt = { name: name || "Calibri", size: Number(size) || 11 };
  const strip = (f: TFont | undefined): TFont | undefined => {
    if (!f) return f;
    if (f.name === dflt.name) delete f.name;
    if (f.size === dflt.size) delete f.size;
    return Object.keys(f).length ? f : undefined;
  };
  for (const s of book.sheets) for (const c of s.cells.values()) {
    if (c.style) { c.style.font = strip(c.style.font); if (!c.style.font) delete c.style.font; }
    if (c.v?.t === "s" && c.v.runs) for (const r of c.v.runs) r.font = strip(r.font ? { ...r.font } : undefined);
  }
  return dflt;
}
