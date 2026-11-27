// The slice of CSS that spreadsheet HTML exports use: rules, declarations, colours, lengths, and the
// mapping of the resulting declarations onto a cell style.
import type { TBorder, TFont, TStyle } from "./model";

export type Decls = Record<string, string>;

/** Undo CSS escapes: "\0022" is a quote, "\@" an at sign. */
export function cssUnescape(s: string): string {
  return s.replace(/\\([0-9a-fA-F]{1,6})\s?|\\([^\n])/g, (_m, hex: string | undefined, ch: string | undefined) =>
    hex ? String.fromCodePoint(parseInt(hex, 16) || 0xfffd) : ch!);
}

/** Split "a:b; c:'d;e'" into declarations, keeping quoted and bracketed parts whole. */
export function parseDecls(text: string, out: Decls = {}): Decls {
  let cur = "", q = "", depth = 0;
  const flush = (): void => {
    const i = cur.indexOf(":");
    if (i > 0) {
      const prop = cur.slice(0, i).trim().toLowerCase();
      const val = cur.slice(i + 1).replace(/!important\s*$/i, "").trim();
      if (prop && val) out[prop] = val;
    }
    cur = "";
  };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (q) { if (ch === "\\" && i + 1 < text.length) { cur += ch + text[++i]; continue; } if (ch === q) q = ""; cur += ch; continue; }
    if (ch === '"' || ch === "'") { q = ch; cur += ch; continue; }
    if (ch === "(") depth++;
    else if (ch === ")") depth = Math.max(0, depth - 1);
    if (ch === ";" && !depth) { flush(); continue; }
    cur += ch;
  }
  flush();
  return out;
}

export interface Compound { tag?: string; classes: string[]; id?: string }
export interface Rule { chain: { sel: Compound; comb: " " | ">" }[]; spec: number; order: number; decls: Decls }

function parseCompound(s: string): Compound | null {
  const c: Compound = { classes: [] };
  const re = /([.#]?)(-?[A-Za-z_][\w-]*)|:[\w-]+(?:\([^)]*\))?|(\*)/g;
  let pos = 0;
  for (let m = re.exec(s); m; m = re.exec(s)) {
    if (m.index !== pos) return null;
    pos = re.lastIndex;
    if (m[3] || m[0][0] === ":") continue; // universal selector, pseudo-classes (a:link reads as a)
    if (m[1] === ".") c.classes.push(m[2]!.toLowerCase());
    else if (m[1] === "#") c.id = m[2]!.toLowerCase();
    else c.tag = m[2]!.toLowerCase();
  }
  return pos === s.length ? c : null;
}

/** Parse a style sheet's rules; at-rules (@page, @media) are skipped. */
export function parseCss(text: string, rules: Rule[] = []): Rule[] {
  const src = text.replace(/\/\*[\s\S]*?\*\//g, "").replace(/<!--|-->/g, " ");
  let i = 0;
  while (i < src.length) {
    const open = src.indexOf("{", i);
    if (open < 0) break;
    const selText = src.slice(i, open).trim();
    // Find the matching close brace (at-rules may nest blocks).
    let depth = 1, j = open + 1;
    while (j < src.length && depth) { if (src[j] === "{") depth++; else if (src[j] === "}") depth--; j++; }
    const body = src.slice(open + 1, j - 1);
    i = j;
    if (!selText || selText.includes("@")) continue;
    const decls = parseDecls(body);
    for (const one of selText.split(",")) {
      const parts = one.trim().replace(/\s*>\s*/g, " > ").split(/\s+/);
      const chain: Rule["chain"] = [];
      let comb: " " | ">" = " ", ok = parts.length > 0;
      for (const p of parts) {
        if (p === ">") { comb = ">"; continue; }
        const sel = parseCompound(p);
        if (!sel) { ok = false; break; }
        chain.push({ sel, comb });
        comb = " ";
      }
      if (!ok || !chain.length) continue;
      let spec = 0;
      for (const { sel } of chain) spec += (sel.id ? 10000 : 0) + sel.classes.length * 100 + (sel.tag ? 1 : 0);
      rules.push({ chain, spec, order: rules.length, decls });
    }
  }
  return rules;
}

// ---- colours ----------------------------------------------------------------------------------

const NAMED = "aliceblue:f0f8ff,antiquewhite:faebd7,aqua:00ffff,aquamarine:7fffd4,azure:f0ffff,beige:f5f5dc,bisque:ffe4c4,black:000000,blanchedalmond:ffebcd,blue:0000ff,blueviolet:8a2be2,brown:a52a2a,burlywood:deb887,cadetblue:5f9ea0,chartreuse:7fff00,chocolate:d2691e,coral:ff7f50,cornflowerblue:6495ed,cornsilk:fff8dc,crimson:dc143c,cyan:00ffff,darkblue:00008b,darkcyan:008b8b,darkgoldenrod:b8860b,darkgray:a9a9a9,darkgreen:006400,darkgrey:a9a9a9,darkkhaki:bdb76b,darkmagenta:8b008b,darkolivegreen:556b2f,darkorange:ff8c00,darkorchid:9932cc,darkred:8b0000,darksalmon:e9967a,darkseagreen:8fbc8f,darkslateblue:483d8b,darkslategray:2f4f4f,darkslategrey:2f4f4f,darkturquoise:00ced1,darkviolet:9400d3,deeppink:ff1493,deepskyblue:00bfff,dimgray:696969,dimgrey:696969,dodgerblue:1e90ff,firebrick:b22222,floralwhite:fffaf0,forestgreen:228b22,fuchsia:ff00ff,gainsboro:dcdcdc,ghostwhite:f8f8ff,gold:ffd700,goldenrod:daa520,gray:808080,green:008000,greenyellow:adff2f,grey:808080,honeydew:f0fff0,hotpink:ff69b4,indianred:cd5c5c,indigo:4b0082,ivory:fffff0,khaki:f0e68c,lavender:e6e6fa,lavenderblush:fff0f5,lawngreen:7cfc00,lemonchiffon:fffacd,lightblue:add8e6,lightcoral:f08080,lightcyan:e0ffff,lightgoldenrodyellow:fafad2,lightgray:d3d3d3,lightgreen:90ee90,lightgrey:d3d3d3,lightpink:ffb6c1,lightsalmon:ffa07a,lightseagreen:20b2aa,lightskyblue:87cefa,lightslategray:778899,lightslategrey:778899,lightsteelblue:b0c4de,lightyellow:ffffe0,lime:00ff00,limegreen:32cd32,linen:faf0e6,magenta:ff00ff,maroon:800000,mediumaquamarine:66cdaa,mediumblue:0000cd,mediumorchid:ba55d3,mediumpurple:9370db,mediumseagreen:3cb371,mediumslateblue:7b68ee,mediumspringgreen:00fa9a,mediumturquoise:48d1cc,mediumvioletred:c71585,midnightblue:191970,mintcream:f5fffa,mistyrose:ffe4e1,moccasin:ffe4b5,navajowhite:ffdead,navy:000080,oldlace:fdf5e6,olive:808000,olivedrab:6b8e23,orange:ffa500,orangered:ff4500,orchid:da70d6,palegoldenrod:eee8aa,palegreen:98fb98,paleturquoise:afeeee,palevioletred:db7093,papayawhip:ffefd5,peachpuff:ffdab9,peru:cd853f,pink:ffc0cb,plum:dda0dd,powderblue:b0e0e6,purple:800080,rebeccapurple:663399,red:ff0000,rosybrown:bc8f8f,royalblue:4169e1,saddlebrown:8b4513,salmon:fa8072,sandybrown:f4a460,seagreen:2e8b57,seashell:fff5ee,sienna:a0522d,silver:c0c0c0,skyblue:87ceeb,slateblue:6a5acd,slategray:708090,slategrey:708090,snow:fffafa,springgreen:00ff7f,steelblue:4682b4,tan:d2b48c,teal:008080,thistle:d8bfd8,tomato:ff6347,turquoise:40e0d0,violet:ee82ee,wheat:f5deb3,white:ffffff,whitesmoke:f5f5f5,yellow:ffff00,yellowgreen:9acd32,windowtext:000000,window:ffffff,buttonface:f0f0f0,infobackground:ffffe1";
let named: Map<string, string> | undefined;

/** A CSS colour as "#rrggbb"; "auto", "transparent" and unknown values give undefined. */
export function cssColor(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const s = v.trim().toLowerCase().replace(/^["']|["']$/g, "");
  let m = /^#?([0-9a-f]{6})$/.exec(s);
  if (m) return "#" + m[1];
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(s);
  if (m) return "#" + m[1]! + m[1] + m[2] + m[2] + m[3] + m[3];
  m = /^rgba?\(\s*(\d+%?)\s*,?\s*(\d+%?)\s*,?\s*(\d+%?)\s*(?:[,/]\s*([\d.]+%?))?\s*\)$/.exec(s);
  if (m) {
    if (m[4] != null && parseFloat(m[4]) === 0) return undefined;
    const ch = (x: string): string => Math.max(0, Math.min(255, x.endsWith("%") ? Math.round(parseFloat(x) * 2.55) : Number(x))).toString(16).padStart(2, "0");
    return "#" + ch(m[1]!) + ch(m[2]!) + ch(m[3]!);
  }
  if (!named) named = new Map(NAMED.split(",").map((p) => p.split(":") as [string, string]));
  const hex = named.get(s);
  return hex ? "#" + hex : undefined;
}

/** A CSS length in points ("10pt", "13px", ".5pt"); `rel` is the size "%"/"em" scale against. */
export function cssPt(v: string | undefined, rel?: number): number | undefined {
  if (!v) return undefined;
  const m = /^(-?[\d.]+)\s*(pt|px|em|rem|%|in|cm|mm)?$/i.exec(v.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]!);
  if (!Number.isFinite(n)) return undefined;
  switch ((m[2] ?? "px").toLowerCase()) {
    case "pt": return n;
    case "px": return n * 0.75;
    case "in": return n * 72;
    case "cm": return (n * 72) / 2.54;
    case "mm": return (n * 72) / 25.4;
    case "%": return rel != null ? (rel * n) / 100 : undefined;
    default: return rel != null ? rel * n : undefined;
  }
}

const FONT_KEYWORDS: Record<string, number> = { "xx-small": 7, "x-small": 7.5, small: 10, medium: 12, large: 13.5, "x-large": 18, "xx-large": 24 };
const HTML_FONT_SIZES = [0, 7.5, 10, 12, 13.5, 18, 24, 36];

/** <font size=N> (1-7, or +n/-n relative to 3) in points. */
export function htmlFontSize(v: string): number | undefined {
  const m = /^\s*([+-]?)(\d)/.exec(v);
  if (!m) return undefined;
  const n = Number(m[2]);
  const idx = m[1] === "+" ? 3 + n : m[1] === "-" ? 3 - n : n;
  return HTML_FONT_SIZES[Math.max(1, Math.min(7, idx))];
}

// ---- number formats ---------------------------------------------------------------------------

const NAMED_FMT: Record<string, string> = {
  "general": "General", "general number": "General", "general date": "m/d/yyyy h:mm", "long date": "dddd, mmmm d, yyyy",
  "medium date": "d-mmm-yy", "short date": "m/d/yyyy", "long time": "h:mm:ss AM/PM", "medium time": "h:mm AM/PM", "short time": "h:mm",
  "currency": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', "euro currency": '[$€-2]\\ #,##0.00', "fixed": "0.00", "standard": "#,##0.00",
  "percent": "0.00%", "scientific": "0.00E+00", "yes/no": '"Yes";"Yes";"No"', "true/false": '"True";"True";"False"', "on/off": '"On";"On";"Off"',
  "@": "@", "text": "@",
};

/** An Excel number format from a SpreadsheetML ss:Format or a CSS mso-number-format value. */
export function excelNumFmt(raw: string | undefined): string | undefined {
  if (raw == null) return undefined;
  let s = raw.trim();
  if (/^".*"$|^'.*'$/.test(s)) s = s.slice(1, -1);
  s = cssUnescape(s);
  if (!s) return undefined;
  return NAMED_FMT[s.toLowerCase()] ?? s;
}

// ---- declarations to a cell style ---------------------------------------------------------------

const BORDER_WIDTH = (w: number, style: string): string => {
  if (style === "double") return "double";
  if (style === "dotted") return w >= 1.5 ? "mediumDashDot" : w < 0.75 ? "hair" : "dotted";
  if (style === "dashed") return w >= 1.5 ? "mediumDashed" : "dashed";
  if (style === "dot-dash") return w >= 1.5 ? "mediumDashDot" : "dashDot";
  if (style === "dot-dot-dash") return w >= 1.5 ? "mediumDashDotDot" : "dashDotDot";
  if (w >= 2.25) return "thick";
  if (w >= 1.5) return "medium";
  if (w > 0 && w < 0.5) return "hair";
  return "thin";
};

/** A CSS border shorthand ("0.5pt solid windowtext", "1px solid #000", "none"). */
export function cssBorder(v: string | undefined): TBorder | undefined {
  if (!v) return undefined;
  const s = v.trim().toLowerCase();
  if (/^(none|hidden|0|0pt|0px|medium none|initial)$/.test(s) || /\bnone\b/.test(s) && !/solid|dotted|dashed|double/.test(s)) return { style: "none" };
  let width = 0.75, style = "", color: string | undefined;
  for (const tok of s.match(/rgba?\([^)]*\)|\S+/g) ?? []) {
    if (/^(solid|dotted|dashed|double|groove|ridge|inset|outset|dot-dash|dot-dot-dash)$/.test(tok)) style = tok;
    else if (/^(thin|medium|thick)$/.test(tok)) width = tok === "thin" ? 0.75 : tok === "medium" ? 1.5 : 2.25;
    else if (/^[\d.]+(pt|px|em)?$/.test(tok)) width = cssPt(tok.replace(/^([\d.]+)$/, "$1px")) ?? width;
    else color = cssColor(tok) ?? color;
  }
  if (!style) style = "solid";
  if (width === 0) return { style: "none" };
  if (/groove|ridge|inset|outset/.test(style)) style = "solid";
  return { style: BORDER_WIDTH(width, style), color: color && color !== "#000000" ? color : undefined };
}

const H_ALIGN: Record<string, string> = { left: "left", right: "right", center: "center", centre: "center", justify: "justify", "center-across": "centerContinuous", fill: "fill", distributed: "distributed", start: "left", end: "right", "-webkit-center": "center", middle: "center" };
const V_ALIGN: Record<string, string> = { top: "top", middle: "center", center: "center", bottom: "bottom", baseline: "bottom", justify: "justify", distributed: "distributed" };

/** Map the effective declarations of a cell onto a style; `base` holds the defaults to leave out. */
export function declsToStyle(d: Decls, opts: { baseSize: number; excel: boolean }): TStyle {
  const st: TStyle = {};
  const font: TFont = {};
  const fam = d["font-family"];
  if (fam) {
    const first = fam.split(",")[0]!.trim().replace(/^["']|["']$/g, "");
    if (first && !/^(serif|sans-serif|monospace|cursive|fantasy|system-ui|inherit|initial)$/i.test(first)) font.name = first;
  }
  const size = d["font-size"] ? FONT_KEYWORDS[d["font-size"].toLowerCase()] ?? cssPt(d["font-size"], opts.baseSize) : undefined;
  if (size) font.size = Math.round(size * 2) / 2;
  const fw = d["font-weight"]?.toLowerCase();
  if (fw && (fw === "bold" || fw === "bolder" || Number(fw) >= 600)) font.bold = true;
  if (/italic|oblique/i.test(d["font-style"] ?? "")) font.italic = true;
  const deco = (d["text-decoration"] ?? "") + " " + (d["text-decoration-line"] ?? "");
  if (/underline/i.test(deco)) {
    const u = (d["text-underline-style"] ?? "").toLowerCase();
    font.underline = u === "double" ? "double" : u === "single-accounting" ? "singleAccounting" : u === "double-accounting" ? "doubleAccounting" : "single";
  }
  if (/line-through/i.test(deco)) font.strike = true;
  const va = d["vertical-align"]?.toLowerCase();
  if (va === "super") font.vertAlign = "superscript";
  else if (va === "sub") font.vertAlign = "subscript";
  const color = cssColor(d["color"]);
  if (color && color !== "#000000") font.color = color;
  if (Object.keys(font).length) st.font = font;

  const bgVal = d["background-color"] ?? d["background"];
  const bg = bgVal ? cssColor(bgVal.split(/\s+(?![^(]*\))/).find((t) => cssColor(t)) ?? bgVal) : undefined;
  const pattern = d["mso-pattern"]?.toLowerCase();
  if (bg) {
    const pm = pattern && /^(\S+)\s+(gray-\d+(?:\.\d+)?|\S+)$/.exec(pattern);
    const kind = pm && pm[2] !== "none" ? MSO_PATTERN[pm[2]!] : undefined;
    st.fill = kind ? { pattern: kind, fg: cssColor(pm![1]) ?? "#000000", bg } : { pattern: "solid", fg: bg };
  }

  const border: NonNullable<TStyle["border"]> = {};
  const all = cssBorder(d["border"]);
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const b = cssBorder(d[`border-${side}`]) ?? all;
    if (b && b.style !== "none") border[side] = b;
  }
  const dd = cssBorder(d["mso-diagonal-down"]), du = cssBorder(d["mso-diagonal-up"]);
  if (dd && dd.style !== "none") { border.diagonal = dd; border.diagDown = true; }
  if (du && du.style !== "none") { border.diagonal = du; border.diagUp = true; }
  if (Object.keys(border).length) st.border = border;

  const align: NonNullable<TStyle["align"]> = {};
  const ta = d["text-align"]?.toLowerCase().replace(/^["']|["']$/g, "");
  if (ta && H_ALIGN[ta]) align.h = H_ALIGN[ta];
  const vv = va ? V_ALIGN[va] : undefined;
  if (vv && vv !== "bottom") align.v = vv;
  // Excel writes "white-space:normal" for wrapped cells and nowrap for the rest.
  if (opts.excel && /^(normal|pre-wrap)$/i.test(d["white-space"] ?? "")) align.wrap = true;
  const rot = Number(d["mso-rotate"]);
  if (rot) align.rotate = rot > 0 ? Math.min(90, rot) : Math.min(180, 90 - rot);
  if (/vertical/i.test(d["layout-flow"] ?? "")) align.rotate = 255;
  const indent = Number(d["mso-char-indent-count"]);
  if (indent > 0) align.indent = Math.min(15, Math.round(indent));
  if (Object.keys(align).length) st.align = align;

  const fmt = excelNumFmt(d["mso-number-format"]);
  if (fmt && fmt !== "General") st.numFmt = fmt;
  if (/unlocked/i.test(d["mso-protection"] ?? "")) st.locked = false;
  if (/hidden/i.test(d["mso-protection"] ?? "")) st.hidden = true;
  return st;
}

const MSO_PATTERN: Record<string, string> = {
  "gray-50": "mediumGray", "gray-75": "darkGray", "gray-25": "lightGray", "gray-125": "gray125", "gray-0625": "gray0625",
  "horz-stripe": "darkHorizontal", "vert-stripe": "darkVertical", "reverse-diag-stripe": "darkDown", "diag-stripe": "darkUp",
  "diag-cross": "darkGrid", "thick-diag-cross": "darkTrellis", "thin-horz-stripe": "lightHorizontal", "thin-vert-stripe": "lightVertical",
  "thin-reverse-diag-stripe": "lightDown", "thin-diag-stripe": "lightUp", "thin-horz-cross": "lightGrid", "thin-diag-cross": "lightTrellis",
};
