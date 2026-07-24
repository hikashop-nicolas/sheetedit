import { cellDisplay, key, type CfDxf, type CfRule, type CondFormat, type Sheet } from "../../core/model";
import type { FormulaEvaluator } from "../../core/recalc";
import { argbToCss, resolveColor } from "./read";

/** Context that lets aggregate/expression rules evaluate cell-ref / formula operands. */
export interface CfEvalCtx { evaluator: FormulaEvaluator; sheetName: string; r0: number; c0: number }

// Conditional formatting: parse the differential formats (dxfs) and cfRules, then evaluate them
// per cell into a visual (fill / text colour / bold / a data bar / a colour-scale fill). A
// supported subset: cellIs and text comparisons, top/bottom N, above/below average, duplicate/
// unique values (all applying a dxf), plus colour scales, data bars and icon sets. cellIs formula
// operands and is-true-formula (expression) rules evaluate through the workbook's formula engine
// when an evaluator is supplied. Time-period rules are left to round-trip untouched (not rendered).

const firstByLocal = (parent: Element, local: string): Element | undefined =>
  Array.from(parent.children).find((e) => e.localName === local);
const childrenByLocal = (parent: Element, local: string): Element[] =>
  Array.from(parent.children).filter((e) => e.localName === local);

/** Parse styles.xml <dxfs> into an array indexed by dxfId. */
export function parseDxfs(stylesDoc: Document | undefined, theme: string[]): CfDxf[] {
  if (!stylesDoc) return [];
  const dxfsEl = stylesDoc.getElementsByTagName("dxfs")[0];
  if (!dxfsEl) return [];
  return childrenByLocal(dxfsEl, "dxf").map((dxf) => {
    const out: CfDxf = {};
    const font = firstByLocal(dxf, "font");
    if (font) {
      if (firstByLocal(font, "b")) out.bold = true;
      if (firstByLocal(font, "i")) out.italic = true;
      out.color = resolveColor(firstByLocal(font, "color"), theme);
    }
    const fill = firstByLocal(dxf, "fill");
    const pat = fill && firstByLocal(fill, "patternFill");
    // dxf solid fills carry the colour in bgColor (a well-known OOXML quirk); fall back to fgColor.
    if (pat) out.bg = resolveColor(firstByLocal(pat, "bgColor"), theme) ?? resolveColor(firstByLocal(pat, "fgColor"), theme);
    return out;
  });
}

const parseRect = (ref: string): { r1: number; c1: number; r2: number; c2: number } | null => {
  const m = (s: string): { r: number; c: number } | null => {
    const mm = /^\$?([A-Z]+)\$?(\d+)$/.exec(s);
    if (!mm) return null;
    let col = 0;
    for (const ch of mm[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { r: Number(mm[2]), c: col };
  };
  const [a, b] = ref.split(":");
  const p1 = m(a ?? "");
  const p2 = b ? m(b) : p1;
  if (!p1 || !p2) return null;
  return { r1: Math.min(p1.r, p2.r), c1: Math.min(p1.c, p2.c), r2: Math.max(p1.r, p2.r), c2: Math.max(p1.c, p2.c) };
};

const cfvoOf = (el: Element): { type: string; val?: number } => {
  const val = el.getAttribute("val");
  return { type: el.getAttribute("type") || "num", val: val != null && val !== "" ? Number(val) : undefined };
};

/** Read the sheet's <conditionalFormatting> blocks into sheet.condFormats. */
export function readCondFormats(sheet: Sheet, doc: Document, dxfs: CfDxf[], theme: string[]): void {
  const blocks = doc.getElementsByTagName("conditionalFormatting");
  if (!blocks.length) return;
  const out: CondFormat[] = [];
  for (const block of Array.from(blocks)) {
    const sqref = block.getAttribute("sqref");
    if (!sqref) continue;
    const ranges = sqref.split(/\s+/).map(parseRect).filter((r): r is NonNullable<typeof r> => r != null);
    if (!ranges.length) continue;
    const rules: CfRule[] = [];
    for (const rule of childrenByLocal(block, "cfRule")) {
      const type = rule.getAttribute("type") || "";
      const priority = Number(rule.getAttribute("priority") || "0");
      const stopIfTrue = rule.getAttribute("stopIfTrue") === "1";
      const dxfId = rule.getAttribute("dxfId");
      const dxf = dxfId != null ? dxfs[Number(dxfId)] : undefined;
      const formulas = childrenByLocal(rule, "formula").map((f) => f.textContent?.trim() ?? "");
      const r: CfRule = { type, priority, stopIfTrue, dxf, formulas, operator: rule.getAttribute("operator") || undefined, text: rule.getAttribute("text") || undefined };
      if (type === "top10") { r.rank = Number(rule.getAttribute("rank") || "10"); r.percent = rule.getAttribute("percent") === "1"; r.bottom = rule.getAttribute("bottom") === "1"; }
      if (type === "aboveAverage") { r.aboveAverage = rule.getAttribute("aboveAverage") !== "0"; r.equalAverage = rule.getAttribute("equalAverage") === "1"; }
      if (type === "colorScale") {
        const cs = firstByLocal(rule, "colorScale");
        if (cs) r.colorScale = { cfvo: childrenByLocal(cs, "cfvo").map(cfvoOf), colors: childrenByLocal(cs, "color").map((c) => argbToCss(c.getAttribute("rgb")) ?? resolveColor(c, theme) ?? "#ffffff") };
      }
      if (type === "dataBar") {
        const db = firstByLocal(rule, "dataBar");
        const cfvos = db ? childrenByLocal(db, "cfvo").map(cfvoOf) : [];
        const color = db ? (argbToCss(firstByLocal(db, "color")?.getAttribute("rgb")) ?? resolveColor(firstByLocal(db, "color"), theme) ?? "#638ec6") : "#638ec6";
        if (cfvos.length >= 2) r.dataBar = { min: cfvos[0], max: cfvos[cfvos.length - 1], color };
      }
      if (type === "iconSet") {
        const is = firstByLocal(rule, "iconSet");
        const cfvo = is ? childrenByLocal(is, "cfvo").map((el) => ({ ...cfvoOf(el), gte: el.getAttribute("gte") !== "0" })) : [];
        if (cfvo.length) r.iconSet = { set: is?.getAttribute("iconSet") || "3TrafficLights1", cfvo, reverse: is?.getAttribute("reverse") === "1" };
      }
      rules.push(r);
    }
    rules.sort((a, b) => a.priority - b.priority); // lower priority number wins
    if (rules.length) out.push({ ranges, rules });
  }
  if (out.length) sheet.condFormats = out;
}

// --- evaluation -------------------------------------------------------------

export interface CfVisual { bg?: string; color?: string; bold?: boolean; italic?: boolean; bar?: { pct: number; color: string }; icon?: string }

const numOf = (s: string): number | null => { const n = Number(s.replace(/,/g, "")); return s.trim() !== "" && Number.isFinite(n) ? n : null; };
const lerp = (a: number, b: number, t: number): number => Math.round(a + (b - a) * t);
const toRgb = (css: string): [number, number, number] => {
  const h = css.replace("#", "");
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
};
const percentile = (sorted: number[], p: number): number => {
  if (!sorted.length) return 0;
  const idx = (p / 100) * (sorted.length - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
};
const stopValue = (c: { type: string; val?: number }, min: number, max: number, sorted: number[]): number => {
  switch (c.type) {
    case "min": return min;
    case "max": return max;
    case "percent": return min + ((c.val ?? 0) / 100) * (max - min);
    case "percentile": return percentile(sorted, c.val ?? 0);
    default: return c.val ?? 0; // num / formula (literal)
  }
};

// Render an icon-set glyph as a small inline SVG. Faithful for the common families (arrows,
// symbols/signs, ratings/quarters, traffic lights); other sets fall back to a colour ramp circle.
// index is the value's bucket (0 = lowest), count the number of icons in the set.
export function cfIconSvg(set: string, index: number, count: number): string {
  const gray = /Gray|Grey|Black/.test(set);
  const hue = count > 1 ? Math.round((index / (count - 1)) * 120) : 120; // 0=red .. 120=green
  const col = gray ? "#808080" : `hsl(${hue},60%,42%)`;
  const wrap = (inner: string): string => `<svg viewBox="0 0 14 14" width="12" height="12" aria-hidden="true">${inner}</svg>`;
  if (/Arrows/.test(set)) {
    const rot = count > 1 ? 180 * (1 - index / (count - 1)) : 0; // lowest points down, highest up
    return wrap(`<g transform="rotate(${rot} 7 7)" stroke="${col}" stroke-width="1.6" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M7 12V3M3.6 6.4 7 3l3.4 3.4"/></g>`);
  }
  if (/Symbols|Signs/.test(set)) {
    if (index === 0) return wrap(`<path d="M4 4l6 6M10 4l-6 6" stroke="${col}" stroke-width="1.8" stroke-linecap="round"/>`); // cross
    if (index === count - 1) return wrap(`<path d="M3.5 7.5 6 10l4.5-5.5" stroke="${col}" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`); // check
    return wrap(`<g stroke="${col}" stroke-width="1.8" stroke-linecap="round"><path d="M7 3.5v4.2"/></g><circle cx="7" cy="10.8" r="1" fill="${col}"/>`); // exclamation
  }
  if (/Rating|Quarters|Stars|Boxes/.test(set)) {
    const frac = count > 1 ? index / (count - 1) : 1;
    let pie = "";
    if (frac >= 1) pie = `<circle cx="7" cy="7" r="5" fill="#555"/>`;
    else if (frac > 0) {
      const a = frac * 2 * Math.PI - Math.PI / 2;
      const x = (7 + 5 * Math.cos(a)).toFixed(2), y = (7 + 5 * Math.sin(a)).toFixed(2);
      pie = `<path d="M7 7 L7 2 A5 5 0 ${frac > 0.5 ? 1 : 0} 1 ${x} ${y} Z" fill="#555"/>`;
    }
    return wrap(`<circle cx="7" cy="7" r="5" fill="none" stroke="#888" stroke-width="1"/>${pie}`);
  }
  if (/Flags/.test(set)) return wrap(`<g stroke="#555" stroke-width="1"><path d="M4 2v10" stroke="${col}"/></g><path d="M4.6 2.5h6l-1.4 2 1.4 2h-6z" fill="${col}"/>`);
  return wrap(`<circle cx="7" cy="7" r="5.2" fill="${col}"/>`); // traffic lights / default
}

/** The 0-based icon bucket for a value given the set's thresholds. */
function iconIndex(v: number, rule: NonNullable<CfRule["iconSet"]>, min: number, max: number, sorted: number[]): number {
  // cfvo[0] is the implicit lower bound; cfvo[1..] are the thresholds between icons.
  let idx = 0;
  for (let i = 1; i < rule.cfvo.length; i++) {
    const t = stopValue(rule.cfvo[i], min, max, sorted);
    if (rule.cfvo[i].gte === false ? v > t : v >= t) idx = i;
    else break;
  }
  return rule.reverse ? rule.cfvo.length - 1 - idx : idx;
}

/** Compute the conditional-format visual for every affected cell of the sheet, once per render. */
export function computeCondVisuals(sheet: Sheet, ev?: { evaluator: FormulaEvaluator; sheetName: string }): Map<string, CfVisual> {
  const out = new Map<string, CfVisual>();
  for (const cf of sheet.condFormats ?? []) {
    const inRange = (r: number, c: number): boolean => cf.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2);
    // Formula operands / expression rules evaluate relative to the range's top-left origin.
    const origin = cf.ranges.reduce((a, g) => (g.r1 < a.r1 || (g.r1 === a.r1 && g.c1 < a.c1) ? g : a), cf.ranges[0]!);
    const evCtx: CfEvalCtx | undefined = ev ? { ...ev, r0: origin.r1, c0: origin.c1 } : undefined;
    // Gather the range's existing cells and their numeric/text values (for aggregate rules).
    const cells: { r: number; c: number; text: string; num: number | null }[] = [];
    for (const cell of sheet.cells.values())
      if (inRange(cell.row, cell.col)) { const text = cellDisplay(cell); cells.push({ r: cell.row, c: cell.col, text, num: numOf(text) }); }
    const nums = cells.map((c) => c.num).filter((n): n is number => n != null);
    const sorted = [...nums].sort((a, b) => a - b);
    const min = sorted[0] ?? 0;
    const max = sorted[sorted.length - 1] ?? 0;
    const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    const counts = new Map<string, number>();
    for (const c of cells) counts.set(c.text, (counts.get(c.text) ?? 0) + 1);

    for (const cell of cells) {
      const vis: CfVisual = {};
      let done = false;
      for (const rule of cf.rules) {
        if (done) break;
        const matched = matchRule(rule, cell, { min, max, avg, sorted, counts }, evCtx);
        if (rule.type === "colorScale" && rule.colorScale && cell.num != null) {
          out.set(key(cell.r, cell.c), { bg: colorScaleAt(cell.num, rule.colorScale, min, max, sorted) });
          done = true;
          break;
        }
        if (rule.type === "dataBar" && rule.dataBar && cell.num != null) {
          const lo = stopValue(rule.dataBar.min, min, max, sorted);
          const hi = stopValue(rule.dataBar.max, min, max, sorted);
          const pct = hi > lo ? Math.max(0, Math.min(1, (cell.num - lo) / (hi - lo))) : 0;
          out.set(key(cell.r, cell.c), { bar: { pct, color: rule.dataBar.color } });
          done = true;
          break;
        }
        if (rule.type === "iconSet" && rule.iconSet && cell.num != null) {
          const n = rule.iconSet.cfvo.length;
          const idx = iconIndex(cell.num, rule.iconSet, min, max, sorted);
          out.set(key(cell.r, cell.c), { icon: cfIconSvg(rule.iconSet.set, idx, n) });
          done = true;
          break;
        }
        if (matched && rule.dxf) {
          if (rule.dxf.bg) vis.bg = rule.dxf.bg;
          if (rule.dxf.color) vis.color = rule.dxf.color;
          if (rule.dxf.bold) vis.bold = true;
          if (rule.dxf.italic) vis.italic = true;
          if (rule.stopIfTrue || true) { done = true; } // first matching dxf wins (common case)
        }
      }
      if (vis.bg || vis.color || vis.bold || vis.italic) out.set(key(cell.r, cell.c), vis);
    }
  }
  return out;
}

function matchRule(rule: CfRule, cell: { r: number; c: number; text: string; num: number | null }, agg: { min: number; max: number; avg: number; sorted: number[]; counts: Map<string, number> }, ev?: CfEvalCtx): boolean {
  const v = cell.num;
  // Operands are literals in the common case; a cell-ref / formula operand ("$A$1", "AVERAGE(...)")
  // is evaluated relative to the range origin through the workbook's formula engine when available.
  const evalNum = (raw: string): number | null => {
    const lit = numOf(raw);
    if (lit != null) return lit;
    if (!ev) return null;
    const res = ev.evaluator.at(raw, ev.r0, ev.c0, cell.r, cell.c, ev.sheetName);
    return typeof res === "number" ? res : numOf(String(res));
  };
  const opnd = (i: number): number | null => { const f = rule.formulas?.[i]; return f != null ? evalNum(f) : null; };
  switch (rule.type) {
    case "cellIs": {
      if (v == null) return false;
      const a = opnd(0);
      const b = opnd(1);
      switch (rule.operator) {
        case "greaterThan": return a != null && v > a;
        case "greaterThanOrEqual": return a != null && v >= a;
        case "lessThan": return a != null && v < a;
        case "lessThanOrEqual": return a != null && v <= a;
        case "equal": return a != null && v === a;
        case "notEqual": return a != null && v !== a;
        case "between": return a != null && b != null && v >= a && v <= b;
        case "notBetween": return a != null && b != null && (v < a || v > b);
        default: return false;
      }
    }
    case "containsText": return !!rule.text && cell.text.includes(rule.text);
    case "notContainsText": return !!rule.text && !cell.text.includes(rule.text);
    case "beginsWith": return !!rule.text && cell.text.startsWith(rule.text);
    case "endsWith": return !!rule.text && cell.text.endsWith(rule.text);
    case "duplicateValues": return cell.text !== "" && (agg.counts.get(cell.text) ?? 0) > 1;
    case "uniqueValues": return cell.text !== "" && (agg.counts.get(cell.text) ?? 0) === 1;
    case "aboveAverage":
      if (v == null) return false;
      return rule.aboveAverage !== false ? (rule.equalAverage ? v >= agg.avg : v > agg.avg) : (rule.equalAverage ? v <= agg.avg : v < agg.avg);
    case "top10": {
      if (v == null || !agg.sorted.length) return false;
      const n = rule.percent ? Math.ceil((rule.rank ?? 10) / 100 * agg.sorted.length) : (rule.rank ?? 10);
      const cut = rule.bottom ? agg.sorted[Math.min(n, agg.sorted.length) - 1] : agg.sorted[Math.max(0, agg.sorted.length - n)];
      return rule.bottom ? v <= cut : v >= cut;
    }
    case "expression": {
      if (!ev || !rule.formulas?.[0]) return false;
      const res = ev.evaluator.at(rule.formulas[0], ev.r0, ev.c0, cell.r, cell.c, ev.sheetName);
      return res === true || (typeof res === "number" && res !== 0) || res === "TRUE";
    }
    default: return false; // iconSet / timePeriod: not evaluated
  }
}

function colorScaleAt(v: number, cs: { cfvo: { type: string; val?: number }[]; colors: string[] }, min: number, max: number, sorted: number[]): string {
  const stops = cs.cfvo.map((c, i) => ({ pos: stopValue(c, min, max, sorted), rgb: toRgb(cs.colors[i] ?? "#ffffff") }));
  if (v <= stops[0].pos) return `#${stops[0].rgb.map((x) => x.toString(16).padStart(2, "0")).join("")}`;
  const last = stops[stops.length - 1];
  if (v >= last.pos) return `#${last.rgb.map((x) => x.toString(16).padStart(2, "0")).join("")}`;
  for (let i = 1; i < stops.length; i++) {
    if (v <= stops[i].pos) {
      const lo = stops[i - 1];
      const hi = stops[i];
      const t = hi.pos > lo.pos ? (v - lo.pos) / (hi.pos - lo.pos) : 0;
      const rgb = [0, 1, 2].map((k) => lerp(lo.rgb[k], hi.rgb[k], t));
      return `#${rgb.map((x) => x.toString(16).padStart(2, "0")).join("")}`;
    }
  }
  return "#ffffff";
}
