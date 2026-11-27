// Typing the text of an HTML cell the way a spreadsheet would: numbers, percents, currency and
// dates become values, while ids with leading zeros and long digit strings stay text.
import { parseDateInput, type DayOrder } from "../../core/dates";
import type { TValue } from "./model";

export interface Typed { v: TValue; fmt?: string }

const SPACES = /[\s   ]/g;
const DOT_NUM = /^(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?$/;
const COMMA_NUM = /^(\d{1,3}(?:[.\s   ]\d{3})+|\d+)(,\d+)?$/;
const CURRENCY = /^([$€£¥])\s?|\s?([$€£¥]|EUR|USD)$/;

/** Whether a table's numbers are written the European way ("1 234,56"), by majority vote. */
export function sniffCommaDecimal(texts: Iterable<string>): boolean {
  let comma = 0, dot = 0;
  for (const raw of texts) {
    const t = raw.trim().replace(CURRENCY, "").replace(/%$/, "").replace(/^[-+]/, "");
    if (!t || t.length > 24) continue;
    if (/,\d+$/.test(t) && COMMA_NUM.test(t) && !/^\d{1,3},\d{3}$/.test(t)) comma++;
    else if (/^\d{1,3}(?:[\s  ]\d{3})+$/.test(t)) comma++;
    else if (/\.\d+$/.test(t) && DOT_NUM.test(t) && !/^\d{1,3}\.\d{3}$/.test(t)) dot++;
  }
  return comma > dot;
}

const decimalsOf = (frac: string | undefined): number => (frac ? frac.length - 1 : 0);
const zeros = (n: number): string => (n ? "." + "0".repeat(Math.min(n, 10)) : "");

/** A number as text the way the page writes it ("1,234.50", "1 234,50", "12%"), or null. */
export function parseLooseNumber(raw: string, commaDecimal: boolean): Typed | null {
  let t = raw.trim();
  if (!t || t.length > 40) return null;
  let sign = 1;
  if (/^[-+−]/.test(t)) { if (t[0] !== "+") sign = -1; t = t.slice(1).trim(); }
  let cur = "";
  const cm = CURRENCY.exec(t);
  if (cm) { cur = cm[1] ?? cm[2]!; t = t.replace(CURRENCY, "").trim(); }
  let pct = false;
  if (/\s?%$/.test(t)) { pct = true; t = t.replace(/\s?%$/, ""); }
  if (cur && pct) return null;
  if (/^[-−]/.test(t) && sign === 1 && cur) { sign = -1; t = t.slice(1).trim(); }
  const sci = /^(\d+(?:\.\d+)?)[eE]([-+]?\d+)$/.exec(t);
  if (sci && !commaDecimal) return { v: { t: "n", v: sign * Number(t) }, fmt: "0.00E+00" };
  const m = (commaDecimal ? COMMA_NUM : DOT_NUM).exec(t) ?? (/^[.,]\d+$/.test(t) && (t[0] === ",") === commaDecimal ? [t, "0", t] : null);
  if (!m) return null;
  const intPart = m[1]!.replace(SPACES, "").replace(commaDecimal ? /\./g : /,/g, "");
  // Ids keep their leading zeros, and digit strings past 15 digits would lose precision.
  if (/^0\d/.test(intPart)) return null;
  if ((intPart + (m[2] ?? "")).replace(/\D/g, "").length > 15) return null;
  const frac = m[2] ? "." + m[2].slice(1) : "";
  let v = sign * Number(intPart + frac);
  if (!Number.isFinite(v)) return null;
  const grouped = m[1]!.length !== intPart.length;
  const dec = decimalsOf(m[2]);
  let fmt: string | undefined;
  if (pct) { v /= 100; fmt = "0" + zeros(dec) + "%"; }
  else if (cur) {
    const sym = cur === "EUR" ? "€" : cur === "USD" ? "$" : cur;
    const body = "#,##0" + zeros(dec || (grouped || cm ? 2 : 0));
    fmt = cm?.[1] ? `"${sym}"${body}` : `${body}\\ "${sym}"`;
  } else if (grouped) fmt = "#,##0" + zeros(dec);
  return { v: { t: "n", v }, fmt };
}

/** Type a cell's text: number, date, boolean or string. */
export function typeCellText(text: string, commaDecimal: boolean, dayOrder: DayOrder): Typed {
  const t = text.trim();
  if (!t) return { v: { t: "s", v: text } };
  const n = parseLooseNumber(t, commaDecimal);
  if (n) return n;
  if (/^\d/.test(t) && /[-/.:]/.test(t)) {
    const d = parseDateInput(t, dayOrder);
    if (d) return { v: { t: "n", v: d.serial }, fmt: d.fmt };
  }
  const up = t.toUpperCase();
  if (up === "TRUE" || up === "FALSE") return { v: { t: "b", v: up === "TRUE" } };
  return { v: { t: "s", v: text } };
}
