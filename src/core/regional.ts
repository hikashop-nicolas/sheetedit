// Regional conventions: how numbers, dates and formulas are typed and shown, as Excel does in
// each language. The workbook always keeps the invariant form ("." decimals, "," separators,
// English function names); this module only converts at the edges, for the user.

import { localeCode } from "./i18n";
import { FUNCTION_NAMES } from "./fn-names";
import { isDateFmt, serialToParts, type DayOrder } from "./dates";
import { formatNumber, type Cell } from "./model";

export interface Regional {
  code: string;
  decimal: "." | ",";
  /** Thousands separator as displayed. */
  group: string;
  /** Function argument separator in formulas. */
  argSep: "," | ";";
  /** Array constant separators: between columns, between rows. */
  arrayCol: string;
  arrayRow: string;
  dayOrder: DayOrder;
  /** Separator of an edited date ("dd/mm/yyyy"); undefined keeps ISO edit text. */
  dateSep?: "/" | ".";
  /** Format shown for the built-in short date (numFmtId 14); undefined keeps SSF's. */
  shortDate?: string;
  /** English function name -> local name (only the names that differ). */
  fn: Record<string, string>;
  bool?: [string, string];
  /** Invariant error literal -> local one (only those that differ). */
  errors: Record<string, string>;
  /** Month and day names: full then short, January..December and Sunday..Saturday. */
  months?: [string[], string[]];
  days?: [string[], string[]];
}

const INVARIANT: Regional = { code: "en", decimal: ".", group: ",", argSep: ",", arrayCol: ",", arrayRow: ";", dayOrder: "mdy", fn: {}, errors: {} };

const commaBase = { decimal: ",", argSep: ";", arrayCol: ".", arrayRow: ";", dayOrder: "dmy" } as const;

const REGIONS: Record<string, Regional> = {
  en: INVARIANT,
  ja: { ...INVARIANT, code: "ja" },
  zh: { ...INVARIANT, code: "zh" },
  fr: {
    ...commaBase, code: "fr", group: " ", dateSep: "/", shortDate: "dd/mm/yyyy",
    fn: FUNCTION_NAMES.fr!, bool: ["VRAI", "FAUX"],
    errors: { "#NULL!": "#NUL!", "#VALUE!": "#VALEUR!", "#NAME?": "#NOM?", "#NUM!": "#NOMBRE!" },
    months: [
      ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"],
      ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."],
    ],
    days: [
      ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"],
      ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."],
    ],
  },
  de: {
    ...commaBase, code: "de", group: ".", dateSep: ".", shortDate: "dd.mm.yyyy",
    fn: FUNCTION_NAMES.de!, bool: ["WAHR", "FALSCH"],
    errors: { "#VALUE!": "#WERT!", "#REF!": "#BEZUG!", "#NUM!": "#ZAHL!", "#N/A": "#NV" },
    months: [
      ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
      ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"],
    ],
    days: [
      ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"],
      ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"],
    ],
  },
  es: {
    ...commaBase, code: "es", group: ".", dateSep: "/", shortDate: "dd/mm/yyyy",
    fn: FUNCTION_NAMES.es!, bool: ["VERDADERO", "FALSO"],
    errors: { "#NULL!": "#¡NULO!", "#DIV/0!": "#¡DIV/0!", "#VALUE!": "#¡VALOR!", "#REF!": "#¡REF!", "#NAME?": "#¿NOMBRE?", "#NUM!": "#¡NUM!" },
    months: [
      ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
      ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"],
    ],
    days: [
      ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"],
      ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"],
    ],
  },
  it: {
    ...commaBase, code: "it", group: ".", dateSep: "/", shortDate: "dd/mm/yyyy",
    fn: FUNCTION_NAMES.it!, bool: ["VERO", "FALSO"],
    errors: { "#NULL!": "#NULLO!", "#VALUE!": "#VALORE!", "#REF!": "#RIF!", "#NAME?": "#NOME?", "#N/A": "#N/D" },
    months: [
      ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"],
      ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"],
    ],
    days: [
      ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"],
      ["dom", "lun", "mar", "mer", "gio", "ven", "sab"],
    ],
  },
  pl: {
    ...commaBase, code: "pl", group: " ", dateSep: ".", shortDate: "dd.mm.yyyy",
    fn: FUNCTION_NAMES.pl!, bool: ["PRAWDA", "FAŁSZ"],
    errors: { "#NULL!": "#ZERO!", "#DIV/0!": "#DZIEL/0!", "#VALUE!": "#ARG!", "#REF!": "#ADR!", "#NAME?": "#NAZWA?", "#NUM!": "#LICZBA!", "#N/A": "#N/D!" },
    months: [
      ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień"],
      ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru"],
    ],
    days: [
      ["niedziela", "poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota"],
      ["niedz.", "pon.", "wt.", "śr.", "czw.", "pt.", "sob."],
    ],
  },
  tr: {
    ...commaBase, code: "tr", group: ".", dateSep: ".", shortDate: "dd.mm.yyyy",
    fn: FUNCTION_NAMES.tr!, bool: ["DOĞRU", "YANLIŞ"],
    errors: { "#NULL!": "#BOŞ!", "#DIV/0!": "#SAYI/0!", "#VALUE!": "#DEĞER!", "#REF!": "#BAŞV!", "#NAME?": "#AD?", "#NUM!": "#SAYI!", "#N/A": "#YOK" },
    months: [
      ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"],
      ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"],
    ],
    days: [
      ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"],
      ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"],
    ],
  },
  // Korean Excel keeps English function names, separators, TRUE/FALSE and mmm/ddd names
  // (as Japanese does); only the built-in short date follows the region.
  ko: { ...INVARIANT, code: "ko", shortDate: "yyyy-mm-dd" },
  pt: { ...commaBase, code: "pt", group: ".", dateSep: "/", shortDate: "dd/mm/yyyy", fn: {}, errors: {} },
  ru: { ...commaBase, code: "ru", group: " ", dateSep: ".", shortDate: "dd.mm.yyyy", fn: {}, errors: {} },
};

let override: string | null = null;

/**
 * Force the regional conventions (numbers, dates, formula language) independently of the UI
 * language; null follows the UI locale again. "en" gives the invariant, English conventions.
 */
export function setRegionalLocale(code: string | null): void {
  override = code == null ? null : code.toLowerCase().split("-")[0]!;
}

/** The active regional conventions. */
export function regional(): Regional {
  return REGIONS[override ?? localeCode()] ?? INVARIANT;
}

/** True when nothing needs converting (English-style numbers and formulas). */
export const isInvariant = (r: Regional): boolean => r.decimal === "." && r.argSep === "," && !Object.keys(r.fn).length && !r.bool;

/** Swap the separators of an invariant (SSF) number text: "1,234.5" -> "1 234,5". */
export function localizeNumberText(s: string, r: Regional = regional()): string {
  if (r.decimal === ".") return s;
  return s.replace(/[.,]/g, (ch, i: number) => {
    const prev = s[i - 1] ?? "";
    const next = s[i + 1] ?? "";
    if (ch === "," && /\d/.test(prev) && /^\d{3}/.test(s.slice(i + 1))) return r.group;
    if (ch === "." && /\d/.test(next) && !/[A-Za-z]/.test(prev)) return r.decimal;
    return ch;
  });
}

const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/**
 * Read a number typed the local way ("1 234,5", "1.234,5") into its invariant text ("1234.5");
 * null when the text is not a local number. Space groups accept any space flavour.
 */
export function parseLocalNumber(text: string, r: Regional = regional()): string | null {
  if (r.decimal === ".") return null;
  const t = text.trim();
  const g = /\s/.test(r.group) ? "[\\u0020\\u00A0\\u202F]" : escapeRe(r.group);
  const d = escapeRe(r.decimal);
  const re = new RegExp(`^([-+]?)(\\d{1,3}(?:${g}\\d{3})+|\\d+)?(?:${d}(\\d+))?([eE][-+]?\\d+)?$`);
  const m = re.exec(t);
  if (!m || (m[2] == null && m[3] == null)) return null;
  const int = (m[2] ?? "0").replace(new RegExp(g, "g"), "");
  return `${m[1] === "-" ? "-" : ""}${int}${m[3] != null ? "." + m[3] : ""}${m[4] ?? ""}`;
}

/** An invariant number as the user edits it locally: "1234.5" -> "1234,5" (no groups). */
export function localNumberInput(value: string, r: Regional = regional()): string {
  if (r.decimal === "." || !/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(value)) return value;
  return value.replace(".", r.decimal);
}

const EN_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const EN_DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const nameMaps = new Map<string, { re: RegExp; map: Map<string, string> }>();

/** Replace SSF's English month and day names with the local ones. */
export function localizeDateNames(s: string, r: Regional = regional()): string {
  if (!r.months || !r.days) return s;
  let nm = nameMaps.get(r.code);
  if (!nm) {
    const map = new Map<string, string>();
    EN_MONTHS.forEach((m, i) => { map.set(m, r.months![0][i]!); map.set(m.slice(0, 3), r.months![1][i]!); });
    EN_DAYS.forEach((d, i) => { map.set(d, r.days![0][i]!); map.set(d.slice(0, 3), r.days![1][i]!); });
    const keys = [...map.keys()].sort((a, b) => b.length - a.length);
    nm = { re: new RegExp(`\\b(${keys.join("|")})\\b`, "g"), map };
    nameMaps.set(r.code, nm);
  }
  const { re, map } = nm;
  return s.replace(re, (w) => map.get(w) ?? w);
}

/** The format a cell shows with locally: the built-in short date follows the region. */
export function localFormat(fmt: string | number, r: Regional = regional()): string | number {
  if (!r.shortDate) return fmt;
  // SSF reads "dd.mm" as a decimal point: the dots must be literal.
  const d = r.shortDate.replace(/\./g, "\\.");
  if (fmt === 14) return d;
  if (fmt === 22) return `${d} hh:mm`;
  return fmt;
}

/** Does an invariant-style display text ("$1,234.50", "12.5%") carry this value? */
export function isInvariantText(text: string, value: string): boolean {
  const m = /\d+(?:\.(\d+))?(?:[eE][-+]?\d+)?/.exec(text.replace(/,(?=\d{3})/g, ""));
  if (!m) return false;
  const shown = Number(m[0]), v = Math.abs(Number(value));
  if (!Number.isFinite(shown) || !Number.isFinite(v)) return false;
  const tol = /[eE]/.test(m[0]) ? Math.abs(shown) * 1e-6 : 0.5 * 10 ** -(m[1]?.length ?? 0) + 1e-9;
  return Math.abs(shown - v) <= tol || (text.includes("%") && Math.abs(shown - v * 100) <= tol);
}

/** Nothing to convert on display: English-style numbers, names and dates. */
const plainDisplay = (r: Regional): boolean =>
  r.decimal === "." && !r.bool && !r.months && !r.shortDate && !Object.keys(r.errors).length;

const shown = new WeakMap<Cell, { k: string; out: string }>();

/** The text a cell shows in the user's region; the stored value and display stay invariant. */
export function localDisplay(cell: Cell | undefined, r: Regional = regional()): string {
  if (!cell) return "";
  const base = cell.display ?? cell.value;
  if (plainDisplay(r)) return base;
  if (cell.kind === "b") return r.bool && (cell.value === "TRUE" || cell.value === "FALSE") ? r.bool[cell.value === "TRUE" ? 0 : 1] : base;
  if (cell.kind === "e") return r.errors[cell.value] ?? base;
  if (cell.kind !== "n" || cell.displayRaw) return base;
  const k = `${r.code}\u0000${base}\u0000${String(cell.numFmt ?? "")}`;
  const hit = shown.get(cell);
  if (hit?.k === k) return hit.out;
  let out: string;
  if (cell.numFmt != null && isDateFmt(cell.numFmt)) {
    const fmt = localFormat(cell.numFmt, r);
    out = localizeDateNames(fmt !== cell.numFmt ? (formatNumber(fmt, cell.value) ?? base) : base, r);
  } else out = localizeNumberText(base, r);
  shown.set(cell, { k, out });
  return out;
}

const p2 = (n: number): string => String(n).padStart(2, "0");

/** Edit text of a date in the user's region ("25/11/2026 13:30"); null keeps the ISO text. */
export function localDateEditText(serial: number, withTime: boolean, r: Regional = regional()): string | null {
  if (!r.dateSep) return null;
  const p = serialToParts(serial);
  if (!p) return null;
  const s = r.dateSep;
  const date = r.dayOrder === "dmy" ? `${p2(p.d)}${s}${p2(p.mo)}${s}${p.y}` : `${p2(p.mo)}${s}${p2(p.d)}${s}${p.y}`;
  if (!withTime && !(p.hh || p.mi || p.ss)) return date;
  return `${date} ${p2(p.hh)}:${p2(p.mi)}${p.ss ? ":" + p2(p.ss) : ""}`;
}
