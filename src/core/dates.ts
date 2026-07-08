// ---------------------------------------------------------------------------
// Dates: Excel 1900-system serial conversions, typed-date input parsing, and
// number-format classification shared by the grid and both file adapters.
// ---------------------------------------------------------------------------

const DAY_MS = 86400000;
// Serial 0 is 1899-12-30 for serials >= 61; the fake 1900-02-29 (serial 60) shifts
// everything before 1900-03-01 by one day.
const EPOCH = Date.UTC(1899, 11, 30);

export function dateToSerial(y: number, mo: number, d: number, hh = 0, mi = 0, ss = 0): number {
  let days = Math.round((Date.UTC(y, mo - 1, d) - EPOCH) / DAY_MS);
  if (days < 61) days -= 1;
  return days + (hh * 3600 + mi * 60 + ss) / 86400;
}

export interface DateParts {
  y: number;
  mo: number;
  d: number;
  hh: number;
  mi: number;
  ss: number;
}

export function serialToParts(serial: number): DateParts | null {
  if (!Number.isFinite(serial) || serial < 0 || serial > 2958465) return null; // 9999-12-31
  let days = Math.floor(serial);
  let frac = serial - days;
  let secs = Math.round(frac * 86400);
  if (secs >= 86400) {
    secs = 0;
    days += 1;
  }
  if (days < 60) days += 1; // undo the fake-leap-day shift
  const dt = new Date(EPOCH + days * DAY_MS);
  return {
    y: dt.getUTCFullYear(),
    mo: dt.getUTCMonth() + 1,
    d: dt.getUTCDate(),
    hh: Math.floor(secs / 3600),
    mi: Math.floor((secs % 3600) / 60),
    ss: secs % 60,
  };
}

const p2 = (n: number): string => String(n).padStart(2, "0");

/** ISO 8601 for ODF office:date-value ("yyyy-mm-dd" or with "Thh:mm:ss"). */
export function serialToIso(serial: number): string | null {
  const p = serialToParts(serial);
  if (!p) return null;
  const date = `${p.y}-${p2(p.mo)}-${p2(p.d)}`;
  return p.hh || p.mi || p.ss ? `${date}T${p2(p.hh)}:${p2(p.mi)}:${p2(p.ss)}` : date;
}

/** Editable text for a date cell ("yyyy-mm-dd", plus " hh:mm[:ss]" when asked). */
export function serialToEditText(serial: number, withTime: boolean): string | null {
  const p = serialToParts(serial);
  if (!p) return null;
  const date = `${p.y}-${p2(p.mo)}-${p2(p.d)}`;
  if (!withTime && !(p.hh || p.mi || p.ss)) return date;
  return p.ss ? `${date} ${p2(p.hh)}:${p2(p.mi)}:${p2(p.ss)}` : `${date} ${p2(p.hh)}:${p2(p.mi)}`;
}

const validYmd = (y: number, mo: number, d: number): boolean => {
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1900 || y > 9999) return false;
  const dt = new Date(Date.UTC(y, mo - 1, d));
  return dt.getUTCMonth() === mo - 1 && dt.getUTCDate() === d;
};

/** Parse an ODF office:date-value / xlsx t="d" ISO string into a serial. */
export function isoToSerial(s: string): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?/.exec(s.trim());
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  if (!validYmd(y, mo, d)) return null;
  return dateToSerial(y, mo, d, Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0));
}

/** Parse an ODF office:time-value duration ("PT13H30M5S") into a day fraction. */
export function durationToSerial(s: string): number | null {
  const m = /^-?PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?$/.exec(s.trim());
  if (!m || (m[1] == null && m[2] == null && m[3] == null)) return null;
  const secs = Number(m[1] ?? 0) * 3600 + Number(m[2] ?? 0) * 60 + Number(m[3] ?? 0);
  return (s.startsWith("-") ? -secs : secs) / 86400;
}

export function serialToDuration(serial: number): string {
  const neg = serial < 0;
  const secs = Math.round(Math.abs(serial) * 86400);
  return `${neg ? "-" : ""}PT${Math.floor(secs / 3600)}H${Math.floor((secs % 3600) / 60)}M${secs % 60}S`;
}

export type DayOrder = "dmy" | "mdy";

export interface ParsedDateInput {
  serial: number;
  /** A number format matching what the user typed, applied when the cell has none. */
  fmt: string;
  /** "13:30" with no date part: a time-of-day fraction. */
  timeOnly?: boolean;
}

/**
 * Recognize a typed date: "2026-07-08", "2026/07/08", "8/7/2026", "08.07.2026",
 * each with an optional " hh:mm[:ss]" tail, or a bare "13:30[:05]" time. Ambiguous
 * d/m vs m/d numbers follow the locale's day order; an out-of-range component
 * disambiguates on its own.
 */
export function parseDateInput(raw: string, order: DayOrder): ParsedDateInput | null {
  const s = raw.trim();
  const timeOnly = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (timeOnly) {
    const hh = Number(timeOnly[1]), mi = Number(timeOnly[2]), ss = Number(timeOnly[3] ?? 0);
    if (hh > 23 || mi > 59 || ss > 59) return null;
    return { serial: (hh * 3600 + mi * 60 + ss) / 86400, fmt: timeOnly[3] != null ? "hh:mm:ss" : "hh:mm", timeOnly: true };
  }
  const time = "(?:[ T](\\d{1,2}):(\\d{2})(?::(\\d{2}))?)?$";
  let y: number, mo: number, d: number, fmt: string;
  let tm: (string | undefined)[];
  let m = new RegExp("^(\\d{4})([-/])(\\d{1,2})\\2(\\d{1,2})" + time).exec(s);
  if (m) {
    y = Number(m[1]);
    mo = Number(m[3]);
    d = Number(m[4]);
    fmt = "yyyy-mm-dd";
    tm = [m[5], m[6], m[7]];
  } else {
    m = new RegExp("^(\\d{1,2})([/.])(\\d{1,2})\\2(\\d{4})" + time).exec(s);
    if (!m) return null;
    const a = Number(m[1]), b = Number(m[3]);
    y = Number(m[4]);
    if (a > 12 && b <= 12) [d, mo] = [a, b];
    else if (b > 12 && a <= 12) [mo, d] = [a, b];
    else if (order === "dmy") [d, mo] = [a, b];
    else [mo, d] = [a, b];
    fmt = order === "dmy" ? "dd/mm/yyyy" : "mm/dd/yyyy";
    tm = [m[5], m[6], m[7]];
  }
  if (!validYmd(y, mo, d)) return null;
  const hh = Number(tm[0] ?? 0), mi = Number(tm[1] ?? 0), ss = Number(tm[2] ?? 0);
  if (hh > 23 || mi > 59 || ss > 59) return null;
  if (tm[0] != null) fmt += tm[2] != null ? " hh:mm:ss" : " hh:mm";
  return { serial: dateToSerial(y, mo, d, hh, mi, ss), fmt };
}

// Built-in xlsx numFmtIds that render as dates and/or times.
const DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
const TIME_IDS = new Set([18, 19, 20, 21, 22, 32, 33, 34, 35, 45, 46, 47]);

const bareCode = (fmt: string): string => fmt.replace(/"[^"]*"|\[[^\]]*\]|\\./g, "");

/** Does this number format (code or built-in id) render its value as a date/time? */
export function isDateFmt(fmt: string | number | undefined): boolean {
  if (fmt == null) return false;
  if (typeof fmt === "number") return DATE_IDS.has(fmt);
  return /[ymdhs]/i.test(bareCode(fmt));
}

export function hasTimeFmt(fmt: string | number | undefined): boolean {
  if (fmt == null) return false;
  if (typeof fmt === "number") return TIME_IDS.has(fmt);
  return /[hs]/i.test(bareCode(fmt));
}

/** Time-of-day with no date part ("hh:mm:ss"): no year/day tokens in the code. */
export function isTimeOnlyFmt(fmt: string | number | undefined): boolean {
  if (fmt == null) return false;
  if (typeof fmt === "number") return fmt >= 18 && fmt <= 21;
  return isDateFmt(fmt) && !/[yd]/i.test(bareCode(fmt));
}

export interface NumFmtPreset {
  /** i18n key for the menu label. */
  key: string;
  /** undefined = General. */
  fmt?: string | number;
  /** ODF currency code for currency presets. */
  currency?: string;
}

/** The number-format picker's preset list, currency and date shapes per locale. */
export function numFmtPresets(locale: string): NumFmtPreset[] {
  const cur =
    locale === "fr"
      ? { fmt: '#,##0.00 "€"', currency: "EUR" }
      : locale === "ja"
        ? { fmt: '"¥"#,##0', currency: "JPY" }
        : { fmt: '"$"#,##0.00', currency: "USD" };
  const date = locale === "fr" ? "dd/mm/yyyy" : locale === "ja" ? "yyyy/mm/dd" : "mm/dd/yyyy";
  return [
    { key: "fmtGeneral" },
    { key: "fmtNumber", fmt: "0.00" },
    { key: "fmtThousands", fmt: "#,##0.00" },
    { key: "fmtPercent", fmt: "0.00%" },
    { key: "fmtCurrency", ...cur },
    { key: "fmtDate", fmt: date },
    { key: "fmtDateTime", fmt: `${date} hh:mm` },
    { key: "fmtTime", fmt: "hh:mm:ss" },
  ];
}
