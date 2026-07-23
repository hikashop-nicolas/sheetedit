// A dependency-free Chart.js date adapter (native Date), so date axes can use a real "time" scale
// with ticks snapped to day/month/year boundaries instead of a plain linear scale. Implements the
// small _adapters._date interface Chart.js calls; timestamps are epoch milliseconds.

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTHS_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const pad = (n: number, w = 2): string => String(Math.abs(n)).padStart(w, "0");

/** Format a Date with the token subset Chart.js's default formats use. */
function fmt(d: Date, format: string): string {
  const y = d.getFullYear(), M = d.getMonth(), D = d.getDate(), h = d.getHours(), m = d.getMinutes(), s = d.getSeconds(), ms = d.getMilliseconds();
  const h12 = h % 12 || 12, ap = h < 12 ? "AM" : "PM", q = Math.floor(M / 3) + 1;
  const tokens: Record<string, string> = {
    yyyy: String(y), yy: pad(y % 100), MMMM: MONTHS_FULL[M]!, MMM: MONTHS[M]!, MM: pad(M + 1), SSS: pad(ms, 3),
    dd: pad(D), HH: pad(h), hh: pad(h12), mm: pad(m), ss: pad(s), qqq: `Q${q}`, ha: `${h12}${ap.toLowerCase()}`,
    aaa: ap, a: ap, M: String(M + 1), d: String(D), H: String(h), h: String(h12), m: String(m), s: String(s), q: String(q),
  };
  // Longest tokens first so e.g. MMM wins over MM/M.
  return format.replace(/yyyy|yy|MMMM|MMM|MM|SSS|dd|HH|hh|mm|ss|qqq|ha|aaa|M|d|H|h|m|s|a|q/g, (t) => tokens[t] ?? t);
}

const UNIT_MS: Record<string, number> = { millisecond: 1, second: 1e3, minute: 6e4, hour: 36e5, day: 864e5, week: 6048e5 };

function add(time: number, amount: number, unit: string): number {
  const x = new Date(time);
  if (unit === "month") { x.setMonth(x.getMonth() + amount); return x.getTime(); }
  if (unit === "quarter") { x.setMonth(x.getMonth() + amount * 3); return x.getTime(); }
  if (unit === "year") { x.setFullYear(x.getFullYear() + amount); return x.getTime(); }
  return time + amount * (UNIT_MS[unit] ?? 1);
}

function startOf(time: number, unit: string, weekday?: number): number {
  const x = new Date(time);
  switch (unit) {
    case "second": x.setMilliseconds(0); break;
    case "minute": x.setSeconds(0, 0); break;
    case "hour": x.setMinutes(0, 0, 0); break;
    case "day": x.setHours(0, 0, 0, 0); break;
    case "week": { x.setHours(0, 0, 0, 0); const diff = (x.getDay() - (weekday ?? 0) + 7) % 7; x.setDate(x.getDate() - diff); break; }
    case "isoWeek": { x.setHours(0, 0, 0, 0); const diff = (x.getDay() + 6) % 7; x.setDate(x.getDate() - diff); break; }
    case "month": x.setDate(1); x.setHours(0, 0, 0, 0); break;
    case "quarter": x.setMonth(Math.floor(x.getMonth() / 3) * 3, 1); x.setHours(0, 0, 0, 0); break;
    case "year": x.setMonth(0, 1); x.setHours(0, 0, 0, 0); break;
  }
  return x.getTime();
}

function diff(a: number, b: number, unit: string): number {
  const da = new Date(a), db = new Date(b);
  if (unit === "month") return (da.getFullYear() - db.getFullYear()) * 12 + (da.getMonth() - db.getMonth());
  if (unit === "quarter") return diff(a, b, "month") / 3;
  if (unit === "year") return da.getFullYear() - db.getFullYear();
  return (a - b) / (UNIT_MS[unit] ?? 1);
}

const DATE_ADAPTER = {
  _id: "sheetedit-native",
  formats: () => ({
    datetime: "MMM d, yyyy, h:mm:ss a", millisecond: "h:mm:ss.SSS a", second: "h:mm:ss a", minute: "h:mm a",
    hour: "ha", day: "MMM d", week: "MMM d", month: "MMM yyyy", quarter: "qqq - yyyy", year: "yyyy",
  }),
  parse(value: unknown): number | null {
    if (value == null) return null;
    if (value instanceof Date) return value.getTime();
    if (typeof value === "number") return value;
    const t = Date.parse(String(value));
    return isNaN(t) ? null : t;
  },
  format(time: number, format: string): string { return fmt(new Date(time), format); },
  add(time: number, amount: number, unit: string): number { return add(time, amount, unit); },
  diff(max: number, min: number, unit: string): number { return diff(max, min, unit); },
  startOf(time: number, unit: string, weekday?: number): number { return startOf(time, unit, weekday); },
  endOf(time: number, unit: string): number { return add(startOf(time, unit), 1, unit) - 1; },
};

/** Install the native date adapter onto Chart.js's _adapters._date singleton (idempotent). */
export function registerDateAdapter(adapters: { _date?: { override(a: unknown): void } } | undefined): void {
  adapters?._date?.override(DATE_ADAPTER);
}
