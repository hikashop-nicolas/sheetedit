import {
  EMPTY, isObject, NULL, numToVbaString, scalarOf, toBool, toNumber, toStr, typeName, VbaArray,
  VbaError, type VbaValue,
} from "./vba-value";

// VBA's built-in function library (Stage 2 of _plans/VBA_PLAN.md). The set real macros actually
// reach for. Anything not here is refused by name at the call site rather than silently returning
// Empty, which is the plan's rule.
//
// Nothing here touches files, the network or the shell: those surfaces do not exist in a browser,
// so a macro that needs them stops with a message saying so.

export interface BuiltinCtx {
  /** MsgBox and Debug.Print go here; a browser must not block on a dialog. */
  output: (text: string) => void;
  lastError: () => VbaError | null;
}

type Builtin = (args: VbaValue[], ctx: BuiltinCtx) => VbaValue;

const arg = (args: VbaValue[], i: number): VbaValue => args[i] ?? EMPTY;
const has = (args: VbaValue[], i: number): boolean => args[i] !== undefined && args[i] !== EMPTY;
const str = (args: VbaValue[], i: number): string => toStr(arg(args, i));
const num = (args: VbaValue[], i: number): number => toNumber(arg(args, i));

/** VBA's Int floors; Fix truncates toward zero. They differ only for negatives. */
const vbaInt = (n: number): number => Math.floor(n);
const vbaFix = (n: number): number => Math.trunc(n);

/** Banker's rounding, which is what VBA's Round does and a common source of surprise. */
function bankersRound(n: number, digits = 0): number {
  const f = Math.pow(10, digits);
  const x = n * f;
  const r = Math.round(x);
  // Exactly .5 rounds to the even neighbour.
  const out = Math.abs(x % 1) === 0.5 && r % 2 !== 0 ? r - Math.sign(x) : r;
  return out / f;
}

/** Excel serial (days since 1899-12-30) from a JS date. */
const toSerial = (d: Date): number => d.getTime() / 86400000 + 25569 - d.getTimezoneOffset() / 1440;

export const builtins: Record<string, Builtin> = {
  // --- strings ---------------------------------------------------------------
  LEN: (a) => (arg(a, 0) === NULL ? NULL : str(a, 0).length),
  LEFT: (a) => str(a, 0).slice(0, Math.max(0, num(a, 1))),
  RIGHT: (a) => {
    const n = Math.max(0, num(a, 1));
    return n === 0 ? "" : str(a, 0).slice(-n);
  },
  MID: (a) => {
    // VBA's Mid is 1-based, and a missing length means "to the end".
    const s = str(a, 0);
    const start = Math.max(1, num(a, 1));
    return has(a, 2) ? s.substr(start - 1, Math.max(0, num(a, 2))) : s.slice(start - 1);
  },
  TRIM: (a) => str(a, 0).replace(/^ +| +$/g, ""),
  LTRIM: (a) => str(a, 0).replace(/^ +/, ""),
  RTRIM: (a) => str(a, 0).replace(/ +$/, ""),
  UCASE: (a) => str(a, 0).toUpperCase(),
  LCASE: (a) => str(a, 0).toLowerCase(),
  INSTR: (a) => {
    // InStr(start, haystack, needle) or InStr(haystack, needle): the arity decides.
    const numeric = typeof scalarOf(arg(a, 0)) === "number";
    const start = numeric ? Math.max(1, num(a, 0)) : 1;
    const hay = numeric ? str(a, 1) : str(a, 0);
    const needle = numeric ? str(a, 2) : str(a, 1);
    return hay.indexOf(needle, start - 1) + 1; // 0 means not found, as in VBA
  },
  INSTRREV: (a) => str(a, 0).lastIndexOf(str(a, 1)) + 1,
  REPLACE: (a) => str(a, 0).split(str(a, 1)).join(str(a, 2)),
  SPLIT: (a) => {
    const parts = str(a, 0).split(has(a, 1) ? str(a, 1) : " ");
    const arr = VbaArray.ofDims([{ lower: 0, upper: parts.length - 1 }]);
    parts.forEach((p, i) => arr.set([i], p));
    return arr;
  },
  JOIN: (a) => {
    const arr = arg(a, 0);
    if (!(arr instanceof VbaArray)) throw new VbaError("Join needs an array", 13);
    return arr.values().map(toStr).join(has(a, 1) ? str(a, 1) : " ");
  },
  SPACE: (a) => " ".repeat(Math.max(0, num(a, 0))),
  STRING: (a) => {
    const c = str(a, 1);
    return (c[0] ?? " ").repeat(Math.max(0, num(a, 0)));
  },
  CHR: (a) => String.fromCharCode(num(a, 0)),
  ASC: (a) => str(a, 0).charCodeAt(0) || 0,
  STRCOMP: (a) => {
    const x = str(a, 0), y = str(a, 1);
    return x < y ? -1 : x > y ? 1 : 0;
  },
  STRREVERSE: (a) => [...str(a, 0)].reverse().join(""),

  // --- numbers ---------------------------------------------------------------
  ABS: (a) => Math.abs(num(a, 0)),
  INT: (a) => vbaInt(num(a, 0)),
  FIX: (a) => vbaFix(num(a, 0)),
  ROUND: (a) => bankersRound(num(a, 0), has(a, 1) ? num(a, 1) : 0),
  SQR: (a) => {
    const n = num(a, 0);
    if (n < 0) throw new VbaError("Sqr of a negative number", 5);
    return Math.sqrt(n);
  },
  SGN: (a) => Math.sign(num(a, 0)),
  EXP: (a) => Math.exp(num(a, 0)),
  LOG: (a) => Math.log(num(a, 0)),
  ATN: (a) => Math.atan(num(a, 0)),
  SIN: (a) => Math.sin(num(a, 0)),
  COS: (a) => Math.cos(num(a, 0)),
  TAN: (a) => Math.tan(num(a, 0)),

  // --- conversion ------------------------------------------------------------
  CSTR: (a) => toStr(arg(a, 0)),
  CINT: (a) => bankersRound(num(a, 0)),
  CLNG: (a) => bankersRound(num(a, 0)),
  CDBL: (a) => num(a, 0),
  CSNG: (a) => num(a, 0),
  CBOOL: (a) => toBool(arg(a, 0)),
  CDATE: (a) => {
    const ms = Date.parse(str(a, 0));
    if (Number.isNaN(ms)) throw new VbaError("that is not a date", 13);
    return ms / 86400000 + 25569;
  },
  VAL: (a) => {
    // Val reads the longest numeric prefix and gives 0 when there is none.
    const m = /^\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/.exec(str(a, 0));
    return m ? Number(m[0]) : 0;
  },
  STR: (a) => {
    // Str puts a leading space where the sign would be, for positives.
    const n = num(a, 0);
    return n >= 0 ? ` ${numToVbaString(n)}` : numToVbaString(n);
  },
  HEX: (a) => {
    const n = Math.trunc(num(a, 0));
    return (n < 0 ? (n >>> 0) : n).toString(16).toUpperCase();
  },
  OCT: (a) => Math.trunc(num(a, 0)).toString(8),

  // --- tests -----------------------------------------------------------------
  ISEMPTY: (a) => arg(a, 0) === EMPTY,
  ISNULL: (a) => arg(a, 0) === NULL,
  ISNUMERIC: (a) => {
    const v = arg(a, 0);
    if (typeof v === "number" || typeof v === "boolean") return true;
    if (typeof v !== "string") return false;
    const t = v.trim();
    return t !== "" && Number.isFinite(Number(t));
  },
  ISARRAY: (a) => arg(a, 0) instanceof VbaArray,
  ISOBJECT: (a) => isObject(arg(a, 0)),
  ISDATE: (a) => !Number.isNaN(Date.parse(str(a, 0))),
  TYPENAME: (a) => typeName(arg(a, 0)),
  VARTYPE: (a) => {
    const v = arg(a, 0);
    if (v === EMPTY) return 0;
    if (v === NULL) return 1;
    if (typeof v === "boolean") return 11;
    if (typeof v === "string") return 8;
    if (typeof v === "number") return Number.isInteger(v) ? 3 : 5;
    return 9;
  },

  // --- arrays ----------------------------------------------------------------
  ARRAY: (a) => {
    const arr = VbaArray.ofDims([{ lower: 0, upper: a.length - 1 }]);
    a.forEach((v, i) => arr.set([i], v));
    return arr;
  },
  UBOUND: (a) => {
    const arr = arg(a, 0);
    if (!(arr instanceof VbaArray)) throw new VbaError("UBound needs an array", 13);
    return arr.upper[has(a, 1) ? num(a, 1) - 1 : 0] ?? 0;
  },
  LBOUND: (a) => {
    const arr = arg(a, 0);
    if (!(arr instanceof VbaArray)) throw new VbaError("LBound needs an array", 13);
    return arr.lower[has(a, 1) ? num(a, 1) - 1 : 0] ?? 0;
  },

  // --- dates -----------------------------------------------------------------
  NOW: () => toSerial(new Date()),
  DATE: () => Math.floor(toSerial(new Date())),
  TIME: () => toSerial(new Date()) % 1,
  YEAR: (a) => new Date((num(a, 0) - 25569) * 86400000).getUTCFullYear(),
  MONTH: (a) => new Date((num(a, 0) - 25569) * 86400000).getUTCMonth() + 1,
  DAY: (a) => new Date((num(a, 0) - 25569) * 86400000).getUTCDate(),

  // --- interaction -----------------------------------------------------------
  MSGBOX: (a, ctx) => {
    // A browser must not block on a dialog, and a macro that stopped for one would never finish.
    // The text is collected and shown after the run instead.
    ctx.output(str(a, 0));
    return 1; // vbOK, so `If MsgBox(...) = vbOK` behaves
  },
  INPUTBOX: (): VbaValue => {
    throw new VbaError("InputBox needs someone to answer it, which a macro run cannot do", 5);
  },

  // --- constants VBA exposes as names ----------------------------------------
  VBCRLF: () => "\r\n",
  VBLF: () => "\n",
  VBCR: () => "\r",
  VBTAB: () => "\t",
  VBNULLSTRING: () => "",
  VBOK: () => 1,
  VBCANCEL: () => 2,
  VBYES: () => 6,
  VBNO: () => 7,
  VBYESNO: () => 4,
  VBINFORMATION: () => 64,
  VBEXCLAMATION: () => 48,
  VBQUESTION: () => 32,
  VBCRITICAL: () => 16,
};
