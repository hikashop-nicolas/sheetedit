// ---------------------------------------------------------------------------
// VBA values and coercion (Stage 2 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// VBA's Variant holds a small set of things, and its coercion rules are specific enough that
// getting them wrong produces answers that look plausible. The ones that matter here:
//   Empty  - an unassigned variable. 0 in arithmetic, "" in concatenation, and equal to both.
//   Null   - propagates: any arithmetic involving Null is Null, which is NOT the same as Empty.
//   Nothing- an object reference that points at nothing.
//   Boolean- True is -1, not 1, once it reaches arithmetic.
// Concatenation with & always makes a string; + adds when both sides are numeric and concatenates
// when both are strings, which is why & exists.

/** An unassigned variable. */
export const EMPTY: unique symbol = Symbol.for("vba.Empty") as never;
/** SQL-style Null, which propagates through operators. */
export const NULL: unique symbol = Symbol.for("vba.Null") as never;
/** An object reference pointing at nothing. */
export const NOTHING: unique symbol = Symbol.for("vba.Nothing") as never;

export type VbaScalar = number | string | boolean | typeof EMPTY | typeof NULL | typeof NOTHING;
export type VbaValue = VbaScalar | VbaArray | VbaObject;

/** A VBA array, which knows its own bounds because they need not start at zero. */
export class VbaArray {
  constructor(readonly lower: number[], readonly upper: number[], readonly data: Map<string, VbaValue> = new Map()) {}
  static ofDims(dims: { lower: number; upper: number }[]): VbaArray {
    return new VbaArray(dims.map((d) => d.lower), dims.map((d) => d.upper));
  }
  private key(idx: number[]): string { return idx.join(","); }
  get(idx: number[]): VbaValue {
    return this.data.get(this.key(idx)) ?? EMPTY;
  }
  set(idx: number[], v: VbaValue): void { this.data.set(this.key(idx), v); }
  inBounds(idx: number[]): boolean {
    return idx.length === this.lower.length && idx.every((n, i) => n >= this.lower[i]! && n <= this.upper[i]!);
  }
  /** Every element in order, for For Each. */
  values(): VbaValue[] {
    const out: VbaValue[] = [];
    const walk = (dim: number, idx: number[]): void => {
      if (dim === this.lower.length) { out.push(this.get(idx)); return; }
      for (let i = this.lower[dim]!; i <= this.upper[dim]!; i++) walk(dim + 1, [...idx, i]);
    };
    walk(0, []);
    return out;
  }
}

/**
 * Anything with members. Stage 3's Excel objects implement this; the interpreter only ever talks to
 * this interface, so the language and the object model stay separable.
 */
export interface VbaObject {
  readonly typeName: string;
  /**
   * Read a property or call a method. `args` is empty for a plain property read. `argNames` holds
   * the names of any `Name:=value` arguments, positionally, since real macros use them constantly.
   */
  get(name: string, args: VbaValue[], argNames?: (string | undefined)[]): VbaValue;
  /** Assign to a property. Absent means the object is read-only. */
  set?(name: string, args: VbaValue[], value: VbaValue, argNames?: (string | undefined)[]): void;
  /** The members For Each walks. Absent means the object is not a collection. */
  enumerate?(): VbaValue[];
  /** The value an object stands for when used where a value is wanted (Range's is its Value). */
  defaultValue?(): VbaValue;
}

export const isObject = (v: VbaValue): v is VbaObject =>
  typeof v === "object" && v !== null && !(v instanceof VbaArray) && "typeName" in v;

/** A VBA runtime error, carrying the number Err.Number would report. */
export class VbaError extends Error {
  constructor(message: string, readonly number = 5, public line?: number) {
    super(message);
    this.name = "VbaError";
  }
}

/** The value an expression yields where a scalar is wanted: an object gives its default member. */
export function scalarOf(v: VbaValue): VbaScalar | VbaArray {
  if (isObject(v)) {
    if (!v.defaultValue) throw new VbaError(`${v.typeName} has no default value`, 438);
    const inner = v.defaultValue();
    return isObject(inner) ? scalarOf(inner) : inner;
  }
  return v;
}

/** Whether a value participates in Null propagation. */
export const isNull = (v: VbaValue): boolean => v === NULL;

/** VBA's numeric coercion. True is -1. Empty is 0. A non-numeric string is a type mismatch. */
export function toNumber(v: VbaValue): number {
  const s = scalarOf(v);
  if (s === EMPTY) return 0;
  if (s === NULL) throw new VbaError("invalid use of Null", 94);
  if (s === NOTHING) throw new VbaError("object variable not set", 91);
  if (typeof s === "number") return s;
  if (typeof s === "boolean") return s ? -1 : 0; // True is -1 in VBA, not 1
  if (s instanceof VbaArray) throw new VbaError("type mismatch", 13);
  const t = s.trim();
  if (t === "") return 0;
  const n = Number(t);
  if (!Number.isFinite(n)) throw new VbaError("type mismatch", 13);
  return n;
}

/** VBA's string coercion. Empty is "", and a boolean spells out True/False. */
export function toStr(v: VbaValue): string {
  const s = scalarOf(v);
  if (s === EMPTY) return "";
  if (s === NULL) throw new VbaError("invalid use of Null", 94);
  if (s === NOTHING) throw new VbaError("object variable not set", 91);
  if (typeof s === "string") return s;
  if (typeof s === "boolean") return s ? "True" : "False";
  if (s instanceof VbaArray) throw new VbaError("type mismatch", 13);
  return numToVbaString(s);
}

/** How VBA renders a number: no exponent for ordinary magnitudes, and no trailing ".0". */
export function numToVbaString(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  // VBA prints up to 15 significant digits and drops what is not needed.
  const s = Number(n.toPrecision(15)).toString();
  return s;
}

/** VBA's truthiness: zero is false, everything else true. Null in a condition is an error. */
export function toBool(v: VbaValue): boolean {
  const s = scalarOf(v);
  if (s === EMPTY) return false;
  if (s === NULL) throw new VbaError("invalid use of Null", 94);
  if (typeof s === "boolean") return s;
  return toNumber(s) !== 0;
}

/** Whether a value should be treated as a number rather than text in a loose comparison. */
const numericish = (v: VbaScalar | VbaArray): boolean =>
  typeof v === "number" || typeof v === "boolean" || v === EMPTY;

/**
 * VBA's comparison, returning -1, 0 or 1. Two numerics compare numerically; anything involving a
 * string compares as text, which is why "10" < "9" is true in VBA as it is in Excel.
 */
export function compare(a: VbaValue, b: VbaValue): number {
  const x = scalarOf(a);
  const y = scalarOf(b);
  if (x === NULL || y === NULL) throw new VbaError("invalid use of Null", 94);
  // Empty compares equal to both 0 and "", so it takes the other side's kind.
  if (x === EMPTY && typeof y === "string") return "" < y ? -1 : "" > y ? 1 : 0;
  if (y === EMPTY && typeof x === "string") return x < "" ? -1 : x > "" ? 1 : 0;
  if (numericish(x) && numericish(y)) {
    const nx = toNumber(x), ny = toNumber(y);
    return nx < ny ? -1 : nx > ny ? 1 : 0;
  }
  if (typeof x === "string" || typeof y === "string") {
    const sx = toStr(x), sy = toStr(y);
    return sx < sy ? -1 : sx > sy ? 1 : 0;
  }
  const nx = toNumber(x), ny = toNumber(y);
  return nx < ny ? -1 : nx > ny ? 1 : 0;
}

/** VBA's Like, whose wildcards are ? * # and [charlist]. */
export function like(text: string, pattern: string): boolean {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i]!;
    if (c === "?") re += ".";
    else if (c === "*") re += "[\\s\\S]*";
    else if (c === "#") re += "\\d";
    else if (c === "[") {
      const end = pattern.indexOf("]", i + 1);
      if (end === -1) { re += "\\["; continue; }
      let set = pattern.slice(i + 1, end);
      // VBA writes negation as [!abc] where a regex wants [^abc].
      if (set.startsWith("!")) set = `^${set.slice(1)}`;
      re += `[${set}]`;
      i = end;
    } else re += c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`).test(text);
}

/** The name VBA's TypeName would give a value. */
export function typeName(v: VbaValue): string {
  if (v === EMPTY) return "Empty";
  if (v === NULL) return "Null";
  if (v === NOTHING) return "Nothing";
  if (v instanceof VbaArray) return "Variant()";
  if (isObject(v)) return v.typeName;
  if (typeof v === "boolean") return "Boolean";
  if (typeof v === "string") return "String";
  return Number.isInteger(v) ? "Long" : "Double";
}
