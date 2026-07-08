import type { Cell } from "./model";
import { isNumeric, numToStr, shiftFormula } from "./model";

// Fill-handle semantics, as pure functions (the editor applies the results).
// Follows Excel where it is unsurprising:
// - formulas copy with relative references shifted;
// - a run of 2+ numbers along the fill axis extends as a linear series;
// - a single number copies as-is;
// - a single text value with a trailing integer increments it ("item1" -> "item2");
// - anything else repeats cyclically.

export interface FillSource {
  value: string;
  formula?: string;
  kind: Cell["kind"];
}

/** Linear series step when every source cell along the axis is numeric (2+ cells). */
export function seriesStep(source: FillSource[]): number | null {
  if (source.length < 2) return null;
  const nums = source.map((s) => (s.formula == null && isNumeric(s.value) ? Number(s.value) : null));
  if (nums.some((v) => v == null)) return null;
  const step = nums[1]! - nums[0]!;
  for (let i = 2; i < nums.length; i++) if (Math.abs(nums[i]! - nums[i - 1]! - step) > 1e-9) return null;
  return step;
}

const TRAILING_INT = /^(.*?)(\d+)$/;

/**
 * Raw inputs for the fill targets. `source` is the run of source cells along
 * the fill axis, `count` how many cells to fill beyond it, `dir` +1 for
 * down/right and -1 for up/left, `axis` which coordinate moves (for formula
 * reference shifting).
 */
export function computeFill(source: FillSource[], count: number, dir: 1 | -1, axis: "row" | "col"): string[] {
  const out: string[] = [];
  if (!source.length || count <= 0) return out;
  const len = source.length;
  const step = seriesStep(source);
  const trailing = len === 1 && source[0]!.formula == null && source[0]!.kind === "s" ? TRAILING_INT.exec(source[0]!.value) : null;
  for (let n = 0; n < count; n++) {
    const r = n % len;
    // Filling backwards walks the pattern in reverse, like Excel.
    const src = source[dir === 1 ? r : len - 1 - r]!;
    // The target sits `distance` cells (signed) from the source cell it copies.
    const distance = dir * len * (Math.floor(n / len) + 1);
    if (src.formula != null) {
      out.push("=" + shiftFormula(src.formula, axis === "row" ? distance : 0, axis === "col" ? distance : 0));
    } else if (step != null) {
      out.push(numToStr(Number(src.value) + step * distance));
    } else if (trailing) {
      out.push(trailing[1]! + String(Number(trailing[2]) + distance));
    } else {
      out.push(src.value);
    }
  }
  return out;
}
