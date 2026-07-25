import type { DataValidation } from "./model";

// Is a cell's value allowed by a data-validation rule? Used to flag invalid cells in the grid.
// list / textLength compare the displayed text; the numeric + date/time rules compare the cell's
// raw value (a date cell stores its serial, which the operands are also given as). A blank cell is
// treated as valid here (the allow-blank flag governs empties on entry, not display). Custom
// (formula) rules are not evaluated live and count as valid.

const num = (s: string | undefined): number | null => {
  if (s == null || s === "") return null;
  const n = Number(String(s).replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
};

function compare(v: number, op: DataValidation["operator"], a: number | null, b: number | null): boolean {
  switch (op ?? "between") {
    case "between": return a != null && b != null && v >= a && v <= b;
    case "notBetween": return a != null && b != null && (v < a || v > b);
    case "equal": return a != null && v === a;
    case "notEqual": return a != null && v !== a;
    case "greaterThan": return a != null && v > a;
    case "lessThan": return a != null && v < a;
    case "greaterThanOrEqual": return a != null && v >= a;
    case "lessThanOrEqual": return a != null && v <= a;
    default: return true;
  }
}

/** True when `display`/`raw` satisfy the rule. `allowed` is the resolved dropdown list (for `list`). */
export function validateCell(dv: DataValidation, raw: string, display: string, allowed: string[]): boolean {
  const type = dv.type ?? (dv.values || dv.rangeRef ? "list" : undefined);
  if (!type || display === "") return true;
  if (type === "list") return allowed.length === 0 || allowed.includes(display);
  if (type === "custom") return true; // not evaluated live
  const v = type === "textLength" ? display.length : num(raw);
  if (v == null) return false; // a numeric / date rule needs a number
  if (type === "whole" && !Number.isInteger(v)) return false;
  return compare(v, dv.operator, num(dv.formula1), num(dv.formula2));
}
