// ---------------------------------------------------------------------------
// IF
// ---------------------------------------------------------------------------
// fast-formula-parser lists IF in funsNeedContextAndNoDataRetrieve so it can leave the untaken
// branch unevaluated, but it never dereferences the CONDITION either: `IF(A1, x, y)` receives a
// reference object, and an object is always truthy, so it always took the true branch whatever the
// cell held. Silently wrong, in one of the most common formulas there is.
//
// The fix keeps IF on that list (so `IF(A1=0, "n/a", 1/A1)` still does not evaluate the division)
// and dereferences the condition here, with Excel's own coercion rules.

/** The subset of the parser context these need: its extractor for a raw reference argument. */
export interface IfCtx {
  utils: { extractRefValue(arg: unknown): { val: unknown; isArray: boolean } };
}

/** A parser error object, which must propagate rather than be coerced. */
const isError = (v: unknown): boolean => typeof v === "object" && v !== null && "_error" in (v as Record<string, unknown>);

const VALUE_ERROR = { _error: "#VALUE!" };

/**
 * Excel's truthiness for an IF condition:
 *   blank    -> FALSE          (an empty cell is not an error)
 *   boolean  -> itself
 *   number   -> non-zero
 *   "TRUE" / "FALSE" (any case) -> the boolean it names
 *   any other text, including "" -> #VALUE!
 * An error propagates unchanged.
 */
export function conditionOf(value: unknown): boolean | typeof VALUE_ERROR | { _error: string } {
  if (isError(value)) return value as { _error: string };
  if (value === null || value === undefined) return false;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (typeof value === "string") {
    const t = value.trim().toUpperCase();
    if (t === "TRUE") return true;
    if (t === "FALSE") return false;
    // Text that does not name a boolean is an error, exactly as Excel reports it.
    return VALUE_ERROR;
  }
  // An array in a scalar context takes its first element, as the parser does elsewhere.
  if (Array.isArray(value)) return conditionOf((value as unknown[][])[0]?.[0] ?? null);
  return VALUE_ERROR;
}

/** Build the IF implementation for the parser's function map. */
export function makeIf(): (ctx: unknown, ...args: unknown[]) => unknown {
  return (ctxRaw: unknown, condArg: unknown, thenArg: unknown, elseArg: unknown): unknown => {
    const ctx = ctxRaw as IfCtx;
    // The arguments arrive un-retrieved, so each one has to be resolved through the parser itself.
    const deref = (arg: unknown): unknown => {
      if (arg === undefined) return undefined;
      try {
        return ctx.utils.extractRefValue(arg).val;
      } catch {
        return arg;
      }
    };
    const cond = conditionOf(deref(condArg));
    if (typeof cond !== "boolean") return cond; // an error condition is the result
    const branch = cond ? thenArg : elseArg;
    // A missing branch is FALSE in Excel, not blank: `IF(FALSE, 1)` is FALSE.
    if (branch === undefined) return cond ? undefined : false;
    const out = deref(branch);
    // A reference to an empty cell reads as 0 in a value context, which is what Excel shows.
    return out === null || out === undefined ? 0 : out;
  };
}
