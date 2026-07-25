// LET(name1, value1, [name2, value2, ...], calculation) binds names for the calculation to reuse.
// The parser evaluates a function's arguments eagerly and has no notion of a local scope, so `x` in
// LET(x,1,x+1) would resolve as an undefined name (#NAME?) before LET ever ran. Instead we expand
// LET away textually before parsing: each name is substituted by its (parenthesised) value
// expression, innermost LET first. Substitution is token-aware, so it never rewrites inside a string
// literal and never touches an identifier used as a function call (`name(`).

/** Split `s` on top-level commas, ignoring commas inside parens, braces or string literals. */
function splitArgs(s: string): string[] {
  const out: string[] = [];
  let depth = 0, inStr = false, start = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (inStr) { if (ch === '"') { if (s[i + 1] === '"') i++; else inStr = false; } continue; }
    if (ch === '"') { inStr = true; continue; }
    if (ch === "(" || ch === "{" || ch === "[") depth++;
    else if (ch === ")" || ch === "}" || ch === "]") depth--;
    else if (ch === "," && depth === 0) { out.push(s.slice(start, i)); start = i + 1; }
  }
  out.push(s.slice(start));
  return out;
}

/** Replace whole-identifier occurrences of `name` with `repl`, outside strings and not as a call. */
function substitute(src: string, name: string, repl: string): string {
  const lower = name.toLowerCase();
  let out = "", i = 0, inStr = false;
  while (i < src.length) {
    const ch = src[i]!;
    if (inStr) { out += ch; if (ch === '"') { if (src[i + 1] === '"') { out += src[++i]; } else inStr = false; } i++; continue; }
    if (ch === '"') { inStr = true; out += ch; i++; continue; }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i;
      while (j < src.length && /[A-Za-z0-9_.]/.test(src[j]!)) j++;
      const word = src.slice(i, j);
      // A trailing "(" means it is a function call, and a leading "!" / trailing "!" a sheet ref.
      const isCall = /^\s*\(/.test(src.slice(j));
      const sheetQualified = src[i - 1] === "!" || src[j] === "!";
      out += word.toLowerCase() === lower && !isCall && !sheetQualified ? repl : word;
      i = j;
      continue;
    }
    out += ch; i++;
  }
  return out;
}

/** Find the index just past the `)` matching the `(` at `open`, or -1. */
function matchParen(s: string, open: number): number {
  let depth = 0, inStr = false;
  for (let i = open; i < s.length; i++) {
    const ch = s[i]!;
    if (inStr) { if (ch === '"') { if (s[i + 1] === '"') i++; else inStr = false; } continue; }
    if (ch === '"') { inStr = true; continue; }
    if (ch === "(") depth++;
    else if (ch === ")") { depth--; if (depth === 0) return i; }
  }
  return -1;
}

/** Locate the LAST (therefore innermost-startable) `LET(` token, so nesting expands inside-out. */
function findLet(s: string): number {
  let found = -1, inStr = false;
  for (let i = 0; i + 3 < s.length; i++) {
    const ch = s[i]!;
    if (inStr) { if (ch === '"') { if (s[i + 1] === '"') i++; else inStr = false; } continue; }
    if (ch === '"') { inStr = true; continue; }
    if ((ch === "L" || ch === "l") && /^let\s*\(/i.test(s.slice(i)) && !/[A-Za-z0-9_.]/.test(s[i - 1] ?? "")) found = i;
  }
  return found;
}

/**
 * Expand every LET(...) in a formula into its calculation with the names substituted in.
 * Returns the formula unchanged when it contains no (well-formed) LET.
 */
export function expandLet(formula: string): string {
  let s = formula;
  for (let guard = 0; guard < 50; guard++) {
    const at = findLet(s);
    if (at < 0) return s;
    const open = s.indexOf("(", at);
    const close = matchParen(s, open);
    if (close < 0) return s;
    const args = splitArgs(s.slice(open + 1, close)).map((a) => a.trim());
    // LET takes name/value pairs plus a final calculation, so an odd count of at least 3.
    if (args.length < 3 || args.length % 2 === 0) return s;
    const bindings: { name: string; value: string }[] = [];
    for (let i = 0; i + 1 < args.length - 1; i += 2) {
      let value = args[i + 1]!;
      // A later value may use an earlier name.
      for (const b of bindings) value = substitute(value, b.name, `(${b.value})`);
      bindings.push({ name: args[i]!, value });
    }
    let calc = args[args.length - 1]!;
    for (const b of bindings) calc = substitute(calc, b.name, `(${b.value})`);
    s = s.slice(0, at) + `(${calc})` + s.slice(close + 1);
  }
  return s;
}

/** True when the formula mentions LET at all (cheap guard before running the expander). */
export const hasLet = (f: string): boolean => /\blet\s*\(/i.test(f);
