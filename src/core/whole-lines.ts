import { SPAN_FUNCTIONS } from "./sheet-span";

// The formula engine misreads an anchored whole column or row ($A:$A, $1:$1): the column became
// hundreds of columns wide and the row an error. Anchors mean nothing once a formula is evaluated
// where it stands, so they are dropped before parsing.
const LINE_RE = /("(?:[^"]|"")*")|(?<![\w.$])\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})(?![\w.(])|(?<![\w.$:])\$?(\d+):\$?(\d+)(?![\w.(:])/g;

/** Whether a formula may hold an anchored whole column or row. */
export const hasAnchoredLine = (f: string): boolean => f.includes("$") && f.includes(":");

/** Drop the $ anchors of whole-column and whole-row references, leaving strings alone. */
export function unanchorLines(f: string): string {
  return f.replace(LINE_RE, (all, str: string | undefined, c1?: string, c2?: string, r1?: string, r2?: string) => {
    if (str) return all;
    return c1 ? `${c1}:${c2}` : `${r1}:${r2}`;
  });
}

// Functions for which a whole column passed straight as an argument reads the same cut at the
// last used row: blanks past the data add nothing and no lookup lands there. Decided per argument,
// so COUNTA(A:A) inside OFFSET, or inside a name, is cut even though OFFSET is not on the list.
// INDEX (a row past the data must stay a blank, not #REF!) and COUNTIF/COUNTBLANK are not.
const CUT_ARG_FUNCTIONS = new Set([
  ...SPAN_FUNCTIONS, "SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS", "MAXIFS", "MINIFS", "SUMPRODUCT",
  "MATCH", "XMATCH", "VLOOKUP", "XLOOKUP",
]);

const COLUMN_RE = /(?<![\p{L}\p{N}_.$'!])((?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!)?\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})(?![\p{L}\p{N}_.(!:])/uy;

/** Whether a formula may hold a whole column at all (cheap pre-check). */
export const mayHaveWholeColumn = (f: string): boolean => /[A-Za-z]\s*:\s*\$?[A-Za-z]/.test(f);

/**
 * Whole columns that are a direct argument of a function in CUT_ARG_FUNCTIONS, rewritten to stop
 * at `rows` (SUM(Data!A:A) reads Data!A1:A40); everything else is left as written.
 */
export function cutWholeColumnArgs(f: string, rows: number): string {
  if (!mayHaveWholeColumn(f)) return f;
  const stack: (string | undefined)[] = [];
  let out = "";
  let at = 0;
  let i = 0;
  while (i < f.length) {
    const ch = f[i]!;
    if (ch === '"') {
      i++;
      while (i < f.length) if (f[i++] === '"') { if (f[i] === '"') i++; else break; }
      continue;
    }
    if (ch === "(") {
      const m = /[A-Za-z0-9_.]+$/.exec(f.slice(Math.max(0, i - 64), i));
      stack.push(m ? m[0].toUpperCase().replace(/^_XLFN\./, "").replace(/^_XLWS\./, "") : undefined);
      i++;
      continue;
    }
    if (ch === ")") {
      stack.pop();
      i++;
      continue;
    }
    COLUMN_RE.lastIndex = i;
    const m = ch === "'" || /[\p{L}_$]/u.test(ch) ? COLUMN_RE.exec(f) : null;
    if (!m) {
      if (ch === "'") {
        // A quoted sheet name: skip it so its text is not read as parentheses.
        i++;
        while (i < f.length) if (f[i++] === "'") { if (f[i] === "'") i++; else break; }
      } else i++;
      continue;
    }
    const end = i + m[0].length;
    const fn = stack[stack.length - 1];
    const before = f.slice(0, i).trimEnd().slice(-1);
    const after = f.slice(end).trimStart().charAt(0);
    if (fn && CUT_ARG_FUNCTIONS.has(fn) && (before === "(" || before === ",") && (after === ")" || after === ",")) {
      out += f.slice(at, i) + `${m[1] ?? ""}${m[2]}1:${m[3]}${Math.max(1, rows)}`;
      at = end;
    }
    i = end;
  }
  return at ? out + f.slice(at) : f;
}

