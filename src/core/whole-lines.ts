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

