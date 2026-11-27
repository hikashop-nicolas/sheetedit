// R1C1 references (SpreadsheetML formulas, names and ranges) to A1.
import { colToLetters } from "../../core/model";

const MAX_R = 1048576, MAX_C = 16384;
const wrap = (n: number, max: number): number => ((n % max) + max) % max;

interface Part { abs: boolean; n: number }

function part(spec: string | undefined, here: number, max: number): Part {
  if (spec == null || spec === "") return { abs: false, n: here };
  if (spec[0] === "[") return { abs: false, n: wrap(here + Number(spec.slice(1, -1)), max) };
  return { abs: true, n: Number(spec) - 1 };
}

const rowA1 = (p: Part): string => (p.abs ? "$" : "") + (p.n + 1);
const colA1 = (p: Part): string => (p.abs ? "$" : "") + colToLetters(p.n + 1);

const RC = /R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?/y;
const ROW = /R(\[-?\d+\]|\d+)?/y;
const COL = /C(\[-?\d+\]|\d+)?/y;
const NAME_CHAR = /[A-Za-z0-9_.\\?]/;

const sticky = (re: RegExp, s: string, i: number): RegExpExecArray | null => { re.lastIndex = i; return re.exec(s); };

/** Convert an R1C1 formula to A1, relative to the 0-based cell (row, col). */
export function r1c1ToA1(s: string, row = 0, col = 0): string {
  let out = "";
  let i = 0;
  while (i < s.length) {
    const ch = s[i]!;
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < s.length) { if (s[j] === ch) { if (s[j + 1] === ch) { j += 2; continue; } break; } j++; }
      out += s.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "[") {
      // An external workbook or a structured reference: copied whole.
      const j = s.indexOf("]", i);
      if (j < 0) { out += s.slice(i); break; }
      out += s.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if ((ch === "R" || ch === "C") && !(i > 0 && NAME_CHAR.test(s[i - 1]!))) {
      const ok = (end: number): boolean => { const n = s[end] ?? ""; return !NAME_CHAR.test(n) && n !== "(" && n !== "["; };
      const inRange = (end: number): boolean => s[end] === ":" || s[i - 1] === ":";
      const rc = ch === "R" ? sticky(RC, s, i) : null;
      if (rc && ok(i + rc[0].length)) {
        out += colA1(part(rc[2], col, MAX_C)) + rowA1(part(rc[1], row, MAX_R));
        i += rc[0].length;
        continue;
      }
      const one = sticky(ch === "R" ? ROW : COL, s, i);
      if (one && ok(i + one[0].length) && inRange(i + one[0].length)) {
        out += ch === "R" ? rowA1(part(one[1], row, MAX_R)) : colA1(part(one[1], col, MAX_C));
        i += one[0].length;
        continue;
      }
    }
    out += ch;
    i++;
  }
  return out;
}

/** An R1C1 range ("R2C1:R5C3", "C2", "R3") as 0-based bounds, relative to (row, col). */
export function r1c1Range(ref: string, row = 0, col = 0): { r1: number; c1: number; r2: number; c2: number } | null {
  const one = (s: string): { r?: [number, number]; c?: [number, number] } | null => {
    const m = /^R(\[-?\d+\]|\d+)?(?:C(\[-?\d+\]|\d+)?)?$|^C(\[-?\d+\]|\d+)?$/i.exec(s.trim());
    if (!m) return null;
    const hasR = /^R/i.test(s.trim()), hasC = /C/i.test(s.trim());
    const out: { r?: [number, number]; c?: [number, number] } = {};
    if (hasR) { const p = part(m[1], row, MAX_R); out.r = [p.n, p.n]; }
    if (hasC) { const p = part(m[2] ?? m[3], col, MAX_C); out.c = [p.n, p.n]; }
    return out;
  };
  const [a, b] = ref.replace(/^.*!/, "").split(":");
  const pa = a ? one(a) : null;
  if (!pa) return null;
  const pb = b ? one(b) : pa;
  if (!pb) return null;
  const r1 = pa.r?.[0] ?? 0, r2 = pb.r?.[1] ?? (pa.r ? pa.r[1] : MAX_R - 1);
  const c1 = pa.c?.[0] ?? 0, c2 = pb.c?.[1] ?? (pa.c ? pa.c[1] : MAX_C - 1);
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}
