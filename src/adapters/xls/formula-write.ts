// A1 formula text compiled into BIFF8 tokens (Ptg RPN) with Excel's operand classes: reference
// for range parameters, value for scalars, array inside array parameters and array formulas.
import { lettersToCol } from "../../core/model";
import { FTAB, FTAB_ARGC, ftabIndex } from "./ftab";

export type Ctx = "R" | "V" | "A";

export interface CompileEnv {
  /** Sheet index of a sheet name, or -1. */
  sheetIndex(name: string): number;
  /** EXTERNSHEET index for a sheet range in this workbook, or an external book ([n]). */
  xti(first: number, last: number, external?: { book: number; first: string; last: string }): number;
  /** 1-based NAME index for a defined name seen from `sheet`. */
  nameIndex(name: string): number;
  /** EXTERNSHEET index and 1-based EXTERNNAME index for an Analysis ToolPak function. */
  addinFunc(name: string): { ixti: number; index: number };
  /** 1-based NAME index of the function name record for a future / user-defined function. */
  funcName(name: string): number;
  /** NameX target for an external workbook's defined name ([n]!Name). */
  externName(book: number, name: string): { ixti: number; index: number };
  /** NameX target for a name scoped to another sheet (Sheet!Name). */
  sheetName(sheet: string, name: string): { ixti: number; index: number };
  /** Index of the sheet the formula lives on (to decide 2D vs 3D refs), or -1 for a defined name. */
  sheet: number;
  /** Force 3D references (defined names). */
  force3d?: boolean;
}

export class CompileError extends Error {}

// Parameter classes per function: one letter per parameter, the last repeating. Unlisted
// functions take value parameters, except that a reference argument stays a reference.
const PARAMS: Record<string, string> = {
  SUM: "R", AVERAGE: "R", MIN: "R", MAX: "R", COUNT: "R", COUNTA: "R", PRODUCT: "R", STDEV: "R", STDEVP: "R", VAR: "R", VARP: "R",
  MEDIAN: "R", MODE: "A", AVEDEV: "R", DEVSQ: "R", GEOMEAN: "R", HARMEAN: "R", SUMSQ: "R", KURT: "R", SKEW: "R", MAXA: "R", MINA: "R",
  AVERAGEA: "R", STDEVA: "R", STDEVPA: "R", VARA: "R", VARPA: "R", AND: "R", OR: "R", NPV: "VR", IRR: "RV", MIRR: "RVV",
  LARGE: "RV", SMALL: "RV", RANK: "VRV", PERCENTILE: "RV", QUARTILE: "RV", PERCENTRANK: "RVV", TRIMMEAN: "RV",
  SUMIF: "RVR", COUNTIF: "RV", COUNTBLANK: "R", SUMPRODUCT: "A", SUMX2MY2: "A", SUMX2PY2: "A", SUMXMY2: "A",
  VLOOKUP: "VRVV", HLOOKUP: "VRVV", MATCH: "VRR", LOOKUP: "VRR", INDEX: "RVVV", OFFSET: "RVVVV", CHOOSE: "VR",
  ROW: "R", COLUMN: "R", ROWS: "R", COLUMNS: "R", AREAS: "R", ISREF: "R", CELL: "VR", SUBTOTAL: "VR", IF: "VRR",
  DSUM: "RRR", DAVERAGE: "RRR", DCOUNT: "RRR", DCOUNTA: "RRR", DMAX: "RRR", DMIN: "RRR", DPRODUCT: "RRR", DSTDEV: "RRR", DSTDEVP: "RRR", DVAR: "RRR", DVARP: "RRR", DGET: "RRR",
  MMULT: "AA", MDETERM: "A", MINVERSE: "A", TRANSPOSE: "A", LINEST: "AAVV", LOGEST: "AAVV", TREND: "AAAV", GROWTH: "AAAV",
  FORECAST: "VAA", SLOPE: "AA", INTERCEPT: "AA", CORREL: "AA", PEARSON: "AA", RSQ: "AA", STEYX: "AA", COVAR: "AA", FREQUENCY: "AA",
  CHITEST: "AA", TTEST: "AAVV", FTEST: "AA", PROB: "AAVV", ZTEST: "RVV", COUNTIFS: "RV", SUMIFS: "RRV", AVERAGEIF: "RVR", AVERAGEIFS: "RRV",
  IFERROR: "VV", CONCATENATE: "V", HYPERLINK: "VV", SERIESSUM: "VVVA", MULTINOMIAL: "R", GCD: "R", LCM: "R", XNPV: "VAA", XIRR: "AAV",
};
// Functions whose result is a reference (kept in reference class where a reference is wanted).
const REF_RESULT = new Set(["OFFSET", "INDEX", "INDIRECT", "IF", "CHOOSE"]);
const VOLATILE = new Set(["RAND", "NOW", "TODAY", "OFFSET", "INDIRECT", "CELL", "INFO", "RANDBETWEEN"]);
const ERR: Record<string, number> = { "#NULL!": 0x00, "#DIV/0!": 0x07, "#VALUE!": 0x0f, "#REF!": 0x17, "#NAME?": 0x1d, "#NUM!": 0x24, "#N/A": 0x2a, "#GETTING_DATA": 0x2b };
// Analysis ToolPak functions: BIFF8 calls them through the add-in EXTERNNAME table.
const ATP = /^(HEX2|DEC2|OCT2|BIN2|IM[A-Z]|COMPLEX$|SERIESSUM$|FACTDOUBLE$|SQRTPI$|QUOTIENT$|DELTA$|GESTEP$|ISEVEN$|ISODD$|MROUND$|ERFC?$|BESSEL|XIRR$|XNPV$|PRICE|YIELD|INTRATE$|RECEIVED$|DISC$|TBILL|DOLLARDE$|DOLLARFR$|NOMINAL$|EFFECT$|CUMPRINC$|CUMIPMT$|EDATE$|EOMONTH$|YEARFRAC$|COUP|DURATION$|MDURATION$|ODD[FL]|RANDBETWEEN$|WEEKNUM$|AMOR|CONVERT$|ACCRINTM?$|WORKDAY$|NETWORKDAYS$|GCD$|MULTINOMIAL$|LCM$|FVSCHEDULE$)/;

type Tok =
  | { t: "num"; v: number } | { t: "str"; v: string } | { t: "bool"; v: boolean } | { t: "err"; v: string }
  | { t: "ref"; text: string } | { t: "name"; v: string } | { t: "func"; v: string } | { t: "op"; v: string }
  | { t: "("; } | { t: ")"; } | { t: ","; } | { t: "{"; } | { t: "}"; } | { t: ";" };

const SHEET_Q = "'(?:[^']|'')+'";
const SHEET_U = "[A-Za-z_\\u00a1-\\uffff][A-Za-z0-9_.\\u00a1-\\uffff]*";
const PREFIX = `(?:(?:${SHEET_Q}|\\[\\d+\\]${SHEET_U}(?::${SHEET_U})?|${SHEET_U}(?::${SHEET_U})?|\\[\\d+\\])!)`;
const CELL = "\\$?[A-Za-z]{1,3}\\$?\\d+";
const REF_RE = new RegExp(`^(${PREFIX})?(${CELL}(?::${CELL})?|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}|\\$?\\d+:\\$?\\d+)(?![A-Za-z0-9_(.!])`);

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i]!;
    if (ch === " " || ch === "\n" || ch === "\r" || ch === "\t") {
      // A space between two references is the intersection operator.
      let j = i;
      while (j < src.length && /\s/.test(src[j]!)) j++;
      const prev = out[out.length - 1];
      if (prev && (prev.t === "ref" || prev.t === "name" || prev.t === ")") && j < src.length && /[A-Za-z_$'(]/.test(src[j]!)) out.push({ t: "op", v: " " });
      i = j;
      continue;
    }
    if (ch === '"') {
      let j = i + 1, s = "";
      while (j < src.length) {
        if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += src[j++];
      }
      out.push({ t: "str", v: s });
      i = j + 1;
      continue;
    }
    if (ch === "#") {
      const m = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA)/i.exec(src.slice(i));
      if (!m) throw new CompileError("bad error literal");
      out.push({ t: "err", v: m[0].toUpperCase() });
      i += m[0].length;
      continue;
    }
    const rest = src.slice(i);
    const prev = out[out.length - 1];
    const operandAllowed = !prev || prev.t === "op" || prev.t === "(" || prev.t === "func" || prev.t === "," || prev.t === "{" || prev.t === ";" ;
    const num = /^(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/.exec(rest);
    if (num && operandAllowed && !/^\d+:\$?\d+/.test(rest)) {
      out.push({ t: "num", v: Number(num[0]) });
      i += num[0].length;
      continue;
    }
    const ref = REF_RE.exec(rest);
    if (ref && operandAllowed && inGrid(ref[0])) {
      out.push({ t: "ref", text: ref[0] });
      i += ref[0].length;
      continue;
    }
    const word = new RegExp(`^(${PREFIX})?[A-Za-z_\\\\\\u00a1-\\uffff][A-Za-z0-9_.\\u00a1-\\uffff]*`).exec(rest);
    if (word) {
      const w = word[0];
      i += w.length;
      if (src[i] === "(") { out.push({ t: "func", v: w }); i++; continue; }
      if (/^(TRUE|FALSE)$/i.test(w)) { out.push({ t: "bool", v: w.toUpperCase() === "TRUE" }); continue; }
      if (/\[/.test(src[i] ?? "")) throw new CompileError("structured reference");
      out.push({ t: "name", v: w });
      continue;
    }
    const two = rest.slice(0, 2);
    if (two === "<=" || two === ">=" || two === "<>") { out.push({ t: "op", v: two }); i += 2; continue; }
    if ("+-*/^&=<>%:".includes(ch)) { out.push({ t: "op", v: ch }); i++; continue; }
    if (ch === "(" || ch === ")" || ch === "," || ch === "{" || ch === "}" || ch === ";") { out.push({ t: ch } as Tok); i++; continue; }
    throw new CompileError(`unexpected "${ch}"`);
  }
  return out;
}

type Node =
  | { k: "num"; v: number } | { k: "str"; v: string } | { k: "bool"; v: boolean } | { k: "err"; v: string } | { k: "miss" }
  | { k: "ref"; text: string } | { k: "name"; v: string } | { k: "func"; name: string; args: Node[] }
  | { k: "bin"; op: string; a: Node; b: Node } | { k: "un"; op: string; a: Node } | { k: "pct"; a: Node } | { k: "paren"; a: Node }
  | { k: "array"; rows: (number | string | boolean | { err: string })[][] };

const BIN_PREC: Record<string, number> = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5, " ": 7, ":": 8 };

/** A cell-shaped token beyond IV65536 is a name in BIFF8, not a reference. */
function inGrid(text: string): boolean {
  const body = text.slice(text.lastIndexOf("!") + 1);
  for (const part of body.split(":")) {
    const m = /^\$?([A-Za-z]{1,3})?\$?(\d+)?$/.exec(part);
    if (!m) return false;
    if (m[1] && lettersToCol(m[1].toUpperCase()) > 256) return false;
    if (m[2] && Number(m[2]) > 65536) return false;
  }
  return true;
}

function parse(toks: Tok[]): Node {
  let p = 0;
  const peek = (): Tok | undefined => toks[p];
  const expr = (min: number): Node => {
    let left = unary();
    for (;;) {
      const t = peek();
      if (t?.t === "op" && t.v === "%") { p++; left = { k: "pct", a: left }; continue; }
      if (t?.t !== "op" || BIN_PREC[t.v] == null || BIN_PREC[t.v]! < min) break;
      p++;
      const prec = BIN_PREC[t.v]!;
      const right = expr(t.v === "^" ? prec : prec + 1);
      left = { k: "bin", op: t.v, a: left, b: right };
    }
    return left;
  };
  const unary = (): Node => {
    const t = peek();
    if (t?.t === "op" && (t.v === "-" || t.v === "+")) { p++; return { k: "un", op: t.v, a: expr(6) }; }
    return primary();
  };
  const primary = (): Node => {
    const t = toks[p++];
    if (!t) throw new CompileError("unexpected end");
    switch (t.t) {
      case "num": return { k: "num", v: t.v };
      case "str": return { k: "str", v: t.v };
      case "bool": return { k: "bool", v: t.v };
      case "err": return { k: "err", v: t.v };
      case "ref": return { k: "ref", text: t.text };
      case "name": return { k: "name", v: t.v };
      case "(": {
        const inner = union();
        if (toks[p]?.t !== ")") throw new CompileError("missing )");
        p++;
        return { k: "paren", a: inner };
      }
      case "{": {
        const rows: (number | string | boolean | { err: string })[][] = [[]];
        for (;;) {
          let neg = false;
          let tk = toks[p++];
          if (tk?.t === "op" && tk.v === "-") { neg = true; tk = toks[p++]; }
          if (!tk) throw new CompileError("bad array");
          if (tk.t === "num") rows[rows.length - 1]!.push(neg ? -tk.v : tk.v);
          else if (tk.t === "str") rows[rows.length - 1]!.push(tk.v);
          else if (tk.t === "bool") rows[rows.length - 1]!.push(tk.v);
          else if (tk.t === "err") rows[rows.length - 1]!.push({ err: tk.v });
          else throw new CompileError("bad array");
          const sep = toks[p++];
          if (sep?.t === "}") break;
          if (sep?.t === ";") rows.push([]);
          else if (sep?.t !== ",") throw new CompileError("bad array");
        }
        return { k: "array", rows };
      }
      case "func": {
        const args: Node[] = [];
        if (toks[p]?.t === ")") { p++; return { k: "func", name: t.v, args }; }
        for (;;) {
          const nt = toks[p];
          if (nt?.t === "," || nt?.t === ")") args.push({ k: "miss" });
          else args.push(expr(0));
          const sep = toks[p++];
          if (sep?.t === ")") break;
          if (sep?.t !== ",") throw new CompileError("missing , or )");
        }
        return { k: "func", name: t.v, args };
      }
      default:
        throw new CompileError("unexpected token");
    }
  };
  // Outside a function's argument list a comma is the union operator.
  function union(): Node {
    let left = expr(0);
    while (toks[p]?.t === ",") { p++; left = { k: "bin", op: ",", a: left, b: expr(0) }; }
    return left;
  }
  const root = union();
  if (p !== toks.length) throw new CompileError("trailing tokens");
  return root;
}

const BIN_PTG: Record<string, number> = { " ": 0x0f, ",": 0x10, "+": 0x03, "-": 0x04, "*": 0x05, "/": 0x06, "^": 0x07, "&": 0x08, "<": 0x09, "<=": 0x0a, "=": 0x0b, ">=": 0x0c, ">": 0x0d, "<>": 0x0e, ":": 0x11 };

class Out {
  b: number[] = [];
  extra: number[] = [];
  u8(v: number): void { this.b.push(v & 255); }
  u16(v: number): void { this.b.push(v & 255, (v >> 8) & 255); }
  f64(v: number, to = this.b): void {
    const d = new DataView(new ArrayBuffer(8));
    d.setFloat64(0, v, true);
    for (let i = 0; i < 8; i++) to.push(d.getUint8(i));
  }
}

const cls = (base: number, c: Ctx): number => base + (c === "R" ? 0 : c === "V" ? 0x20 : 0x40);
const unquote = (s: string): string => (s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s);

/** Compile formula text (no leading "=") to rgce + rgcb. Throws CompileError when BIFF8 cannot hold it. */
export function compileFormula(text: string, env: CompileEnv, top: Ctx = "V"): { rgce: Uint8Array; rgcb: Uint8Array } {
  const root = parse(tokenize(text));
  const o = new Out();
  let volatile = false;
  const ref = (text: string, c: Ctx): void => {
    let prefix = "";
    let body = text;
    const bang = text.lastIndexOf("!");
    if (bang >= 0) { prefix = text.slice(0, bang); body = text.slice(bang + 1); }
    const parts = body.split(":");
    type P = { row: number; col: number; rowRel: boolean; colRel: boolean };
    const parsePart = (s: string, kind: "cell" | "col" | "row"): P => {
      const m = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/.exec(s);
      if (!m) throw new CompileError("bad ref");
      const col = m[2] ? lettersToCol(m[2].toUpperCase()) - 1 : kind === "row" ? 0 : 0;
      const row = m[4] ? Number(m[4]) - 1 : 0;
      return { row, col, rowRel: !(m[2] ? m[3] : m[1]), colRel: !m[1] || !m[2] ? !m[1] && !!m[2] : !m[1] };
    };
    let area: [P, P] | null = null;
    let one: P | null = null;
    if (parts.length === 2 && /^\$?[A-Za-z]{1,3}$/.test(parts[0]!) && /^\$?[A-Za-z]{1,3}$/.test(parts[1]!)) {
      const a = parsePart(parts[0]!, "col"), b = parsePart(parts[1]!, "col");
      area = [{ ...a, row: 0, rowRel: false, colRel: !parts[0]!.startsWith("$") }, { ...b, row: 0xffff, rowRel: false, colRel: !parts[1]!.startsWith("$") }];
    } else if (parts.length === 2 && /^\$?\d+$/.test(parts[0]!) && /^\$?\d+$/.test(parts[1]!)) {
      const r1 = Number(parts[0]!.replace("$", "")) - 1, r2 = Number(parts[1]!.replace("$", "")) - 1;
      area = [{ row: r1, col: 0, rowRel: !parts[0]!.startsWith("$"), colRel: false }, { row: r2, col: 0xff, rowRel: !parts[1]!.startsWith("$"), colRel: false }];
    } else if (parts.length === 2) {
      const cell = (s: string): P => {
        const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(s);
        if (!m) throw new CompileError("bad ref");
        return { col: lettersToCol(m[2]!.toUpperCase()) - 1, row: Number(m[4]) - 1, colRel: !m[1], rowRel: !m[3] };
      };
      area = [cell(parts[0]!), cell(parts[1]!)];
    } else {
      const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(parts[0]!);
      if (!m) throw new CompileError("bad ref");
      one = { col: lettersToCol(m[2]!.toUpperCase()) - 1, row: Number(m[4]) - 1, colRel: !m[1], rowRel: !m[3] };
    }
    for (const q of area ?? [one!]) if (q.row > 0xffff || q.col > 0xff || q.row < 0) throw new CompileError("outside the BIFF8 grid");
    const colWord = (q: P): number => (q.col & 0x3fff) | (q.colRel ? 0x4000 : 0) | (q.rowRel ? 0x8000 : 0);
    let ixti = -1;
    if (prefix || env.force3d) {
      let first: number, last: number;
      const ext = /^'?\[(\d+)\](.*?)'?$/.exec(prefix);
      if (ext) {
        const [s1, s2] = unquote(prefix.replace(/^'?\[\d+\]/, "'").replace(/^''/, "'")).split(":");
        ixti = env.xti(0, 0, { book: Number(ext[1]), first: s1 ?? "", last: s2 ?? s1 ?? "" });
      } else if (prefix) {
        const names = unquote(prefix).split(":");
        first = env.sheetIndex(names[0]!);
        last = names[1] != null ? env.sheetIndex(names[1]) : first;
        if (first < 0 || last < 0) throw new CompileError("unknown sheet");
        ixti = env.xti(first, last); // an explicit sheet stays a 3D reference, as Excel keeps it
      } else if (env.sheet >= 0) ixti = env.xti(env.sheet, env.sheet);
    }
    if (ixti >= 0) {
      o.u8(cls(area ? 0x3b : 0x3a, c));
      o.u16(ixti);
    } else o.u8(cls(area ? 0x25 : 0x24, c));
    if (area) { o.u16(area[0].row); o.u16(area[1].row); o.u16(colWord(area[0])); o.u16(colWord(area[1])); }
    else { o.u16(one!.row); o.u16(colWord(one!)); }
  };
  const emit = (n: Node, c: Ctx): void => {
    switch (n.k) {
      case "num":
        if (Number.isInteger(n.v) && n.v >= 0 && n.v <= 0xffff) { o.u8(0x1e); o.u16(n.v); }
        else { o.u8(0x1f); o.f64(n.v); }
        return;
      case "str": {
        if (n.v.length > 255) throw new CompileError("string longer than 255 characters");
        o.u8(0x17); o.u8(n.v.length);
        const wide = /[^\u0000-ÿ]/.test(n.v);
        o.u8(wide ? 1 : 0);
        for (const ch of n.v) { const code = ch.charCodeAt(0); if (wide) o.u16(code); else o.u8(code); }
        return;
      }
      case "bool": o.u8(0x1d); o.u8(n.v ? 1 : 0); return;
      case "err":
        // A deleted reference is a PtgRefErr, which is how Excel stores #REF! in a formula.
        if (n.v === "#REF!") { o.u8(cls(0x2a, c)); o.u16(0); o.u16(0); return; }
        o.u8(0x1c); o.u8(ERR[n.v] ?? 0x2a); return;
      case "miss": o.u8(0x16); return;
      case "ref": ref(n.text, c === "A" ? "A" : c); return;
      case "name": {
        const bang = n.v.lastIndexOf("!");
        const nm = bang >= 0 ? n.v.slice(bang + 1) : n.v;
        if (bang < 0) { o.u8(cls(0x23, c)); o.u16(env.nameIndex(nm)); o.u16(0); return; }
        // [n]!Name is an external workbook's name; Sheet!Name a name scoped to another sheet.
        const prefix = unquote(n.v.slice(0, bang));
        const ext = /^\[(\d+)\]$/.exec(prefix);
        const target = ext ? env.externName(Number(ext[1]), nm) : env.sheetName(prefix, nm);
        o.u8(cls(0x39, c)); o.u16(target.ixti); o.u16(target.index); o.u16(0);
        return;
      }
      case "paren": emit(n.a, c); o.u8(0x15); return;
      case "un": emit(n.a, c === "R" ? "V" : c); o.u8(n.op === "-" ? 0x13 : 0x12); return;
      case "pct": emit(n.a, c === "R" ? "V" : c); o.u8(0x14); return;
      case "bin": {
        if (n.op === ":" || n.op === " " || n.op === ",") { emit(n.a, "R"); emit(n.b, "R"); o.u8(BIN_PTG[n.op]!); return; }
        const sub = c === "A" ? "A" : "V";
        emit(n.a, sub); emit(n.b, sub); o.u8(BIN_PTG[n.op]!);
        return;
      }
      case "array": {
        o.u8(cls(0x20, c === "R" ? "A" : c));
        for (let i = 0; i < 7; i++) o.u8(0);
        const rows = n.rows, cols = rows[0]?.length ?? 0;
        if (!cols || rows.some((r) => r.length !== cols)) throw new CompileError("ragged array");
        const x = o.extra;
        x.push((cols - 1) & 255, (rows.length - 1) & 255, ((rows.length - 1) >> 8) & 255);
        for (const r of rows) for (const v of r) {
          if (typeof v === "number") { x.push(1); o.f64(v, x); }
          else if (typeof v === "string") {
            x.push(2, v.length & 255, v.length >> 8, 1);
            for (const ch of v) { const code = ch.charCodeAt(0); x.push(code & 255, code >> 8); }
          } else if (typeof v === "boolean") x.push(4, v ? 1 : 0, 0, 0, 0, 0, 0, 0, 0);
          else x.push(16, ERR[v.err] ?? 0x2a, 0, 0, 0, 0, 0, 0, 0);
        }
        return;
      }
      case "func": {
        const raw = n.name.replace(/^_xlfn\./i, "").replace(/^_xlws\./i, "");
        const name = raw.toUpperCase();
        if (VOLATILE.has(name)) volatile = true;
        const params = PARAMS[name] ?? "";
        const argClass = (i: number, arg: Node): Ctx => {
          if (c === "A") return "A";
          const p = params ? (params[i] ?? params[params.length - 1]!) : arg.k === "ref" || arg.k === "name" ? "R" : "V";
          return p as Ctx;
        };
        let idx = ftabIndex(name);
        if (idx != null && idx >= 380) idx = undefined;
        const retCls: Ctx = c === "R" && !REF_RESULT.has(name) ? "V" : c;
        if (idx == null) {
          // Add-in, future or user-defined function: PtgFuncVar(255) whose first argument names it.
          if (ATP.test(name)) {
            const ext = env.addinFunc(name);
            o.u8(cls(0x39, "R")); o.u16(ext.ixti); o.u16(ext.index); o.u16(0);
          } else { o.u8(0x23); o.u16(env.funcName(raw)); o.u16(0); }
          n.args.forEach((a, i) => emit(a, argClass(i, a)));
          o.u8(cls(0x22, retCls)); o.u8(n.args.length + 1); o.u16(255);
          return;
        }
        n.args.forEach((a, i) => emit(a, argClass(i, a)));
        const fixed = FTAB_ARGC.get(idx);
        if (fixed != null && fixed === n.args.length) { o.u8(cls(0x21, retCls)); o.u16(idx); }
        else { o.u8(cls(0x22, retCls)); o.u8(n.args.length); o.u16(idx); }
        void FTAB;
        return;
      }
    }
  };
  emit(root, top);
  let rgce = o.b;
  if (volatile) rgce = [0x19, 0x01, 0x00, 0x00, ...rgce];
  return { rgce: new Uint8Array(rgce), rgcb: new Uint8Array(o.extra) };
}
