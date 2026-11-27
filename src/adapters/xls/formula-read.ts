// BIFF formula tokens (Ptg, [MS-XLS] 2.5.198) back into A1 formula text, as the xlsx side stores it.
import { colToLetters } from "../../core/model";
import { Cursor, readStringNoCch } from "./biff";
import { FTAB, FTAB_ARGC } from "./ftab";

export interface FormulaContext {
  biff: number;
  /** 0-based position of the cell the formula belongs to (relative RefN / AreaN tokens). */
  row: number;
  col: number;
  /** BIFF8: an EXTERNSHEET index to its sheet prefix ("Sheet1!", "'[1]Data'!"), or null for #REF. */
  xti(ixti: number): string | null;
  /** BIFF5: an internal 3D reference's sheet prefix by sheet indexes. */
  tabs(first: number, last: number, ixals: number): string | null;
  /** A defined name by its 1-based index. */
  name(index: number): string;
  /** An external / add-in name. */
  nameX(ixti: number, index: number): string;
  /** Decode PtgRef as relative to the cell (CF / DV / shared formulas may use either form). */
  maxRow?: number;
}

export const ERRORS: Record<number, string> = { 0x00: "#NULL!", 0x07: "#DIV/0!", 0x0f: "#VALUE!", 0x17: "#REF!", 0x1d: "#NAME?", 0x24: "#NUM!", 0x2a: "#N/A", 0x2b: "#GETTING_DATA" };

const BIN: Record<number, string> = { 0x03: "+", 0x04: "-", 0x05: "*", 0x06: "/", 0x07: "^", 0x08: "&", 0x09: "<", 0x0a: "<=", 0x0b: "=", 0x0c: ">=", 0x0d: ">", 0x0e: "<>", 0x0f: " ", 0x10: ",", 0x11: ":" };

const numText = (n: number): string => {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const s = String(parseFloat(n.toPrecision(15)));
  return s.replace("e+", "E+").replace("e-", "E-");
};

const strLit = (s: string): string => '"' + s.replace(/"/g, '""') + '"';

class Unsupported extends Error {}

/** Decode a token array (rgce) plus its trailing extra data (rgcb) into formula text, or null. */
export function decodeFormula(rgce: Uint8Array, rgcb: Uint8Array | undefined, ctx: FormulaContext): string | null {
  try {
    return decode(rgce, rgcb, ctx);
  } catch {
    return null;
  }
}

function decode(rgce: Uint8Array, rgcb: Uint8Array | undefined, ctx: FormulaContext): string {
  const b8 = ctx.biff >= 8;
  const c = new Cursor(rgce, 1252);
  const extra = rgcb ? new Cursor(rgcb) : undefined;
  const st: string[] = [];
  const pop = (): string => {
    const v = st.pop();
    if (v === undefined) throw new Unsupported("stack");
    return v;
  };
  const maxRow = b8 ? 0xffff : 0x3fff;
  const maxCol = 0xff;
  const cell = (row: number, col: number, rowRel: boolean, colRel: boolean): string =>
    (colRel ? "" : "$") + colToLetters((col & 0xff) + 1) + (rowRel ? "" : "$") + String(row + 1);
  // Read a cell reference; `n` = relative-offset form (RefN / AreaN).
  const ref = (n: boolean): [number, number, boolean, boolean] => {
    let row: number, col: number, rowRel: boolean, colRel: boolean;
    if (b8) {
      row = c.u16();
      const cw = c.u16();
      col = cw & 0x3fff;
      colRel = (cw & 0x4000) !== 0;
      rowRel = (cw & 0x8000) !== 0;
    } else {
      const rw = c.u16();
      row = rw & 0x3fff;
      colRel = (rw & 0x4000) !== 0;
      rowRel = (rw & 0x8000) !== 0;
      col = c.u8();
    }
    if (n) {
      if (rowRel) row = (ctx.row + (b8 ? (row << 16) >> 16 : (row << 18) >> 18)) & maxRow;
      if (colRel) col = (ctx.col + ((col << 24) >> 24)) & 0xff;
    }
    return [row, col, rowRel, colRel];
  };
  const area = (n: boolean): string => {
    let r1: number, r2: number, c1: number, c2: number, r1Rel: boolean, r2Rel: boolean, c1Rel: boolean, c2Rel: boolean;
    if (b8) {
      r1 = c.u16(); r2 = c.u16();
      const w1 = c.u16(), w2 = c.u16();
      c1 = w1 & 0x3fff; c2 = w2 & 0x3fff;
      c1Rel = (w1 & 0x4000) !== 0; r1Rel = (w1 & 0x8000) !== 0;
      c2Rel = (w2 & 0x4000) !== 0; r2Rel = (w2 & 0x8000) !== 0;
    } else {
      const a = c.u16(), bb = c.u16();
      r1 = a & 0x3fff; r2 = bb & 0x3fff;
      c1Rel = (a & 0x4000) !== 0; r1Rel = (a & 0x8000) !== 0;
      c2Rel = (bb & 0x4000) !== 0; r2Rel = (bb & 0x8000) !== 0;
      c1 = c.u8(); c2 = c.u8();
    }
    if (n) {
      const sr = (v: number): number => (b8 ? (v << 16) >> 16 : (v << 18) >> 18);
      if (r1Rel) r1 = (ctx.row + sr(r1)) & maxRow;
      if (r2Rel) r2 = (ctx.row + sr(r2)) & maxRow;
      if (c1Rel) c1 = (ctx.col + ((c1 << 24) >> 24)) & 0xff;
      if (c2Rel) c2 = (ctx.col + ((c2 << 24) >> 24)) & 0xff;
    }
    // Whole columns / whole rows read the way Excel shows them.
    if (r1 === 0 && r2 === maxRow && !(c1 === 0 && c2 === maxCol))
      return (c1Rel ? "" : "$") + colToLetters((c1 & 0xff) + 1) + ":" + (c2Rel ? "" : "$") + colToLetters((c2 & 0xff) + 1);
    if (c1 === 0 && c2 === maxCol && !(r1 === 0 && r2 === maxRow))
      return (r1Rel ? "" : "$") + String(r1 + 1) + ":" + (r2Rel ? "" : "$") + String(r2 + 1);
    const a = cell(r1, c1, r1Rel, c1Rel), z = cell(r2, c2, r2Rel, c2Rel);
    return a === z ? a : a + ":" + z;
  };
  // BIFF5 3D prefix: ixals, 8 reserved bytes, itabFirst, itabLast.
  const prefix5 = (): string => {
    const ixals = c.i16();
    c.skip(8);
    const t1 = c.i16(), t2 = c.i16();
    const p = ctx.tabs(t1, t2, ixals);
    if (p == null) throw new Unsupported("3d");
    return p;
  };
  const prefix8 = (): string | null => ctx.xti(c.u16());
  while (c.left > 0) {
    const ptg = c.u8();
    if (ptg >= 0x03 && ptg <= 0x11) {
      const b = pop(), a = pop();
      st.push(a + BIN[ptg] + b);
      continue;
    }
    switch (ptg) {
      case 0x01: // PtgExp: resolved by the caller (shared / array formula)
      case 0x02: // PtgTbl: data table
        throw new Unsupported("exp");
      case 0x12: st.push("+" + pop()); break;
      case 0x13: st.push("-" + pop()); break;
      case 0x14: st.push(pop() + "%"); break;
      case 0x15: st.push("(" + pop() + ")"); break;
      case 0x16: st.push(""); break;
      case 0x17: {
        const cch = c.u8();
        st.push(strLit(b8 ? readStringNoCch(c, cch, 8) : c.chars7(cch)));
        break;
      }
      case 0x18: { // PtgExtended (BIFF8 elf / sxname): not used in cell formulas
        throw new Unsupported("ext");
      }
      case 0x19: {
        const grbit = c.u8();
        const w = c.u16();
        if (grbit & 0x04) { c.skip((w + 1) * 2); break; } // choose jump table
        if (grbit & 0x10) st.push("SUM(" + pop() + ")");
        // volatile / if / goto / space carry no text of their own
        break;
      }
      case 0x1c: st.push(ERRORS[c.u8()] ?? "#N/A"); break;
      case 0x1d: st.push(c.u8() ? "TRUE" : "FALSE"); break;
      case 0x1e: st.push(String(c.u16())); break;
      case 0x1f: st.push(numText(c.f64())); break;
      default: {
        if (ptg < 0x20 || ptg > 0x7f) throw new Unsupported("ptg " + ptg);
        const base = (ptg & 0x1f) | 0x20;
        switch (base) {
          case 0x20: { // PtgArray: the values live in rgcb
            c.skip(7);
            if (!extra) throw new Unsupported("array");
            st.push(readArray(extra, ctx.biff));
            break;
          }
          case 0x21: case 0x22: {
            let argc: number, iftab: number;
            if (base === 0x22) {
              argc = c.u8() & 0x7f;
              iftab = ctx.biff <= 3 ? c.u8() : c.u16() & 0x7fff;
            } else {
              iftab = ctx.biff <= 3 ? c.u8() : c.u16();
              const n = FTAB_ARGC.get(iftab);
              if (n == null) throw new Unsupported("argc");
              argc = n;
            }
            const args: string[] = [];
            for (let i = 0; i < argc; i++) args.unshift(pop());
            if (iftab === 255) {
              // User-defined / add-in: the first argument is the function's name.
              const fn = args.shift() ?? "";
              st.push(fn + "(" + args.join(",") + ")");
            } else {
              const name = FTAB[iftab];
              if (!name) throw new Unsupported("func " + iftab);
              st.push(name + "(" + args.join(",") + ")");
            }
            break;
          }
          case 0x23: { // PtgName
            const idx = c.u16();
            c.skip(b8 ? 2 : 12);
            st.push(ctx.name(idx));
            break;
          }
          case 0x24: { const [r, cc, rr, cr] = ref(false); st.push(cell(r, cc, rr, cr)); break; }
          case 0x25: st.push(area(false)); break;
          case 0x26: case 0x27: case 0x28: case 0x29: case 0x2e: case 0x2f: {
            // Mem tokens: the sub-expression that follows is decoded normally.
            c.skip(base <= 0x28 ? 6 : 2);
            if (base === 0x26 && extra) { const n = extra.u16(); extra.skip(n * (b8 ? 8 : 6)); }
            break;
          }
          case 0x2a: c.skip(b8 ? 4 : 3); st.push("#REF!"); break;
          case 0x2b: c.skip(b8 ? 8 : 6); st.push("#REF!"); break;
          case 0x2c: { const [r, cc, rr, cr] = ref(true); st.push(cell(r, cc, rr, cr)); break; }
          case 0x2d: st.push(area(true)); break;
          case 0x39: { // PtgNameX
            if (b8) {
              const ixti = c.u16(), idx = c.u16();
              c.skip(2);
              st.push(ctx.nameX(ixti, idx));
            } else {
              const ixals = c.i16();
              c.skip(8);
              const idx = c.u16();
              c.skip(12);
              st.push(ctx.nameX(ixals, idx));
            }
            break;
          }
          case 0x3a: {
            const p = b8 ? prefix8() : prefix5();
            const [r, cc, rr, cr] = ref(false);
            st.push(p == null ? "#REF!" : p + cell(r, cc, rr, cr));
            break;
          }
          case 0x3b: {
            const p = b8 ? prefix8() : prefix5();
            const a = area(false);
            st.push(p == null ? "#REF!" : p + a);
            break;
          }
          case 0x3c: { if (b8) c.skip(6); else { prefix5(); c.skip(3); } st.push("#REF!"); break; }
          case 0x3d: { if (b8) c.skip(10); else { prefix5(); c.skip(6); } st.push("#REF!"); break; }
          default:
            throw new Unsupported("ptg " + ptg);
        }
      }
    }
  }
  if (st.length !== 1) throw new Unsupported("stack " + st.length);
  return st[0]!;
}

/** A PtgArray's constant values from rgcb: {1,2;3,4}. */
function readArray(c: Cursor, biff: number): string {
  const cols = c.u8() + 1;
  const rows = c.u16() + 1;
  const out: string[] = [];
  for (let r = 0; r < rows; r++) {
    const row: string[] = [];
    for (let k = 0; k < cols; k++) {
      const t = c.u8();
      if (t === 0x01) row.push(numText(c.f64()));
      else if (t === 0x02) {
        if (biff >= 8) { const cch = c.u16(); row.push(strLit(readStringNoCch(c, cch, 8))); }
        else { const cch = c.u8(); row.push(strLit(c.chars7(cch))); }
      } else if (t === 0x04) { row.push(c.u8() ? "TRUE" : "FALSE"); c.skip(7); }
      else if (t === 0x10) { row.push(ERRORS[c.u8()] ?? "#N/A"); c.skip(7); }
      else { c.skip(8); row.push(""); }
    }
    out.push(row.join(","));
  }
  return "{" + out.join(";") + "}";
}
