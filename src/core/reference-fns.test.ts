import { describe, expect, it } from "vitest";
import { readCsv } from "../adapters/csv/read";
import { getCell, ensureCell } from "./model";
import { recalc } from "./recalc";
import { expandLet } from "./let-expand";

// A1:A5 = 3,1,4,1,5 ; B1:B5 = a,b,c,d,e
const DATA = "3,a\n1,b\n4,c\n1,d\n5,e\n";

function ev(formula: string, data = DATA): string {
  const wb = readCsv(data);
  const s = wb.sheets[0]!;
  s.cells.set("9:9", { row: 9, col: 9, value: "", kind: "blank", formula } as never);
  s.maxRow = Math.max(s.maxRow, 9); s.maxCol = Math.max(s.maxCol, 9);
  recalc(wb);
  return getCell(s, 9, 9)?.value ?? "";
}

describe("OFFSET", () => {
  it("shifts a reference by rows and columns", () => {
    expect(ev("OFFSET(A1,2,0)")).toBe("4"); // A3
    expect(ev("OFFSET(A1,1,1)")).toBe("b"); // B2
    expect(ev("OFFSET(A3,-2,0)")).toBe("3"); // back up to A1
  });

  it("resizes with height / width, and the range feeds an aggregate", () => {
    expect(ev("SUM(OFFSET(A1,0,0,3,1))")).toBe("8");  // A1:A3 = 3+1+4
    expect(ev("SUM(OFFSET(A1,1,0,2,1))")).toBe("5");  // A2:A3 = 1+4
    expect(ev("COUNT(OFFSET(A1,0,0,5,1))")).toBe("5");
  });

  it("returns #REF! when the result falls off the sheet or is degenerate", () => {
    expect(ev("OFFSET(A1,-1,0)")).toBe("#REF!");
    expect(ev("OFFSET(A1,0,-1)")).toBe("#REF!");
    expect(ev("OFFSET(A1,0,0,0,1)")).toBe("#REF!"); // zero height
  });
});

describe("INDIRECT", () => {
  it("resolves A1-style text to a cell or range", () => {
    expect(ev('INDIRECT("A3")')).toBe("4");
    expect(ev('INDIRECT("B2")')).toBe("b");
    expect(ev('SUM(INDIRECT("A1:A3"))')).toBe("8");
  });

  it("supports R1C1 style when the second argument is FALSE", () => {
    expect(ev('INDIRECT("R2C1",FALSE)')).toBe("1");
    expect(ev('INDIRECT("R3C2",FALSE)')).toBe("c");
  });

  it("builds the address from other cells, and errors on nonsense", () => {
    // "A" & 3 -> "A3"
    expect(ev('INDIRECT("A"&3)')).toBe("4");
    expect(ev('INDIRECT("not a ref")')).toBe("#REF!");
  });
});

describe("LET", () => {
  it("binds a name for the calculation, including several in sequence", () => {
    expect(ev("LET(x,1,x+1)")).toBe("2");
    expect(ev("LET(x,2,y,3,x*y)")).toBe("6");
    expect(ev("LET(x,2,y,x*5,y+1)")).toBe("11"); // a later value may use an earlier name
  });

  it("binds ranges, references and text", () => {
    expect(ev("LET(total,SUM(A1:A5),total/2)")).toBe("7");
    expect(ev("LET(x,A1,x*2)")).toBe("6");
    expect(ev('LET(s,"ab",s&"c")')).toBe("abc");
    expect(ev('LET(x,5,IF(x>3,"big","small"))')).toBe("big");
  });

  it("nests, and composes inside another call", () => {
    expect(ev("LET(x,1,LET(y,2,x+y))")).toBe("3");
    expect(ev("SUM(LET(x,2,x),10)")).toBe("12");
  });

  it("never substitutes a name inside a string literal", () => {
    expect(ev('LET(x,1,"x is literal")')).toBe("x is literal");
    expect(ev('LET(x,1,"x"&x)')).toBe("x1");
  });

  it("tracks dependencies, so editing a source cell recalculates", () => {
    const wb = readCsv(DATA);
    const s = wb.sheets[0]!;
    s.cells.set("9:9", { row: 9, col: 9, value: "", kind: "blank", formula: "LET(t,SUM(A1:A5),t*2)" } as never);
    s.maxRow = 9; s.maxCol = 9;
    recalc(wb);
    expect(getCell(s, 9, 9)?.value).toBe("28"); // (3+1+4+1+5)*2
    const a1 = ensureCell(s, 1, 1);
    a1.value = "13"; a1.kind = "n";
    recalc(wb);
    expect(getCell(s, 9, 9)?.value).toBe("48"); // (13+1+4+1+5)*2
  });

  it("leaves a malformed LET alone rather than mangling the formula", () => {
    expect(expandLet("LET(x,1)")).toBe("LET(x,1)");       // too few args
    expect(expandLet("LET(x,1,y,2)")).toBe("LET(x,1,y,2)"); // even count: no calculation
    expect(expandLet("SUM(A1:A5)")).toBe("SUM(A1:A5)");   // no LET at all
  });
});
