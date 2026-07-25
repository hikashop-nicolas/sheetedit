import { describe, expect, it } from "vitest";
import { readCsv } from "../adapters/csv/read";
import { getCell } from "./model";
import { recalc } from "./recalc";

// A1:A5 = 3,1,4,1,5 ; B1:B5 = a,b,c,b,a ; C1:C5 = 44000..44004
const DATA = "3,a,44000\n1,b,44001\n4,c,44002\n1,b,44003\n5,a,44004\n";

/** Evaluate one formula against DATA and return the anchor cell's value. */
function ev(formula: string, data = DATA): string {
  const wb = readCsv(data);
  const s = wb.sheets[0]!;
  s.cells.set("9:9", { row: 9, col: 9, value: "", kind: "blank", formula } as never);
  s.maxRow = Math.max(s.maxRow, 9); s.maxCol = Math.max(s.maxCol, 9);
  recalc(wb);
  return getCell(s, 9, 9)?.value ?? "";
}
const num = (f: string): number => Number(ev(f));

describe("statistics", () => {
  it("computes the spread functions Excel-compatibly", () => {
    expect(ev("MEDIAN(A1:A5)")).toBe("3");
    expect(ev("MODE(A1:A5)")).toBe("1");
    expect(num("STDEV(A1:A5)")).toBeCloseTo(1.7888543819998, 9); // sample (n-1)
    expect(num("STDEV.P(A1:A5)")).toBeCloseTo(1.6, 9);
    expect(num("VAR(A1:A5)")).toBeCloseTo(3.2, 9);
    expect(num("VAR.P(A1:A5)")).toBeCloseTo(2.56, 9);
  });

  it("orders values for LARGE / SMALL / PERCENTILE / QUARTILE / RANK", () => {
    expect(ev("LARGE(A1:A5,1)")).toBe("5");
    expect(ev("LARGE(A1:A5,2)")).toBe("4");
    expect(ev("SMALL(A1:A5,1)")).toBe("1");
    expect(ev("PERCENTILE(A1:A5,0.5)")).toBe("3"); // = the median
    expect(ev("QUARTILE(A1:A5,1)")).toBe("1");
    expect(ev("RANK(4,A1:A5)")).toBe("2"); // only 5 is larger
    expect(ev("RANK(4,A1:A5,1)")).toBe("4"); // ascending
    expect(ev("COUNTBLANK(A1:A5)")).toBe("0");
  });

  it("SUBTOTAL and AGGREGATE select the aggregate by function number", () => {
    expect(ev("SUBTOTAL(9,A1:A5)")).toBe("14");  // SUM
    expect(ev("SUBTOTAL(1,A1:A5)")).toBe("2.8"); // AVERAGE
    expect(ev("SUBTOTAL(4,A1:A5)")).toBe("5");   // MAX
    expect(ev("SUBTOTAL(109,A1:A5)")).toBe("14"); // the 10x form
    expect(ev("AGGREGATE(9,0,A1:A5)")).toBe("14");
    expect(ev("AGGREGATE(14,0,A1:A5,1)")).toBe("5"); // LARGE
  });
});

describe("multi-criteria aggregates", () => {
  it("sums / counts / averages over one and several criteria", () => {
    expect(ev('SUMIFS(A1:A5,A1:A5,">1")')).toBe("12");
    expect(ev('COUNTIFS(A1:A5,">1")')).toBe("3");
    expect(ev('AVERAGEIFS(A1:A5,A1:A5,">1")')).toBe("4");
    expect(ev('MAXIFS(A1:A5,A1:A5,">1")')).toBe("5");
    expect(ev('MINIFS(A1:A5,A1:A5,">1")')).toBe("3");
    // two criteria: rows where B is "b" AND A < 4  -> the two 1s
    expect(ev('SUMIFS(A1:A5,B1:B5,"b",A1:A5,"<4")')).toBe("2");
    expect(ev('COUNTIFS(B1:B5,"a")')).toBe("2");
  });

  it("supports wildcards and <> in criteria", () => {
    expect(ev('COUNTIFS(B1:B5,"<>a")')).toBe("3");
    expect(ev('COUNTIFS(B1:B5,"?")')).toBe("5"); // every value is one char
  });
});

describe("lookup", () => {
  it("MATCH exact / approximate, and the INDEX+MATCH idiom", () => {
    expect(ev("MATCH(4,A1:A5,0)")).toBe("3");
    expect(ev("INDEX(B1:B5,MATCH(4,A1:A5,0))")).toBe("c");
    expect(ev("MATCH(99,A1:A5,0)")).toBe("#N/A");
    expect(ev("MATCH(3,{1,2,3,4},1)")).toBe("3"); // ascending, largest <= 3
  });

  it("XLOOKUP / XMATCH, with the not-found fallback", () => {
    expect(ev("XLOOKUP(4,A1:A5,B1:B5)")).toBe("c");
    expect(ev('XLOOKUP(99,A1:A5,B1:B5,"none")')).toBe("none");
    expect(ev("XMATCH(4,A1:A5)")).toBe("3");
  });

  it("CHOOSE picks the nth argument (its args arrive context-shifted)", () => {
    expect(ev('CHOOSE(2,"a","b","c")')).toBe("b");
    expect(ev("CHOOSE(1,10,20)")).toBe("10");
    expect(ev('CHOOSE(9,"a")')).toBe("#VALUE!");
  });

  it("LOOKUP finds the last value <= the key", () => {
    expect(ev("LOOKUP(3,{1,2,3,4},{10,20,30,40})")).toBe("30");
    expect(ev("LOOKUP(2.5,{1,2,3,4},{10,20,30,40})")).toBe("20");
  });
});

describe("text", () => {
  it("UPPER / SUBSTITUTE / TEXTJOIN / VALUE", () => {
    expect(ev('UPPER("abc")')).toBe("ABC");
    expect(ev('SUBSTITUTE("a-a-a","a","b")')).toBe("b-b-b");
    expect(ev('SUBSTITUTE("a-a-a","a","b",2)')).toBe("a-b-a"); // only the 2nd instance
    expect(ev('TEXTJOIN(",",TRUE,B1:B5)')).toBe("a,b,c,b,a");
    expect(ev('VALUE("1.5")')).toBe("1.5");
    expect(ev('VALUE("50%")')).toBe("0.5");
  });

  it("SEARCH is case-insensitive (unlike FIND) and honours wildcards", () => {
    expect(ev('SEARCH("B","abc")')).toBe("2");
    expect(ev('FIND("B","abc")')).toBe("#VALUE!"); // case-sensitive: not found
    expect(ev('SEARCH("b?","abc")')).toBe("2");
    expect(ev('SEARCH("z","abc")')).toBe("#VALUE!");
  });

  it("TEXTBEFORE / TEXTAFTER split around a delimiter", () => {
    expect(ev('TEXTBEFORE("a-b-c","-")')).toBe("a");
    expect(ev('TEXTAFTER("a-b-c","-")')).toBe("b-c");
    expect(ev('TEXTBEFORE("a-b-c","-",2)')).toBe("a-b");
    expect(ev('TEXTAFTER("a-b-c","-",-1)')).toBe("c"); // last instance
  });
});

describe("logical", () => {
  it("SWITCH matches a value or falls back to the default", () => {
    expect(ev('SWITCH(2,1,"one",2,"two","other")')).toBe("two");
    expect(ev('SWITCH(9,1,"one",2,"two","other")')).toBe("other");
    expect(ev('SWITCH(9,1,"one")')).toBe("#N/A");
  });
});

describe("financial", () => {
  it("annuities match Excel to 6 decimals", () => {
    expect(num("PMT(0.01,12,1000)")).toBeCloseTo(-88.848789, 6);
    expect(num("FV(0.01,12,-100)")).toBeCloseTo(1268.250301, 6);
    expect(num("PV(0.01,12,-100)")).toBeCloseTo(1125.507747, 6);
    expect(num("NPER(0.01,-100,1000)")).toBeCloseTo(10.588644, 5);
    expect(num("RATE(12,-100,1000)")).toBeCloseTo(0.029229, 5);
    expect(num("IPMT(0.01,1,12,1000)")).toBeCloseTo(-10, 9);
    expect(num("PPMT(0.01,1,12,1000)")).toBeCloseTo(-78.848789, 6);
    expect(num("PMT(0,12,1200)")).toBeCloseTo(-100, 9); // zero-rate branch
  });

  it("cash flows: NPV / IRR / XNPV", () => {
    expect(num("NPV(0.1,A1:A5)")).toBeCloseTo(10.346598, 6);
    // -100 then three inflows of 40: NPV(r)=0 at r ~ 9.701%
    expect(num("IRR({-100,40,40,40})")).toBeCloseTo(0.09701, 5);
    expect(ev("IRR(A1:A5)")).toBe("#NUM!"); // no sign change -> no solution
    expect(num("XNPV(0.1,A1:A5,C1:C5)")).toBeCloseTo(13.991648, 5);
  });

  it("depreciation: SLN / SYD / DDB / DB", () => {
    expect(ev("SLN(1000,100,5)")).toBe("180");
    expect(ev("SYD(1000,100,5,1)")).toBe("300");
    expect(ev("DDB(1000,100,5,1)")).toBe("400");
    expect(num("DB(1000,100,5,1)")).toBeCloseTo(369, 6);
  });
});

describe("modern array functions spill", () => {
  it("TAKE / DROP slice from either end", () => {
    const wb = readCsv(DATA);
    const s = wb.sheets[0]!;
    s.cells.set("1:5", { row: 1, col: 5, value: "", kind: "blank", formula: "TAKE(A1:A5,2)" } as never);
    s.maxRow = 5; s.maxCol = 5;
    recalc(wb);
    expect(getCell(s, 1, 5)?.value).toBe("3");
    expect(getCell(s, 2, 5)?.value).toBe("1");
    expect(getCell(s, 3, 5)?.value ?? "").toBe(""); // only 2 rows spilled
  });

  it("VSTACK stacks vertically and HSTACK horizontally", () => {
    const wb = readCsv("1,2\n3,4\n");
    const s = wb.sheets[0]!;
    s.cells.set("1:5", { row: 1, col: 5, value: "", kind: "blank", formula: "VSTACK(A1:B1,A2:B2)" } as never);
    s.maxRow = 5; s.maxCol = 6;
    recalc(wb);
    expect(getCell(s, 1, 5)?.value).toBe("1");
    expect(getCell(s, 1, 6)?.value).toBe("2");
    expect(getCell(s, 2, 5)?.value).toBe("3");
    expect(getCell(s, 2, 6)?.value).toBe("4");
  });

  it("SORTBY orders rows by a key column", () => {
    const wb = readCsv(DATA);
    const s = wb.sheets[0]!;
    s.cells.set("1:5", { row: 1, col: 5, value: "", kind: "blank", formula: "SORTBY(A1:A5,B1:B5)" } as never);
    s.maxRow = 5; s.maxCol = 5;
    recalc(wb);
    // keys a,b,c,b,a -> rows for a (3,5), then b (1,1), then c (4)
    expect([1, 2, 3, 4, 5].map((r) => getCell(s, r, 5)?.value)).toEqual(["3", "5", "1", "1", "4"]);
  });

  it("TOCOL flattens and WRAPROWS reshapes", () => {
    const wb = readCsv("1,2\n3,4\n");
    const s = wb.sheets[0]!;
    s.cells.set("1:5", { row: 1, col: 5, value: "", kind: "blank", formula: "TOCOL(A1:B2)" } as never);
    s.maxRow = 6; s.maxCol = 6;
    recalc(wb);
    expect([1, 2, 3, 4].map((r) => getCell(s, r, 5)?.value)).toEqual(["1", "2", "3", "4"]);
  });
});
