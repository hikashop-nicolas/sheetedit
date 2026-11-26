import { afterEach, describe, expect, it } from "vitest";
import { setLocale } from "./i18n";
import { localDateEditText, localDisplay, parseLocalNumber, regional, setRegionalLocale } from "./regional";
import { formulaFromLocal, formulaToLocal, inputFromLocal } from "./formula-locale";
import { FUNCTION_NAMES } from "./fn-names";
import type { Cell } from "./model";

afterEach(() => {
  setRegionalLocale(null);
  setLocale("en");
});

const cell = (c: Partial<Cell>): Cell => ({ row: 1, col: 1, value: "", kind: "n", ...c });

describe("Italian, Polish and Turkish formulas", () => {
  const cases: [string, string, string][] = [
    ["it", "SUM(A1,0.5)", "SOMMA(A1;0,5)"],
    ["it", "IF(VLOOKUP(A1,B:C,2,FALSE)>1.5,TRUE,#N/A)", "SE(CERCA.VERT(A1;B:C;2;FALSO)>1,5;VERO;#N/D)"],
    ["it", "IFNA(A1,0)", "SE.NON.DISP.(A1;0)"],
    ["pl", "SUM(A1,0.5)", "SUMA(A1;0,5)"],
    ["pl", "IF(VLOOKUP(A1,B:C,2,FALSE)>1.5,TRUE,#VALUE!)", "JEŻELI(WYSZUKAJ.PIONOWO(A1;B:C;2;FAŁSZ)>1,5;PRAWDA;#ARG!)"],
    ["tr", "SUM(A1,0.5)", "TOPLA(A1;0,5)"],
    ["tr", "IF(VLOOKUP(A1,B:C,2,FALSE)>1.5,TRUE,#REF!)", "EĞER(DÜŞEYARA(A1;B:C;2;YANLIŞ)>1,5;DOĞRU;#BAŞV!)"],
  ];

  it("translate names, separators, booleans and errors both ways", () => {
    for (const [code, inv, local] of cases) {
      setRegionalLocale(code);
      expect(formulaToLocal(inv)).toBe(local);
      expect(formulaFromLocal(local)).toBe(inv);
    }
  });

  it("round-trip every local name", () => {
    for (const code of ["it", "pl", "tr"]) {
      setRegionalLocale(code);
      for (const en of Object.keys(FUNCTION_NAMES[code]!)) {
        const f = `${en}(A1,2.5)`;
        expect(formulaFromLocal(formulaToLocal(f)), `${code} ${en}`).toBe(f);
      }
    }
  });

  it("read lower-case local names and numbers", () => {
    setRegionalLocale("tr");
    expect(inputFromLocal("=topla(a1;0,5)")).toBe("=SUM(a1,0.5)");
    expect(inputFromLocal("1.234,5")).toBe("1234.5");
    expect(inputFromLocal("doğru")).toBe("TRUE");
    expect(inputFromLocal("=tarihsayısı(A1)+eğersay(B:B;1)")).toBe("=DATEVALUE(A1)+COUNTIF(B:B,1)");
    setRegionalLocale("pl");
    expect(inputFromLocal("1 234,5")).toBe("1234.5");
    expect(parseLocalNumber("1 234,5")).toBe("1234.5");
    setRegionalLocale("it");
    expect(inputFromLocal("1.234,5")).toBe("1234.5");
    expect(inputFromLocal("=somma(A1:A3)*2,5%")).toBe("=SUM(A1:A3)*2.5%");
  });
});

describe("regional display", () => {
  it("shows dates, booleans and errors the local way", () => {
    const date = cell({ value: "46351", numFmt: 14, display: "11/25/26" });
    const long = cell({ value: "46351", numFmt: "d mmmm yyyy", display: "25 November 2026" });
    const want: Record<string, [string, string, string, string, string]> = {
      it: ["25/11/2026", "25 novembre 2026", "FALSO", "#NOME?", "1.234,5"],
      pl: ["25.11.2026", "25 listopad 2026", "FAŁSZ", "#NAZWA?", "1 234,5"],
      tr: ["25.11.2026", "25 Kasım 2026", "YANLIŞ", "#AD?", "1.234,5"],
    };
    for (const [code, [d, l, b, e, n]] of Object.entries(want)) {
      setRegionalLocale(code);
      expect(localDisplay(date)).toBe(d);
      expect(localDisplay(long)).toBe(l);
      expect(localDisplay(cell({ value: "FALSE", kind: "b" }))).toBe(b);
      expect(localDisplay(cell({ value: "#NAME?", kind: "e" }))).toBe(e);
      expect(localDisplay(cell({ value: "1234.5", numFmt: "#,##0.0", display: "1,234.5" }))).toBe(n);
      expect(localDateEditText(46351, false)).toBe(d);
    }
    // German had the same dotted short date, which SSF misread before.
    setRegionalLocale("de");
    expect(localDisplay(date)).toBe("25.11.2026");
  });

  it("keeps English formulas in Korean, with an ISO short date", () => {
    setLocale("ko");
    const r = regional();
    expect([r.decimal, r.argSep, r.bool]).toEqual([".", ",", undefined]);
    expect(formulaToLocal("SUM(A1,0.5)")).toBe("SUM(A1,0.5)");
    expect(inputFromLocal("TRUE")).toBe("TRUE");
    expect(localDisplay(cell({ value: "46351", numFmt: 14, display: "11/25/26" }))).toBe("2026-11-25");
    expect(localDisplay(cell({ value: "1234.5", numFmt: "#,##0.0", display: "1,234.5" }))).toBe("1,234.5");
    expect(localDateEditText(46351, false)).toBeNull();
  });
});
