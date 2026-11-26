import { afterEach, describe, expect, it } from "vitest";
import FormulaParser from "fast-formula-parser";
import { setLocale } from "./i18n";
import { isInvariantText, localDateEditText, localDisplay, localizeNumberText, parseLocalNumber, regional, setRegionalLocale } from "./regional";
import { formulaFromLocal, formulaToLocal, inputFromLocal } from "./formula-locale";
import { FUNCTION_NAMES } from "./fn-names";
import { extraFunctions } from "./functions";
import { financialFunctions } from "./financial";
import { referenceFunctions } from "./reference-fns";
import type { Cell } from "./model";

afterEach(() => {
  setRegionalLocale(null);
  setLocale("en");
});

const fr = () => {
  setRegionalLocale("fr");
  return regional();
};

describe("formulas in French", () => {
  it("shows SUM with SOMME, ; and a decimal comma", () => {
    const r = fr();
    expect(formulaToLocal("SUM(A1:A3,1.5)", r)).toBe("SOMME(A1:A3;1,5)");
    expect(formulaToLocal("IF(AND(A1>0,B1<>\"\"),VLOOKUP(A1,Sheet2!A:B,2,FALSE),TRUE)", r)).toBe(
      "SI(ET(A1>0;B1<>\"\");RECHERCHEV(A1;Sheet2!A:B;2;FAUX);VRAI)",
    );
  });

  it("stores what was typed in French as the invariant formula", () => {
    const r = fr();
    expect(formulaFromLocal("SOMME(A1:A3;1,5)", r)).toBe("SUM(A1:A3,1.5)");
    expect(formulaFromLocal("nb.si(A:A;\">0,5\")", r)).toBe("COUNTIF(A:A,\">0,5\")");
    expect(formulaFromLocal("MOYENNE(A1;A2)*2,5%", r)).toBe("AVERAGE(A1,A2)*2.5%");
  });

  it("round-trips every French name", () => {
    const r = fr();
    for (const en of Object.keys(FUNCTION_NAMES.fr!)) {
      const f = `${en}(A1,2.5)`;
      expect(formulaFromLocal(formulaToLocal(f, r), r)).toBe(f);
    }
  });

  it("leaves strings, quoted sheets, structured refs and names alone", () => {
    const r = fr();
    const f = `'Q1, 2026'!A1&"a,b;1.5"&SUM(Table1[[#Headers],[Amount, net]])+TRUE_RATE`;
    const local = formulaToLocal(f, r);
    expect(local).toBe(`'Q1, 2026'!A1&"a,b;1.5"&SOMME(Table1[[#Headers];[Amount, net]])+TRUE_RATE`);
    expect(formulaFromLocal(local, r)).toBe(f);
  });

  it("converts array constants, errors and future-function prefixes", () => {
    const r = fr();
    expect(formulaToLocal("SUM({1.5,2;3,4})", r)).toBe("SOMME({1,5.2;3.4})");
    expect(formulaFromLocal("SOMME({1,5.2;3.4})", r)).toBe("SUM({1.5,2;3,4})");
    expect(formulaToLocal("IFERROR(A1,#VALUE!)", r)).toBe("SIERREUR(A1;#VALEUR!)");
    expect(formulaFromLocal("SIERREUR(A1;#VALEUR!)", r)).toBe("IFERROR(A1,#VALUE!)");
    expect(formulaToLocal("_xlfn.XLOOKUP(A1,B:B,C:C)", r)).toBe("_xlfn.RECHERCHEX(A1;B:B;C:C)");
    expect(formulaFromLocal("_xlfn.RECHERCHEX(A1;B:B;C:C)", r)).toBe("_xlfn.XLOOKUP(A1,B:B,C:C)");
  });

  it("still accepts English names typed in French, as long as no French name means otherwise", () => {
    const r = fr();
    expect(formulaFromLocal("SUM(A1;A2)", r)).toBe("SUM(A1,A2)");
    // TRIM is the French name of MIRR, as in French Excel.
    expect(formulaFromLocal("TRIM(A1:A3;0,1;0,2)", r)).toBe("MIRR(A1:A3,0.1,0.2)");
    expect(formulaToLocal("TRIM(A1)", r)).toBe("SUPPRESPACE(A1)");
  });

  it("evaluates the stored form with the engine", () => {
    const parser = new FormulaParser({ onCell: () => 2 });
    expect(parser.parse(formulaFromLocal("SOMME(1,5;2,5)*ARRONDI(A1;0)", fr()), { sheet: "S", row: 1, col: 1 })).toBe(8);
  });

  it("is a no-op in English", () => {
    setRegionalLocale("en");
    expect(formulaToLocal("SUM(A1,1.5)")).toBe("SUM(A1,1.5)");
    expect(formulaFromLocal("SUM(A1,1.5)")).toBe("SUM(A1,1.5)");
  });
});

describe("formulas in German and Spanish", () => {
  it("round-trips every German and Spanish name", () => {
    for (const code of ["de", "es"]) {
      setRegionalLocale(code);
      for (const en of Object.keys(FUNCTION_NAMES[code]!)) {
        const f = `${en}(A1,2.5)`;
        expect(formulaFromLocal(formulaToLocal(f))).toBe(f);
      }
    }
  });

  it("translates names and separators", () => {
    setRegionalLocale("de");
    expect(formulaToLocal("IF(SUM(A1:A2)>1.5,TRUE,FALSE)")).toBe("WENN(SUMME(A1:A2)>1,5;WAHR;FALSCH)");
    expect(formulaFromLocal("ZÄHLENWENN(A:A;\"x\")")).toBe("COUNTIF(A:A,\"x\")");
    setRegionalLocale("es");
    expect(formulaToLocal("YEAR(TODAY())")).toBe("AÑO(HOY())");
    expect(formulaFromLocal("si(a1>0;suma(b1;b2);0)")).toBe("IF(a1>0,SUM(b1,b2),0)");
  });
});

describe("function name tables", () => {
  it("cover the functions the engine implements, in French", () => {
    const p = new FormulaParser({}) as unknown as { functions: Record<string, unknown> };
    const implemented = new Set([
      ...Object.keys(p.functions),
      ...Object.keys(extraFunctions()),
      ...Object.keys(financialFunctions()),
      ...Object.keys(referenceFunctions()),
    ]);
    // Names French Excel keeps as in English; every other implemented function needs a French name.
    const same = new Set(`ABS ACOS ACOSH ACOT ACOTH ASC ASIN ASINH ATAN ATAN2 ATANH AVERAGEA BAHTTEXT BASE
      BESSELI BESSELJ BESSELK BESSELY CHISQ.TEST CODE COMBIN COMBINA CONCAT COS COSH COT COTH CSC CSCH DATE
      DATEDIF DB DBCS DDB DECIMAL DELTA ERF ERFC EXACT EXP F.TEST FACT FACTDOUBLE FISHER GAMMA GAUSS
      INDEX INDIRECT ISO.CEILING LENB LN LOG LOG10 MAX MAXA MIN MINA MINUTE MIRR MOD MODE N NA PEARSON
      PERMUTATIONA PHI PI QUARTILE QUOTIENT RADIANS REPT SEC SECH SEQUENCE SIN SINH STDEVA STDEVPA SYD T
      T.TEST TAN TANH TRANSPOSE TYPE UNICODE UNIQUE VAR VAR.S VARA VARPA Z.TEST`.split(/\s+/));
    const missing = [...implemented].filter((n) => !FUNCTION_NAMES.fr![n] && !same.has(n));
    expect(missing).toEqual([]);
  });

  it("cover the functions the engine implements, in Italian, Polish and Turkish", () => {
    const p = new FormulaParser({}) as unknown as { functions: Record<string, unknown> };
    const implemented = new Set([
      ...Object.keys(p.functions),
      ...Object.keys(extraFunctions()),
      ...Object.keys(financialFunctions()),
      ...Object.keys(referenceFunctions()),
    ]);
    // Names each Excel keeps as in English (Microsoft's articles); every other one needs a local name.
    const same: Record<string, string> = {
    it: `ASC BASE BITAND BITOR BITXOR CONCAT COS COSH COT COTH CSC CSCH DBCS DELTA EXP FISHER GAMMA GAUSS
      INT LN LOG LOG10 MAX MIN PEARSON PERCENTILE PHI QUARTILE SEC SECH SWITCH T TAN TANH UNICODE VAR
      VAR.P XOR`,
    pl: `ACOS ACOSH ACOT ACOTH ASC ASIN ASINH ATAN ATAN2 ATANH BITAND BITOR BITXOR COS COSH COT COTH CSC
      CSCH DB DBCS DDB ENCODEURL EXP F.TEST FILTERXML FV GAMMA GAUSS INFO IPMT IRR LN LOG LOG10 MAX
      MIN MIRR MOD N NPER NPV PEARSON PHI PHONETIC PI PMT PPMT PV RATE SEC SECH SIN SINH SLN SYD T
      T.TEST TAN TANH UNICODE WEBSERVICE XIRR XNPV XOR Z.TEST`,
    tr: `ACOS ACOSH ACOT ACOTH ASC ATAN ATAN2 ATANH BESSELI BESSELJ BESSELK BESSELY BIN2DEC BIN2HEX
      BIN2OCT COS COSH COT COTH CSC CSCH DBCS DEC2BIN DEC2HEX DEC2OCT DELTA F.TEST FISHER GAUSS
      HEX2BIN HEX2DEC HEX2OCT LN LOG LOG10 MOD OCT2BIN OCT2DEC OCT2HEX PEARSON PHI RANK SEC SECH
      T.TEST TAN TANH UNICODE VAR VAR.P VAR.S VARA Z.TEST`,
    };
    for (const [code, list] of Object.entries(same)) {
      const keep = new Set(list.split(/\s+/));
      const missing = [...implemented].filter((n) => !FUNCTION_NAMES[code]![n] && !keep.has(n));
      expect(missing, code).toEqual([]);
    }
  });

  it("match Microsoft's names for the tricky ones", () => {
    const pick = (code: string, names: string[]) => names.map((n) => FUNCTION_NAMES[code]![n]);
    expect(pick("fr", ["MIRR", "TRIM", "NORM.DIST", "MAXIFS", "DOLLAR", "FINDB"])).toEqual(
      ["TRIM", "SUPPRESPACE", "LOI.NORMALE.N", "MAX.SI.ENS", "DEVISE", "TROUVERB"]);
    expect(pick("de", ["ACOS", "SWITCH", "COUNTBLANK", "T.DIST.2T", "LENB", "Z.TEST"])).toEqual(
      ["ARCCOS", "ERSTERWERT", "ANZAHLLEEREZELLEN", "T.VERT.2S", "LÄNGEB", "G.TEST"]);
    expect(pick("es", ["AND", "FIXED", "TIME", "NORM.S.DIST", "T.TEST", "BITOR"])).toEqual(
      ["Y", "DECIMAL", "NSHORA", "DISTR.NORM.ESTAND.N", "PRUEBA.T.N", "BIT.O"]);
    expect(pick("it", ["IFNA", "SUMPRODUCT", "COLUMN", "REPLACE", "PV", "LENB"])).toEqual(
      ["SE.NON.DISP.", "MATR.SOMMA.PRODOTTO", "RIF.COLONNA", "RIMPIAZZA", "VA", "LUNGB"]);
    expect(pick("pl", ["ABS", "AND", "COUNTIF", "GROWTH", "ROWS", "ISNA"])).toEqual(
      ["MODUŁ.LICZBY", "ORAZ", "LICZ.JEŻELI", "REGEXPW", "ILE.WIERSZY", "CZY.BRAK"]);
    expect(pick("tr", ["OR", "SWITCH", "TIMEVALUE", "MID", "MIDB", "XLOOKUP"])).toEqual(
      ["YADA", "İLKEŞLEŞEN", "ZAMANSAYISI", "PARÇAAL", "ORTAB", "ÇAPRAZARA"]);
  });

  it("never give two functions the same local name", () => {
    for (const [code, map] of Object.entries(FUNCTION_NAMES)) {
      const seen = new Map<string, string>();
      for (const [en, local] of Object.entries(map)) {
        const prev = seen.get(local.toUpperCase());
        expect(prev, `${code}: ${local} is both ${prev} and ${en}`).toBeUndefined();
        seen.set(local.toUpperCase(), en);
      }
    }
  });
});

describe("numbers in French", () => {
  it("reads typed numbers with a decimal comma and space groups", () => {
    const r = fr();
    expect(parseLocalNumber("1234,5", r)).toBe("1234.5");
    expect(parseLocalNumber("1 234,56", r)).toBe("1234.56");
    expect(parseLocalNumber("-1 234 567", r)).toBe("-1234567");
    expect(parseLocalNumber(",5", r)).toBe("0.5");
    expect(parseLocalNumber("1,5E3", r)).toBe("1.5E3");
    expect(parseLocalNumber("12 34", r)).toBeNull();
    expect(parseLocalNumber("abc", r)).toBeNull();
    setRegionalLocale("de");
    expect(parseLocalNumber("1.234,5")).toBe("1234.5");
  });

  it("makes typed input invariant", () => {
    const r = fr();
    expect(inputFromLocal("=SOMME(A1;0,5)", r)).toBe("=SUM(A1,0.5)");
    expect(inputFromLocal("3,25", r)).toBe("3.25");
    expect(inputFromLocal("vrai", r)).toBe("TRUE");
    expect(inputFromLocal("bonjour, monde", r)).toBe("bonjour, monde");
  });

  it("swaps the separators of a formatted number", () => {
    const r = fr();
    expect(localizeNumberText("1,234,567.89", r)).toBe("1 234 567,89");
    expect(localizeNumberText("$1,234.50", r)).toBe("$1 234,50");
    expect(localizeNumberText("12.5%", r)).toBe("12,5%");
    expect(localizeNumberText("1.5E+10", r)).toBe("1,5E+10");
  });

  it("shows cells the French way and leaves producer text alone", () => {
    const r = fr();
    const cell = (c: Partial<Cell>): Cell => ({ row: 1, col: 1, value: "", kind: "n", ...c });
    expect(localDisplay(cell({ value: "1234.5" }), r)).toBe("1234,5");
    expect(localDisplay(cell({ value: "1234.5", numFmt: "#,##0.00", display: "1,234.50" }), r)).toBe("1 234,50");
    expect(localDisplay(cell({ value: "46351", numFmt: 14, display: "11/25/26" }), r)).toBe("25/11/2026");
    expect(localDisplay(cell({ value: "46351", numFmt: "d mmmm yyyy", display: "25 November 2026" }), r)).toBe("25 novembre 2026");
    expect(localDisplay(cell({ value: "1234.5", display: "1 234,50", displayRaw: true }), r)).toBe("1 234,50");
    expect(localDisplay(cell({ value: "TRUE", kind: "b" }), r)).toBe("VRAI");
    expect(localDisplay(cell({ value: "#VALUE!", kind: "e" }), r)).toBe("#VALEUR!");
    expect(localDisplay(cell({ value: "a.b, c", kind: "s" }), r)).toBe("a.b, c");
    expect(localDateEditText(46351.5, true, r)).toBe("25/11/2026 12:00");
  });

  it("tells an English-style producer text from a local one", () => {
    expect(isInvariantText("1,234.50", "1234.5")).toBe(true);
    expect(isInvariantText("$1,234.57", "1234.5678")).toBe(true);
    expect(isInvariantText("12.50%", "0.125")).toBe(true);
    expect(isInvariantText("1 234,56", "1234.56")).toBe(false);
    expect(isInvariantText("1,234", "1.234")).toBe(false);
  });

  it("follows the UI locale unless forced", () => {
    setLocale("fr");
    expect(regional().decimal).toBe(",");
    setRegionalLocale("en");
    expect(regional().decimal).toBe(".");
  });
});
