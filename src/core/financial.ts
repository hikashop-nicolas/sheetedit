import FormulaParser from "fast-formula-parser";
import { asMatrix, asNumber } from "./dynamic-arrays";

// Financial functions. fast-formula-parser ships none of them, so the whole family is supplied
// here: the annuity set (PMT / FV / PV / NPER / RATE / IPMT / PPMT / CUMIPMT / CUMPRINC), the
// cash-flow set (NPV / IRR / MIRR / XNPV / XIRR) and depreciation (SLN / SYD / DB / DDB).
// Excel's sign convention is followed: money you pay out is negative, money you receive positive.

const NUM = (FormulaParser as unknown as { FormulaError: { NUM: unknown } }).FormulaError.NUM;

const numsOf = (arg: unknown): number[] => asMatrix(arg).flat().filter((v): v is number => typeof v === "number" && Number.isFinite(v));

/** (1+r)^n and the annuity factor ((1+r)^n - 1)/r * (1 + r*type), handling r = 0. */
function factors(r: number, n: number, type: number): { pow: number; ann: number } {
  if (r === 0) return { pow: 1, ann: n };
  const pow = Math.pow(1 + r, n);
  return { pow, ann: ((pow - 1) / r) * (1 + r * type) };
}

function PMT(rate: number, nper: number, pv: number, fv: number, type: number): number {
  const { pow, ann } = factors(rate, nper, type);
  return ann === 0 ? NaN : -(pv * pow + fv) / ann;
}
function FV(rate: number, nper: number, pmt: number, pv: number, type: number): number {
  const { pow, ann } = factors(rate, nper, type);
  return -(pv * pow + pmt * ann);
}
function PV(rate: number, nper: number, pmt: number, fv: number, type: number): number {
  const { pow, ann } = factors(rate, nper, type);
  return pow === 0 ? NaN : -(fv + pmt * ann) / pow;
}
function NPER(rate: number, pmt: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return pmt === 0 ? NaN : -(pv + fv) / pmt;
  const c = (pmt * (1 + rate * type)) / rate;
  const ratio = (c - fv) / (pv + c);
  return ratio <= 0 ? NaN : Math.log(ratio) / Math.log(1 + rate);
}
/** The remaining balance after `per` periods, used by IPMT. */
function balanceAfter(rate: number, per: number, pmt: number, pv: number, type: number): number {
  return FV(rate, per, pmt, pv, type);
}
function IPMT(rate: number, per: number, nper: number, pv: number, fv: number, type: number): number {
  if (per < 1 || per > nper) return NaN;
  const pmt = PMT(rate, nper, pv, fv, type);
  if (per === 1) return type === 1 ? 0 : -pv * rate;
  const prev = balanceAfter(rate, per - 1, pmt, pv, type);
  // With payments at the start of the period, the interest accrues on the post-payment balance.
  return type === 1 ? (prev - pmt) * rate : prev * rate;
}

/** Solve f(x) = 0 by Newton with a bisection fallback; used by RATE / IRR / XIRR. */
function solve(f: (x: number) => number, guess: number): number {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (!Number.isFinite(y)) break;
    if (Math.abs(y) < 1e-10) return x;
    const h = Math.abs(x) > 1e-7 ? x * 1e-6 : 1e-7;
    const d = (f(x + h) - y) / h;
    if (!Number.isFinite(d) || d === 0) break;
    const next = x - y / d;
    if (!Number.isFinite(next)) break;
    if (Math.abs(next - x) < 1e-12) return next;
    x = next;
  }
  // Bisection over a wide bracket when Newton wanders off.
  let lo = -0.9999, hi = 10;
  let flo = f(lo);
  if (!Number.isFinite(flo)) return NaN;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    const fm = f(mid);
    if (!Number.isFinite(fm)) return NaN;
    if (Math.abs(fm) < 1e-10) return mid;
    if ((flo < 0) === (fm < 0)) { lo = mid; flo = fm; } else hi = mid;
  }
  return NaN;
}

const ok = (v: number): unknown => (Number.isFinite(v) ? v : NUM);

export function financialFunctions(): Record<string, (...args: unknown[]) => unknown> {
  const n = (a: unknown, d = 0): number => asNumber(a, d);
  return {
    PMT: (r, np, pv, fv, t) => ok(PMT(n(r), n(np), n(pv), n(fv), n(t))),
    FV: (r, np, pmt, pv, t) => ok(FV(n(r), n(np), n(pmt), n(pv), n(t))),
    PV: (r, np, pmt, fv, t) => ok(PV(n(r), n(np), n(pmt), n(fv), n(t))),
    NPER: (r, pmt, pv, fv, t) => ok(NPER(n(r), n(pmt), n(pv), n(fv), n(t))),
    IPMT: (r, per, np, pv, fv, t) => ok(IPMT(n(r), n(per), n(np), n(pv), n(fv), n(t))),
    PPMT: (r, per, np, pv, fv, t) => {
      const rate = n(r), nper = n(np), pval = n(pv), fval = n(fv), ty = n(t);
      return ok(PMT(rate, nper, pval, fval, ty) - IPMT(rate, n(per), nper, pval, fval, ty));
    },
    RATE: (np, pmt, pv, fv, t, guess) => {
      const nper = n(np), p = n(pmt), pval = n(pv), fval = n(fv), ty = n(t);
      return ok(solve((r) => (r === 0 ? pval + p * nper + fval : pval * Math.pow(1 + r, nper) + p * (1 + r * ty) * ((Math.pow(1 + r, nper) - 1) / r) + fval), asNumber(guess, 0.1)));
    },
    CUMIPMT: (r, np, pv, start, end, t) => {
      const rate = n(r), nper = n(np), pval = n(pv), ty = n(t);
      let sum = 0;
      for (let p = Math.trunc(n(start)); p <= Math.trunc(n(end)); p++) sum += IPMT(rate, p, nper, pval, 0, ty);
      return ok(sum);
    },
    CUMPRINC: (r, np, pv, start, end, t) => {
      const rate = n(r), nper = n(np), pval = n(pv), ty = n(t);
      const pmt = PMT(rate, nper, pval, 0, ty);
      let sum = 0;
      for (let p = Math.trunc(n(start)); p <= Math.trunc(n(end)); p++) sum += pmt - IPMT(rate, p, nper, pval, 0, ty);
      return ok(sum);
    },
    // Cash flows. NPV discounts from period 1; IRR includes the period-0 outlay.
    NPV: (r, ...vals) => {
      const rate = n(r);
      const cf = vals.flatMap(numsOf);
      let acc = 0;
      cf.forEach((v, i) => { acc += v / Math.pow(1 + rate, i + 1); });
      return ok(acc);
    },
    IRR: (vals, guess) => {
      const cf = numsOf(vals);
      if (cf.length < 2) return NUM;
      return ok(solve((r) => cf.reduce((a, v, i) => a + v / Math.pow(1 + r, i), 0), asNumber(guess, 0.1)));
    },
    MIRR: (vals, financeRate, reinvestRate) => {
      const cf = numsOf(vals);
      const fr = n(financeRate), rr = n(reinvestRate);
      const nn = cf.length - 1;
      if (nn < 1) return NUM;
      let neg = 0, pos = 0;
      cf.forEach((v, i) => { if (v < 0) neg += v / Math.pow(1 + fr, i); else pos += v * Math.pow(1 + rr, nn - i); });
      if (neg === 0) return NUM;
      return ok(Math.pow(-pos / neg, 1 / nn) - 1);
    },
    XNPV: (r, vals, dates) => {
      const rate = n(r), cf = numsOf(vals), ds = numsOf(dates);
      if (cf.length !== ds.length || !cf.length) return NUM;
      const d0 = ds[0]!;
      let acc = 0;
      cf.forEach((v, i) => { acc += v / Math.pow(1 + rate, (ds[i]! - d0) / 365); });
      return ok(acc);
    },
    XIRR: (vals, dates, guess) => {
      const cf = numsOf(vals), ds = numsOf(dates);
      if (cf.length !== ds.length || cf.length < 2) return NUM;
      const d0 = ds[0]!;
      return ok(solve((r) => cf.reduce((a, v, i) => a + v / Math.pow(1 + r, (ds[i]! - d0) / 365), 0), asNumber(guess, 0.1)));
    },
    // Depreciation.
    SLN: (cost, salvage, life) => { const l = n(life); return l === 0 ? NUM : ok((n(cost) - n(salvage)) / l); },
    SYD: (cost, salvage, life, per) => {
      const l = n(life), p = n(per);
      if (l <= 0 || p < 1 || p > l) return NUM;
      return ok(((n(cost) - n(salvage)) * (l - p + 1) * 2) / (l * (l + 1)));
    },
    DDB: (cost, salvage, life, per, factor) => {
      const c = n(cost), s = n(salvage), l = n(life), p = n(per), f = asNumber(factor, 2);
      if (l <= 0 || p < 1 || p > l) return NUM;
      let acc = 0;
      for (let i = 1; i <= p; i++) {
        const dep = Math.min((c - acc) * (f / l), Math.max(0, c - s - acc));
        if (i === p) return ok(dep);
        acc += dep;
      }
      return NUM;
    },
    DB: (cost, salvage, life, per, month) => {
      const c = n(cost), s = n(salvage), l = n(life), p = n(per), m = asNumber(month, 12);
      if (c <= 0 || l <= 0 || p < 1) return NUM;
      const rate = Math.round((1 - Math.pow(s / c, 1 / l)) * 1000) / 1000;
      let acc = 0, dep = 0;
      for (let i = 1; i <= p; i++) {
        dep = i === 1 ? (c * rate * m) / 12 : i === l + 1 ? ((c - acc) * rate * (12 - m)) / 12 : (c - acc) * rate;
        acc += dep;
      }
      return ok(dep);
    },
    EFFECT: (nominal, npery) => { const np = Math.trunc(n(npery)); return np < 1 ? NUM : ok(Math.pow(1 + n(nominal) / np, np) - 1); },
    NOMINAL: (effect, npery) => { const np = Math.trunc(n(npery)); return np < 1 ? NUM : ok((Math.pow(1 + n(effect), 1 / np) - 1) * np); },
  };
}
