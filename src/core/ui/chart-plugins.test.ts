import { describe, expect, it } from "vitest";
import { fitTrendline } from "./chart-plugins";

describe("fitTrendline", () => {
  it("fits a perfect line (r2 = 1, correct slope/intercept)", () => {
    const pts = [0, 1, 2, 3, 4].map((x) => ({ x, y: 2 * x + 1 }));
    const fit = fitTrendline({ type: "linear" }, pts)!;
    expect(fit.f!(10)).toBeCloseTo(21, 6);
    expect(fit.r2).toBeCloseTo(1, 6);
  });

  it("fits an exponential curve", () => {
    const pts = [0, 1, 2, 3].map((x) => ({ x, y: 3 * Math.exp(0.5 * x) }));
    const fit = fitTrendline({ type: "exp" }, pts)!;
    expect(fit.f!(4)).toBeCloseTo(3 * Math.exp(2), 4);
    expect(fit.r2).toBeCloseTo(1, 6);
  });

  it("fits a quadratic exactly with poly order 2", () => {
    const pts = [-2, -1, 0, 1, 2].map((x) => ({ x, y: x * x - x + 2 }));
    const fit = fitTrendline({ type: "poly", order: 2 }, pts)!;
    expect(fit.f!(3)).toBeCloseTo(9 - 3 + 2, 4);
    expect(fit.r2).toBeCloseTo(1, 6);
  });

  it("produces moving-average points offset by the period", () => {
    const pts = [10, 20, 30, 40].map((y, x) => ({ x, y }));
    const fit = fitTrendline({ type: "movingAvg", order: 2 }, pts)!;
    // first avg is over points 0..1 = 15, at x = 1
    expect(fit.points![0]).toEqual({ x: 1, y: 15 });
    expect(fit.points![2]).toEqual({ x: 3, y: 35 });
  });
});
