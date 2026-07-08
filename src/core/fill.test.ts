import { describe, expect, it } from "vitest";
import { computeFill, seriesStep, type FillSource } from "./fill";

const num = (v: number): FillSource => ({ value: String(v), kind: "n" });
const txt = (v: string): FillSource => ({ value: v, kind: "s" });
const fx = (f: string): FillSource => ({ value: "", formula: f, kind: "blank" });

describe("seriesStep", () => {
  it("detects linear series and rejects everything else", () => {
    expect(seriesStep([num(1), num(2)])).toBe(1);
    expect(seriesStep([num(10), num(20), num(30)])).toBe(10);
    expect(seriesStep([num(5), num(3)])).toBe(-2);
    expect(seriesStep([num(1)])).toBeNull();
    expect(seriesStep([num(1), num(2), num(4)])).toBeNull();
    expect(seriesStep([num(1), txt("x")])).toBeNull();
  });
});

describe("computeFill", () => {
  it("extends a numeric series downward and upward", () => {
    expect(computeFill([num(1), num(2)], 4, 1, "row")).toEqual(["3", "4", "5", "6"]);
    expect(computeFill([num(3), num(5)], 3, -1, "row")).toEqual(["1", "-1", "-3"]);
  });

  it("copies a single number as-is", () => {
    expect(computeFill([num(7)], 3, 1, "row")).toEqual(["7", "7", "7"]);
  });

  it("increments a single trailing-integer text", () => {
    expect(computeFill([txt("item1")], 3, 1, "col")).toEqual(["item2", "item3", "item4"]);
    expect(computeFill([txt("item5")], 2, -1, "col")).toEqual(["item4", "item3"]);
  });

  it("repeats a mixed pattern cyclically", () => {
    expect(computeFill([txt("a"), txt("b")], 5, 1, "row")).toEqual(["a", "b", "a", "b", "a"]);
  });

  it("shifts formula references relative to the distance moved", () => {
    expect(computeFill([fx("A1*2")], 3, 1, "row")).toEqual(["=A2*2", "=A3*2", "=A4*2"]);
    expect(computeFill([fx("SUM(A1:B1)")], 2, 1, "col")).toEqual(["=SUM(B1:C1)", "=SUM(C1:D1)"]);
    expect(computeFill([fx("$A$1+B2")], 2, 1, "row")).toEqual(["=$A$1+B3", "=$A$1+B4"]);
  });

  it("keeps a two-formula pattern aligned", () => {
    const out = computeFill([fx("A1"), fx("B1")], 4, 1, "row");
    expect(out).toEqual(["=A3", "=B3", "=A5", "=B5"]);
  });
});
