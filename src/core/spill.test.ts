import { describe, expect, it } from "vitest";
import { cannotOverflow, fontPx } from "./spill";

// Measuring a cell's text forces a layout, and the grid did it for every text cell it drew.
// cannotOverflow rules out the cells whose text cannot possibly reach the next column.
describe("ruling out a spill without measuring", () => {
  it("skips a short label in an ordinary column", () => {
    expect(cannotOverflow("Rent", fontPx(undefined), 96)).toBe(true);
  });

  it("still measures a label that might overflow", () => {
    expect(cannotOverflow("Studio budget 2027", fontPx(undefined), 96)).toBe(false);
  });

  it("stays safe for full-width text, which is about 1em a character", () => {
    // Six CJK characters at 13px are about 78px wide: close enough to 96px that it must be measured.
    expect(cannotOverflow("契約書に記入", fontPx(undefined), 96)).toBe(false);
    expect(cannotOverflow("契約", fontPx(undefined), 96)).toBe(true);
  });

  it("counts a character outside the basic plane once", () => {
    expect(cannotOverflow("😀😀", fontPx(undefined), 60)).toBe(true);
  });

  it("uses the cell's own font size", () => {
    // "Total" at 8pt is at most 16 + 5 x 10.7px x 1.35 = 88px: fits a 96px column. At 24pt it may not.
    expect(cannotOverflow("Total", fontPx(8), 96)).toBe(true);
    expect(cannotOverflow("Total", fontPx(24), 96)).toBe(false);
  });
});
