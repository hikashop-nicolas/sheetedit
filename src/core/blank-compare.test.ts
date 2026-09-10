import { describe, expect, it } from "vitest";
import { readCsv } from "../adapters/csv/read";
import { getCell } from "./model";
import { recalc } from "./recalc";

// Excel's blank cell takes the type of whatever it is compared against: an empty cell equals both
// "" and 0. `IF(A1="","",...)` is the standard way to leave a row blank until it is filled in, and
// with the comparison coming out FALSE those cells showed the other branch's arithmetic instead of
// nothing at all.

/** Evaluate one formula on a sheet whose A1 is empty and B1 holds 5. */
function ev(formula: string): string {
  const wb = readCsv(",5\n");
  const s = wb.sheets[0]!;
  s.cells.delete("1:1"); // A1 genuinely absent, as an untouched cell is
  s.cells.set("9:9", { row: 9, col: 9, value: "", kind: "blank", formula });
  s.maxRow = 9; s.maxCol = 9;
  recalc(wb);
  return getCell(s, 9, 9)?.value ?? "";
}

describe("comparing against an empty cell", () => {
  it("equals the empty string", () => {
    expect(ev('A1=""')).toBe("TRUE");
    expect(ev('A1<>""')).toBe("FALSE");
  });

  it("still equals zero", () => {
    expect(ev("A1=0")).toBe("TRUE");
    expect(ev("A1<>0")).toBe("FALSE");
  });

  it("equals FALSE, the way a blank boolean does", () => {
    expect(ev("A1=FALSE")).toBe("TRUE");
  });

  it("blanks the cell in the IF idiom the sheets are built on", () => {
    expect(ev('IF(A1="","",B1-100)')).toBe("");
    expect(ev('IF(A1<>"",B1-100,"")')).toBe("");
  });

  it("takes the other branch once the cell has something in it", () => {
    expect(ev('IF(B1="","",B1-100)')).toBe("-95");
  });

  // Two cells that are both empty compare equal, and an ordering comparison still works.
  it("leaves the ordinary comparisons alone", () => {
    expect(ev("A1=A2")).toBe("TRUE");
    expect(ev("B1>1")).toBe("TRUE");
    expect(ev('"b">"a"')).toBe("TRUE");
    expect(ev('A1<"a"')).toBe("TRUE"); // "" sorts before "a"
  });
});
