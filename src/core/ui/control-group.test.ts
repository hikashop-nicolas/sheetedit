import { describe, expect, it } from "vitest";
import type { SheetControl } from "../model";
import { radioPeersToClear } from "./control-layer";

// One radio on means the rest of its group off, which is the only thing that makes a radio a
// radio. The group is the group box drawn around it, exactly as Excel decides it.

const at = (kind: SheetControl["kind"], r1: number, c1: number, r2: number, c2: number, checked = false): SheetControl => ({
  kind,
  name: `${kind}-${r1}-${c1}`,
  checked,
  anchor: { fromRow: r1, fromCol: c1, toRow: r2, toCol: c2, fromRowOff: 0, fromColOff: 0, toRowOff: 0, toColOff: 0 },
});

describe("radio groups", () => {
  it("clears the other checked radios when there is no group box", () => {
    const a = at("radio", 1, 1, 1, 2, true);
    const b = at("radio", 2, 1, 2, 2, true);
    const c = at("radio", 3, 1, 3, 2);
    expect(radioPeersToClear([a, b, c], c)).toEqual([a, b]);
  });

  it("keeps two group boxes independent", () => {
    const boxA = at("groupBox", 1, 1, 5, 5);
    const boxB = at("groupBox", 1, 10, 5, 15);
    const inA = at("radio", 2, 2, 2, 3, true);
    const inB = at("radio", 2, 11, 2, 12, true);
    const all = [boxA, boxB, inA, inB];
    const newInA = at("radio", 3, 2, 3, 3);
    expect(radioPeersToClear([...all, newInA], newInA)).toEqual([inA]);
  });

  it("does not pull a grouped radio into the ungrouped set", () => {
    const box = at("groupBox", 1, 1, 5, 5);
    const inside = at("radio", 2, 2, 2, 3, true);
    const outside = at("radio", 9, 9, 9, 10);
    expect(radioPeersToClear([box, inside, outside], outside)).toEqual([]);
  });

  it("leaves checkboxes alone, which are not exclusive", () => {
    const check = at("checkbox", 1, 1, 1, 2, true);
    const radio = at("radio", 2, 1, 2, 2);
    expect(radioPeersToClear([check, radio], radio)).toEqual([]);
  });

  it("ignores a radio the file never placed, since it belongs to no group", () => {
    const placed = at("radio", 1, 1, 1, 2, true);
    const floating: SheetControl = { kind: "radio", name: "floating", checked: false };
    // With no anchor it cannot be inside any box, so it shares the ungrouped set.
    expect(radioPeersToClear([placed, floating], floating)).toEqual([placed]);
  });
});
