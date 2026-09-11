import { describe, expect, it } from "vitest";
import { drawingExtent, type Sheet } from "./model";

// How far a sheet reaches is not only where its cells stop. A picture anchored below the last row
// of data used to be drawn past the end of the grid, with no rows under it and no way to scroll
// down to it: the sheet ended at row 41 while the image sat over rows 43 to 55.

const sheet = (over: Partial<Sheet> = {}): Sheet =>
  ({ name: "S", cells: new Map(), maxRow: 10, maxCol: 4, ...over }) as Sheet;

const anchor = (fromRow: number, toRow: number, fromCol = 1, toCol = 3) =>
  ({ fromRow, toRow, fromCol, toCol, fromRowOff: 0, toRowOff: 0, fromColOff: 0, toColOff: 0 });

describe("how far a sheet's drawings reach", () => {
  it("is nothing at all on a sheet with no drawings", () => {
    expect(drawingExtent(sheet())).toEqual({ rows: 0, cols: 0 });
  });

  it("counts a picture anchored below the data", () => {
    const s = sheet({ images: [{ cid: "i1", anchor: anchor(43, 55) }] as Sheet["images"] });
    expect(drawingExtent(s).rows, "a row of slack under the last one it covers").toBe(56);
  });

  it("counts shapes and charts, and reaches as far as the furthest one", () => {
    const s = sheet({
      shapes: [{ cid: "s1", geom: "rect", anchor: anchor(2, 8) }] as Sheet["shapes"],
      charts: [{ id: "c1", kind: "column", anchor: anchor(20, 30, 2, 9) }] as unknown as Sheet["charts"],
    });
    expect(drawingExtent(s)).toEqual({ rows: 31, cols: 10 });
  });

  it("ignores a shape inside a group, whose anchor is the group's own box", () => {
    // A grouped member is placed within its group, so its anchor numbers are coordinates in the
    // group rather than cells on the sheet: counting them stretches the grid to nowhere.
    const s = sheet({
      shapes: [
        { cid: "g1", geom: "rect", anchor: anchor(1, 5) },
        { cid: "m1", geom: "ellipse", anchor: anchor(900, 4000), within: "g1" },
      ] as Sheet["shapes"],
    });
    expect(drawingExtent(s).rows).toBe(6);
  });
});
