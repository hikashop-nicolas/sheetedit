import { describe, expect, it } from "vitest";
import type { Sheet } from "./model";
import { clearOutline, groupLines, groupSpan, maxOutlineLevel, outlineGroups, setGroupCollapsed, showOutlineLevel, ungroupLines } from "./outline";

const sheet = (): Sheet => ({ name: "S", cells: new Map(), maxRow: 20, maxCol: 10 });

describe("outline grouping", () => {
  it("raises and lowers the level of a span", () => {
    const s = sheet();
    groupLines(s, "row", 3, 6);
    expect([...s.rowOutline!.entries()]).toEqual([[3, 1], [4, 1], [5, 1], [6, 1]]);
    groupLines(s, "row", 4, 5);
    expect(s.rowOutline!.get(4)).toBe(2);
    expect(s.rowOutline!.get(3)).toBe(1);
    ungroupLines(s, "row", 4, 5);
    expect(s.rowOutline!.get(4)).toBe(1);
    ungroupLines(s, "row", 3, 6);
    expect(s.rowOutline).toBeUndefined();
  });

  it("caps the depth at 7, the way Excel does", () => {
    const s = sheet();
    for (let i = 0; i < 10; i++) groupLines(s, "row", 2, 3);
    expect(maxOutlineLevel(s, "row")).toBe(7);
  });

  it("finds the run of consecutive lines at a level", () => {
    const s = sheet();
    groupLines(s, "row", 3, 6);
    expect(groupSpan(s, "row", 5, 1)).toEqual({ from: 3, to: 6 });
    expect(groupSpan(s, "row", 8, 1)).toBeNull();
    // A deeper level narrows the run to the inner group.
    groupLines(s, "row", 4, 5);
    expect(groupSpan(s, "row", 4, 2)).toEqual({ from: 4, to: 5 });
    expect(groupSpan(s, "row", 4, 1)).toEqual({ from: 3, to: 6 });
  });

  it("lists a group per level with its summary line below", () => {
    const s = sheet();
    groupLines(s, "row", 3, 6);
    groupLines(s, "row", 4, 5);
    const gs = outlineGroups(s, "row", 20);
    expect(gs).toEqual([
      { level: 1, from: 3, to: 6, summary: 7, collapsed: false },
      { level: 2, from: 4, to: 5, summary: 6, collapsed: false },
    ]);
  });

  it("puts the summary line above when summaryBelow is false", () => {
    const s = sheet();
    s.summaryBelow = false;
    groupLines(s, "row", 3, 6);
    expect(outlineGroups(s, "row", 20)[0]!.summary).toBe(2);
  });

  it("collapses a group by hiding its members and marking the summary", () => {
    const s = sheet();
    groupLines(s, "row", 3, 6);
    setGroupCollapsed(s, "row", 4, 1, true, 20);
    expect([...s.hiddenRows!].sort((a, b) => a - b)).toEqual([3, 4, 5, 6]);
    expect(s.rowCollapsed!.has(7)).toBe(true);
    expect(outlineGroups(s, "row", 20)[0]!.collapsed).toBe(true);
    setGroupCollapsed(s, "row", 4, 1, false, 20);
    expect(s.hiddenRows).toBeUndefined();
    expect(s.rowCollapsed).toBeUndefined();
  });

  it("keeps an inner collapsed group collapsed when the outer one re-expands", () => {
    const s = sheet();
    groupLines(s, "row", 3, 8);
    groupLines(s, "row", 5, 6);
    setGroupCollapsed(s, "row", 5, 2, true, 20); // inner
    setGroupCollapsed(s, "row", 4, 1, true, 20); // outer
    setGroupCollapsed(s, "row", 4, 1, false, 20);
    // The outer group's own rows come back, but the inner group's stay hidden.
    expect([...s.hiddenRows!].sort((a, b) => a - b)).toEqual([5, 6]);
  });

  it("shows a level and collapses everything deeper", () => {
    const s = sheet();
    groupLines(s, "row", 3, 8);
    groupLines(s, "row", 5, 6);
    showOutlineLevel(s, "row", 1, 20);
    expect([...s.hiddenRows!].sort((a, b) => a - b)).toEqual([5, 6]);
    showOutlineLevel(s, "row", 2, 20);
    expect(s.hiddenRows).toBeUndefined();
    showOutlineLevel(s, "row", 0, 20);
    expect([...s.hiddenRows!].sort((a, b) => a - b)).toEqual([3, 4, 5, 6, 7, 8]);
  });

  it("ungrouping reveals what the group had hidden", () => {
    const s = sheet();
    groupLines(s, "row", 3, 6);
    setGroupCollapsed(s, "row", 4, 1, true, 20);
    ungroupLines(s, "row", 3, 6);
    expect(s.hiddenRows?.size ?? 0).toBe(0);
    expect(s.rowOutline).toBeUndefined();
  });

  it("clears every group on an axis", () => {
    const s = sheet();
    groupLines(s, "col", 2, 5);
    setGroupCollapsed(s, "col", 3, 1, true, 10);
    clearOutline(s, "col");
    expect(s.colOutline).toBeUndefined();
    expect(s.hiddenCols?.size ?? 0).toBe(0);
  });

  it("keeps a row hidden by hand out of the outline's way", () => {
    const s = sheet();
    s.hiddenRows = new Set([12]);
    groupLines(s, "row", 3, 6);
    setGroupCollapsed(s, "row", 4, 1, true, 20);
    setGroupCollapsed(s, "row", 4, 1, false, 20);
    expect([...s.hiddenRows!]).toEqual([12]);
  });
});
