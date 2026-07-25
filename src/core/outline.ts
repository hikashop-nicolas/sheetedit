import type { Sheet } from "./model";

// Outline (row / column grouping): the model operations behind the gutter and the header menu.
// A group is a run of consecutive lines sharing an outline level; its SUMMARY line is the one just
// past the run (below / right by default, per <outlinePr>). Collapsing hides the run's members and
// marks the summary line, which is exactly what the file records.

export type Axis = "row" | "col";

const levels = (sheet: Sheet, axis: Axis): Map<number, number> =>
  axis === "row" ? (sheet.rowOutline ??= new Map()) : (sheet.colOutline ??= new Map());
const hiddenSet = (sheet: Sheet, axis: Axis): Set<number> =>
  axis === "row" ? (sheet.hiddenRows ??= new Set()) : (sheet.hiddenCols ??= new Set());
const collapsedSet = (sheet: Sheet, axis: Axis): Set<number> =>
  axis === "row" ? (sheet.rowCollapsed ??= new Set()) : (sheet.colCollapsed ??= new Set());
/** Which side the summary line sits on for this axis (Excel defaults to below / right). */
const summaryAfter = (sheet: Sheet, axis: Axis): boolean =>
  axis === "row" ? sheet.summaryBelow ?? true : sheet.summaryRight ?? true;

export const outlineLevel = (sheet: Sheet, axis: Axis, line: number): number =>
  (axis === "row" ? sheet.rowOutline : sheet.colOutline)?.get(line) ?? 0;

/** The deepest level in use on an axis (0 when nothing is grouped). */
export function maxOutlineLevel(sheet: Sheet, axis: Axis): number {
  const map = axis === "row" ? sheet.rowOutline : sheet.colOutline;
  let max = 0;
  for (const v of map?.values() ?? []) if (v > max) max = v;
  return Math.min(7, max);
}

/** Raise the outline level of every line in [from, to] by one (Excel caps depth at 7). */
export function groupLines(sheet: Sheet, axis: Axis, from: number, to: number): void {
  const map = levels(sheet, axis);
  for (let i = Math.min(from, to); i <= Math.max(from, to); i++) map.set(i, Math.min(7, (map.get(i) ?? 0) + 1));
  sheet.outlineDirty = true;
}

/** Lower the outline level of every line in [from, to] by one; a line reaching 0 leaves the map. */
export function ungroupLines(sheet: Sheet, axis: Axis, from: number, to: number): void {
  const map = levels(sheet, axis);
  const hidden = hiddenSet(sheet, axis);
  const collapsed = collapsedSet(sheet, axis);
  for (let i = Math.min(from, to); i <= Math.max(from, to); i++) {
    const next = (map.get(i) ?? 0) - 1;
    if (next > 0) map.set(i, next);
    else {
      map.delete(i);
      // A line that is no longer grouped must not stay hidden by a group that no longer exists.
      hidden.delete(i);
      collapsed.delete(i);
    }
  }
  if (!map.size) { if (axis === "row") sheet.rowOutline = undefined; else sheet.colOutline = undefined; }
  sheet.outlineDirty = true;
}

/** Remove every group on the axis and reveal what the groups had hidden. */
export function clearOutline(sheet: Sheet, axis: Axis): void {
  const map = axis === "row" ? sheet.rowOutline : sheet.colOutline;
  if (!map) return;
  const hidden = hiddenSet(sheet, axis);
  const collapsed = collapsedSet(sheet, axis);
  for (const line of map.keys()) { hidden.delete(line); collapsed.delete(line); }
  if (axis === "row") sheet.rowOutline = undefined; else sheet.colOutline = undefined;
  sheet.outlineDirty = true;
}

/** The run of consecutive lines at or above `level` that contains `line`. */
export function groupSpan(sheet: Sheet, axis: Axis, line: number, level: number): { from: number; to: number } | null {
  if (level < 1 || outlineLevel(sheet, axis, line) < level) return null;
  let from = line, to = line;
  while (outlineLevel(sheet, axis, from - 1) >= level) from--;
  while (outlineLevel(sheet, axis, to + 1) >= level) to++;
  return { from, to };
}

/** Every group on an axis, as a span per level, over lines 1..limit. */
export function outlineGroups(sheet: Sheet, axis: Axis, limit: number): { level: number; from: number; to: number; summary: number; collapsed: boolean }[] {
  const out: { level: number; from: number; to: number; summary: number; collapsed: boolean }[] = [];
  const after = summaryAfter(sheet, axis);
  const hidden = axis === "row" ? sheet.hiddenRows : sheet.hiddenCols;
  for (let level = 1; level <= maxOutlineLevel(sheet, axis); level++) {
    let i = 1;
    while (i <= limit) {
      if (outlineLevel(sheet, axis, i) < level) { i++; continue; }
      let to = i;
      while (to + 1 <= limit && outlineLevel(sheet, axis, to + 1) >= level) to++;
      // The summary line sits just past the run; clamp to the run when there is nothing past it.
      const summary = after ? (to + 1 <= limit ? to + 1 : to) : (i - 1 >= 1 ? i - 1 : i);
      let allHidden = true;
      for (let k = i; k <= to; k++) if (!hidden?.has(k)) { allHidden = false; break; }
      out.push({ level, from: i, to, summary, collapsed: allHidden });
      i = to + 1;
    }
  }
  return out;
}

/**
 * Collapse or expand the group of `line` at `level`: its members are hidden or revealed, and the
 * summary line carries the collapsed marker. Nested deeper groups stay collapsed when re-expanding,
 * matching Excel: expanding one level does not blow every inner group open.
 */
export function setGroupCollapsed(sheet: Sheet, axis: Axis, line: number, level: number, collapse: boolean, limit: number): void {
  const span = groupSpan(sheet, axis, line, level);
  if (!span) return;
  const hidden = hiddenSet(sheet, axis);
  const collapsed = collapsedSet(sheet, axis);
  const after = summaryAfter(sheet, axis);
  const summary = after ? (span.to + 1 <= limit ? span.to + 1 : span.to) : (span.from - 1 >= 1 ? span.from - 1 : span.from);
  for (let i = span.from; i <= span.to; i++) {
    if (collapse) hidden.add(i);
    else {
      // Re-expanding reveals a line only when no deeper collapsed group still covers it.
      const inner = outlineGroups(sheet, axis, limit).find((g) => g.level > level && g.from <= i && i <= g.to && collapsed.has(g.summary));
      if (!inner) hidden.delete(i);
    }
  }
  if (collapse) collapsed.add(summary); else collapsed.delete(summary);
  if (!hidden.size) { if (axis === "row") sheet.hiddenRows = undefined; else sheet.hiddenCols = undefined; }
  if (!collapsed.size) { if (axis === "row") sheet.rowCollapsed = undefined; else sheet.colCollapsed = undefined; }
  sheet.outlineDirty = true;
}

/** Show every line up to `level` and collapse everything deeper (the 1/2/3 level buttons). */
export function showOutlineLevel(sheet: Sheet, axis: Axis, level: number, limit: number): void {
  const hidden = hiddenSet(sheet, axis);
  const collapsed = collapsedSet(sheet, axis);
  collapsed.clear();
  for (let i = 1; i <= limit; i++) {
    if (outlineLevel(sheet, axis, i) > level) hidden.add(i); else hidden.delete(i);
  }
  for (const g of outlineGroups(sheet, axis, limit)) if (g.level > level) collapsed.add(g.summary);
  if (!hidden.size) { if (axis === "row") sheet.hiddenRows = undefined; else sheet.hiddenCols = undefined; }
  if (!collapsed.size) { if (axis === "row") sheet.rowCollapsed = undefined; else sheet.colCollapsed = undefined; }
  sheet.outlineDirty = true;
}
