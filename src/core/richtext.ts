import type { Cell, StyleChange, TextRun } from "./model";

// In-cell rich text: turn a cell + a character selection + a style change into per-run styling
// (TextRun[]), so a user can format part of a cell's text. Pure and unit-tested; both the xlsx and
// ods writers consume the resulting richRuns.

/** The whole-cell style expressed as a single run over the cell's value (the base when a cell has
    no per-run styling yet). */
function baseRun(cell: Cell): TextRun {
  const s = cell.cellStyle;
  const r: TextRun = { text: cell.value };
  if (s?.bold) r.bold = true;
  if (s?.italic) r.italic = true;
  if (s?.underline) r.underline = true;
  if (s?.strike) r.strike = true;
  if (s?.fontSize) r.size = s.fontSize;
  if (s?.color) r.color = s.color;
  if (s?.fontFamily) r.font = s.fontFamily;
  return r;
}

/** The cell's current runs: its rich runs (copied), or a single base run over the whole value. */
export function cellRuns(cell: Cell): TextRun[] {
  if (cell.richRuns?.length) return cell.richRuns.map((r) => ({ ...r }));
  return cell.value === "" ? [] : [baseRun(cell)];
}

function sameStyle(a: TextRun, b: TextRun): boolean {
  return !!a.bold === !!b.bold && !!a.italic === !!b.italic && !!a.underline === !!b.underline &&
    !!a.strike === !!b.strike && (a.size ?? 0) === (b.size ?? 0) && (a.color ?? "") === (b.color ?? "") && (a.font ?? "") === (b.font ?? "");
}

/** Merge adjacent runs with identical styling and drop empty ones. */
export function normalizeRuns(runs: TextRun[]): TextRun[] {
  const out: TextRun[] = [];
  for (const r of runs) {
    if (r.text === "") continue;
    const last = out[out.length - 1];
    if (last && sameStyle(last, r)) last.text += r.text;
    else out.push({ ...r });
  }
  return out;
}

/** Split the run list so a run boundary falls exactly at char offset `at`. */
function splitAt(runs: TextRun[], at: number): TextRun[] {
  const out: TextRun[] = [];
  let pos = 0;
  for (const r of runs) {
    const end = pos + r.text.length;
    if (at > pos && at < end) {
      out.push({ ...r, text: r.text.slice(0, at - pos) });
      out.push({ ...r, text: r.text.slice(at - pos) });
    } else out.push({ ...r });
    pos = end;
  }
  return out;
}

/** The runs overlapping [start,end) (after the boundaries have been split in). */
function coveredRuns(parts: TextRun[], start: number, end: number): TextRun[] {
  const out: TextRun[] = [];
  let pos = 0;
  for (const r of parts) { const e = pos + r.text.length; if (pos >= start && e <= end && r.text !== "") out.push(r); pos = e; }
  return out;
}

/** Apply a style change to characters [start,end) of the runs, returning normalized runs. Boolean
    attributes (bold/italic/underline/strike) toggle: if the whole selection already has the
    attribute they clear it, otherwise they set it. Value attributes (colour/size/font) set the
    change's value (an empty string / 0 clears). Attributes absent from the change are untouched. */
export function applyRunStyle(runs: TextRun[], start: number, end: number, change: StyleChange): TextRun[] {
  if (start >= end) return normalizeRuns(runs);
  const parts = splitAt(splitAt(runs, start), end);
  const covered = coveredRuns(parts, start, end);
  const toggle = (key: "bold" | "italic" | "underline" | "strike") => {
    const allOn = covered.length > 0 && covered.every((r) => !!r[key]);
    for (const r of covered) { if (allOn) delete r[key]; else r[key] = true; }
  };
  if (change.bold !== undefined) toggle("bold");
  if (change.italic !== undefined) toggle("italic");
  if (change.underline !== undefined) toggle("underline");
  if (change.strike !== undefined) toggle("strike");
  if (change.color !== undefined) for (const r of covered) r.color = change.color || undefined;
  if (change.fontSize !== undefined) for (const r of covered) r.size = change.fontSize || undefined;
  if (change.fontFamily !== undefined) for (const r of covered) r.font = change.fontFamily || undefined;
  return normalizeRuns(parts);
}

/** Set (rather than toggle) the change's attributes on characters [start,end) of the runs. Used
    when a whole-cell style change carries a resolved absolute value that rich cells must adopt. */
export function setRunStyle(runs: TextRun[], start: number, end: number, change: StyleChange): TextRun[] {
  if (start >= end) return normalizeRuns(runs);
  const parts = splitAt(splitAt(runs, start), end);
  const covered = coveredRuns(parts, start, end);
  const setBool = (key: "bold" | "italic" | "underline" | "strike", on: boolean | undefined) => {
    if (on === undefined) return;
    for (const r of covered) { if (on) r[key] = true; else delete r[key]; }
  };
  setBool("bold", change.bold); setBool("italic", change.italic); setBool("underline", change.underline); setBool("strike", change.strike);
  if (change.color !== undefined) for (const r of covered) r.color = change.color || undefined;
  if (change.fontSize !== undefined) for (const r of covered) r.size = change.fontSize || undefined;
  if (change.fontFamily !== undefined) for (const r of covered) r.font = change.fontFamily || undefined;
  return normalizeRuns(parts);
}

/** True when the runs all share the cell's base (whole-cell) style, i.e. there is no genuine
    per-run formatting and richRuns can be dropped in favour of the plain value + cellStyle. */
export function runsUniform(runs: TextRun[], cell: Cell): boolean {
  if (runs.length <= 1) { const b = baseRun(cell); return runs.length === 0 || sameStyle(runs[0]!, b); }
  return false;
}

/** Only these style attributes apply to a text run; align/valign/wrap/bg/border are whole-cell. */
export function isRunStyleChange(change: StyleChange): boolean {
  const runKeys = change.bold !== undefined || change.italic !== undefined || change.underline !== undefined ||
    change.strike !== undefined || change.color !== undefined || change.fontSize !== undefined || change.fontFamily !== undefined;
  const cellKeys = change.bg !== undefined || change.align !== undefined || change.valign !== undefined ||
    change.wrap !== undefined || change.border !== undefined || change.borderSides !== undefined;
  return runKeys && !cellKeys;
}
