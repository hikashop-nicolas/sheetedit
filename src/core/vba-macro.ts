import { applyFields, snapFields, type CellFields } from "./history";
import { getCell, type Sheet, type Workbook } from "./model";
import { parseModule } from "./vba-parse";
import { VbaInterpreter } from "./vba-run";
import { VbaError } from "./vba-value";
import { VbaSyntaxError } from "./vba-lex";
import { excelGlobals, type ExcelHost, type Rect } from "./vba-excel";

// ---------------------------------------------------------------------------
// Running a macro against the workbook (Stage 4 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// A whole run is one undo step, and a run that stops part-way is rolled back before the caller
// sees the error. That is the plan's central rule: a macro that half-runs leaves a workbook in a
// state its author never intended, and the user then saves it.

/** The sheet-level state a macro can change without touching a cell. */
interface SheetState {
  name: string;
  hiddenRows: number[];
  hiddenCols: number[];
  merges: Rect[];
  protection: Sheet["protection"];
}

interface Snapshot {
  cells: { sheet: number; r: number; c: number; fields: CellFields | null }[];
  sheets: SheetState[];
}

const captureSheets = (wb: Workbook): SheetState[] => wb.sheets.map((s) => ({
  name: s.name,
  hiddenRows: [...(s.hiddenRows ?? [])],
  hiddenCols: [...(s.hiddenCols ?? [])],
  merges: (s.merges ?? []).map((m) => ({ ...m })),
  protection: s.protection ? { ...s.protection } : undefined,
}));

function restoreSheets(wb: Workbook, states: SheetState[]): void {
  states.forEach((st, i) => {
    const sheet = wb.sheets[i];
    if (!sheet) return;
    sheet.name = st.name;
    sheet.hiddenRows = new Set(st.hiddenRows);
    sheet.hiddenCols = new Set(st.hiddenCols);
    sheet.merges = st.merges.map((m) => ({ ...m }));
    sheet.protection = st.protection ? { ...st.protection } : undefined;
    sheet.outlineDirty = true;
  });
}

export interface MacroRunOptions {
  /** The sheet ActiveSheet starts on, and what Selection / ActiveCell stand for. */
  activeSheet?: number;
  selection?: Rect;
  activeCell?: { r: number; c: number };
  /** Name shown in the workbook's own Name property. */
  fileName?: string;
  /** Lowered in tests; the default is high enough that no honest macro reaches it. */
  maxSteps?: number;
}

export interface MacroRunResult {
  ok: boolean;
  /** What MsgBox and Debug.Print produced, in order. */
  messages: string[];
  /** Set when the run stopped. The workbook has already been rolled back. */
  error?: { message: string; module: string; line?: number };
  /** Where the macro left the view. Absent when nothing moved. */
  activeSheet?: number;
  selection?: Rect;
  activeCell?: { r: number; c: number };
  /** Replay the run's changes, for undo/redo. Absent when the run changed nothing. */
  undo?: () => void;
  redo?: () => void;
}

/** The Subs a user can pick: public, and taking no argument that must be supplied. */
export function runnableSubs(source: string): string[] {
  try {
    return new VbaInterpreter(parseModule(source)).runnableSubs();
  } catch {
    // A module sheetedit cannot parse offers nothing to run; the viewer still shows its source.
    return [];
  }
}

/**
 * Run one procedure against `wb`. Every cell the macro touches is snapshotted before it changes,
 * so the run collapses into a single undo step, and a failed run is undone before returning.
 */
export function runWorkbookMacro(
  wb: Workbook,
  source: string,
  moduleName: string,
  procName: string,
  options: MacroRunOptions = {},
): MacroRunResult {
  let module: ReturnType<typeof parseModule>;
  try {
    module = parseModule(source, moduleName);
  } catch (e) {
    const line = e instanceof VbaSyntaxError ? e.line : undefined;
    return { ok: false, messages: [], error: { message: (e as Error).message, module: moduleName, line } };
  }

  const before: Snapshot = { cells: [], sheets: captureSheets(wb) };
  const seen = new Set<string>();
  const host: ExcelHost = {
    wb,
    activeSheet: options.activeSheet ?? 0,
    selection: options.selection,
    activeCell: options.activeCell,
    onBeforeWrite: (sheet, r, c) => {
      // Only the first write to a cell matters: that is the state undo has to get back to.
      const k = `${sheet}:${r}:${c}`;
      if (seen.has(k)) return;
      seen.add(k);
      before.cells.push({ sheet, r, c, fields: snapFields(getCell(wb.sheets[sheet]!, r, c)) });
    },
  };

  const interp = new VbaInterpreter(module, {
    globals: excelGlobals(host, options.fileName ?? "workbook"),
    maxSteps: options.maxSteps,
  });

  let result: { messages: string[] };
  try {
    result = interp.run(procName);
  } catch (e) {
    // Roll back before the caller sees anything, so a stopped macro leaves no trace.
    restore(wb, before);
    const line = e instanceof VbaError ? e.line : undefined;
    return { ok: false, messages: [], error: { message: (e as Error).message, module: moduleName, line } };
  }

  const after: Snapshot = { cells: capture(wb, before.cells), sheets: captureSheets(wb) };
  const changed = before.cells.length > 0 || JSON.stringify(before.sheets) !== JSON.stringify(after.sheets);
  return {
    ok: true,
    messages: result.messages,
    activeSheet: host.activeSheet,
    selection: host.selection,
    activeCell: host.activeCell,
    ...(changed ? { undo: () => restore(wb, before), redo: () => restore(wb, after) } : {}),
  };
}

const capture = (wb: Workbook, positions: Snapshot["cells"]): Snapshot["cells"] =>
  positions.map((p) => ({ sheet: p.sheet, r: p.r, c: p.c, fields: snapFields(getCell(wb.sheets[p.sheet]!, p.r, p.c)) }));

function restore(wb: Workbook, snap: Snapshot): void {
  for (const c of snap.cells) {
    const sheet = wb.sheets[c.sheet];
    if (sheet) applyFields(sheet, c.r, c.c, c.fields);
  }
  restoreSheets(wb, snap.sheets);
}
