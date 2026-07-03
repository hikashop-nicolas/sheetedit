import type { Cell, CellKind, CellStyle, Sheet } from "./model";
import { ensureCell, getCell } from "./model";

// Undo/redo for the grid. Each user action records the affected cells' mutable
// fields before and after (plus optional extra closures for structural changes
// like merges); undoing restores the "before" fields onto the live model and
// the caller recalcs and re-renders. The stacks are bounded.

export interface CellFields {
  value: string;
  kind: CellKind;
  display?: string;
  formula?: string;
  odfFormula?: string;
  style?: string;
  cellStyle?: CellStyle;
  fDirty?: boolean;
}

export interface UndoCellChange {
  r: number;
  c: number;
  before: CellFields | null;
  after: CellFields | null;
}

export interface UndoStep {
  sheet: number;
  cells: UndoCellChange[];
  /** Structural inverses (e.g. merge/unmerge replays) run after the cell fields. */
  undoExtra?: () => void;
  redoExtra?: () => void;
}

export function snapFields(cell: Cell | undefined): CellFields | null {
  if (!cell) return null;
  return {
    value: cell.value,
    kind: cell.kind,
    display: cell.display,
    formula: cell.formula,
    odfFormula: cell.odfFormula,
    style: cell.style,
    cellStyle: cell.cellStyle
      ? { ...cell.cellStyle, borders: cell.cellStyle.borders ? { ...cell.cellStyle.borders } : undefined }
      : undefined,
    fDirty: cell.fDirty,
  };
}

export function applyFields(sheet: Sheet, r: number, c: number, f: CellFields | null): void {
  const cur = getCell(sheet, r, c);
  if (!f) {
    // The cell did not exist before: reset it to blank (the writer then clears it).
    if (!cur) return;
    const hadFormula = cur.formula != null;
    cur.value = "";
    cur.kind = "blank";
    cur.display = undefined;
    cur.formula = undefined;
    cur.odfFormula = undefined;
    cur.edited = true;
    cur.fDirty = cur.fDirty || hadFormula;
    return;
  }
  const cell = ensureCell(sheet, r, c);
  const formulaChanges = cell.formula !== f.formula;
  cell.value = f.value;
  cell.kind = f.kind;
  cell.display = f.display;
  cell.formula = f.formula;
  cell.odfFormula = f.odfFormula;
  cell.style = f.style;
  cell.cellStyle = f.cellStyle;
  cell.edited = true;
  cell.fDirty = cell.fDirty || f.fDirty || formulaChanges;
}

const CAP = 100;

export class UndoHistory {
  private un: UndoStep[] = [];
  private re: UndoStep[] = [];

  push(step: UndoStep): void {
    this.un.push(step);
    if (this.un.length > CAP) this.un.shift();
    this.re.length = 0;
  }
  popUndo(): UndoStep | null {
    const s = this.un.pop();
    if (!s) return null;
    this.re.push(s);
    return s;
  }
  popRedo(): UndoStep | null {
    const s = this.re.pop();
    if (!s) return null;
    this.un.push(s);
    return s;
  }
  get canUndo(): boolean {
    return this.un.length > 0;
  }
  get canRedo(): boolean {
    return this.re.length > 0;
  }
}
