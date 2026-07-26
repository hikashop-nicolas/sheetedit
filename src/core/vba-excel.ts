import {
  cellDisplay, colToLetters, ensureCell, getCell, lettersToCol, numToStr, parseA1Ref, type Cell, type CellStyle,
  type Sheet, type SheetControl, type StyleChange, type Workbook,
} from "./model";
import { clearXlsxCellStyle, setXlsxCellStyle } from "../adapters/xlsx/styles";
import { setXlsxRowHidden } from "../adapters/xlsx/write";
import { clearOdsCellStyle, setOdsCellStyle } from "../adapters/ods/styles";
import { setCellInput } from "./workbook";
import { createTable, listWorkbookTables, renameTable, tableStyle, type WorkbookTable } from "../adapters/xlsx/tables";
import { makeFormulaEvaluator } from "./recalc";
import { canEditCell } from "./protection";
import { filterHiddenRows, sortedPositions, sortRange, sortText, type SortKey } from "./range-ops";
import { addSheet, copySheet, deleteSheet, moveSheet, setSheetVisibility, sheetsEditable } from "./sheet-ops";
import { EMPTY, NOTHING, toNumber, toStr, VbaArray, VbaError, type LazyGlobal, type VbaObject, type VbaValue } from "vbalang";

// ---------------------------------------------------------------------------
// The Excel object model for VBA (Stage 3 of _plans/done/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// Everything here maps onto sheetedit's own model, so a macro edit is an ordinary edit: it
// recalculates, renders and undoes like any other. Nothing reaches outside the tab, and the
// members that would (Shell, CreateObject, SaveAs, Quit) refuse by name.
//
// The plan's rule holds: an unmodelled member STOPS the run saying which one. A macro that
// half-runs leaves a workbook in a state its author never intended, and the user then saves it.

/** A rectangle of cells, 1-based and inclusive, as Excel counts them. */
export interface Rect { r1: number; c1: number; r2: number; c2: number }

export interface ExcelHost {
  wb: Workbook;
  /** Index of the sheet ActiveSheet refers to. A macro may change it via Activate. */
  activeSheet: number;
  /** What Selection and ActiveCell stand for. */
  selection?: Rect;
  activeCell?: { r: number; c: number };
  /** Called before a cell is written, so a run can be snapshotted into one undo step. */
  onBeforeWrite?: (sheetIndex: number, row: number, col: number) => void;
  /** Called when something other than a cell changed (sheet added, row hidden, protection). */
  onStructureChange?: () => void;
  /** Print a sheet. Absent when the host has no print path, which makes PrintOut refuse. */
  print?: (sheetIndex: number) => void;
  /**
   * What Copy / Cut put down and PasteSpecial picks up. It lives on the host rather than on a
   * Range so a macro can copy from one sheet and paste into another, as Excel does.
   */
  clipboard?: { values: string[][]; cut?: { sheetIndex: number; rect: Rect } };
  /** The items a list control draws from, which needs the host's defined-name resolution. */
  controlItems?: (ctl: SheetControl) => string[];
  /** A control's own state changed, so the host writes it back (the persisted binary, the linked
      cell) and redraws. Absent = OLEObjects is readable but not settable. */
  onControlChange?: (sheetIndex: number, ctl: SheetControl) => void;
}

const MAX_ROW = 1048576;
const MAX_COL = 16384;

const rectOf = (r1: number, c1: number, r2: number, c2: number): Rect =>
  ({ r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) });

const a1 = (r: number, c: number, absolute = true): string =>
  absolute ? `$${colToLetters(c)}$${r}` : `${colToLetters(c)}${r}`;

/** Excel's Color is a BGR long, not the RGB a CSS colour spells. */
function cssToBgr(css: string | undefined): number {
  const m = /^#?([0-9a-f]{6})$/i.exec((css ?? "").trim());
  if (!m) return 0;
  const n = parseInt(m[1]!, 16);
  return ((n & 0xff) << 16) | (n & 0xff00) | ((n >> 16) & 0xff);
}

function bgrToCss(bgr: number): string {
  const n = Math.max(0, Math.trunc(bgr));
  const r = n & 0xff, g = (n >> 8) & 0xff, b = (n >> 16) & 0xff;
  return `#${[r, g, b].map((x) => x.toString(16).padStart(2, "0")).join("")}`;
}

/** Parse the reference forms Range() accepts: A1, A1:B2, A:C, 3:5, and a workbook defined name. */
function parseRef(wb: Workbook, sheet: Sheet, text: string): Rect[] {
  const named = wb.definedNames?.get(text) ?? wb.definedNames?.get(text.toUpperCase());
  // A defined name may itself carry a sheet qualifier, which the caller's sheet then loses; the
  // area is right either way, and a name pointing at another sheet is rare in a macro.
  const src = named ? named.replace(/^.*!/, "") : text;
  const out: Rect[] = [];
  for (const part of src.split(",")) {
    const p = part.trim().replace(/\$/g, "").replace(/^.*!/, "");
    if (!p) continue;
    let m = /^([A-Za-z]{1,3})(\d{1,7}):([A-Za-z]{1,3})(\d{1,7})$/.exec(p);
    if (m) { out.push(rectOf(Number(m[2]), lettersToCol(m[1]!), Number(m[4]), lettersToCol(m[3]!))); continue; }
    m = /^([A-Za-z]{1,3})(\d{1,7})$/.exec(p);
    if (m) { const r = Number(m[2]), c = lettersToCol(m[1]!); out.push({ r1: r, c1: c, r2: r, c2: c }); continue; }
    m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(p);
    if (m) { out.push(rectOf(1, lettersToCol(m[1]!), MAX_ROW, lettersToCol(m[2]!))); continue; }
    m = /^(\d{1,7}):(\d{1,7})$/.exec(p);
    if (m) { out.push(rectOf(Number(m[1]), 1, Number(m[2]), MAX_COL)); continue; }
    throw new VbaError(`${text} is not a reference sheetedit understands`, 1004);
  }
  if (!out.length) throw new VbaError(`${text} is not a reference sheetedit understands`, 1004);
  void sheet;
  return out;
}

/** The value a cell stands for in VBA. A blank cell is Empty, which is what macros test for. */
function cellValue(cell: Cell | undefined): VbaValue {
  if (!cell || cell.value === "") return EMPTY;
  if (cell.kind === "n") {
    const n = Number(cell.value);
    return Number.isFinite(n) ? n : cell.value;
  }
  if (cell.kind === "b") return cell.value === "TRUE" || cell.value === "1" || cell.value === "true";
  if (cell.kind === "blank") return EMPTY;
  return cell.value;
}

// ---------------------------------------------------------------------------

class FontObject implements VbaObject {
  readonly typeName = "Font";
  constructor(private readonly range: RangeObject) {}
  get(name: string): VbaValue {
    const s = this.range.firstStyle();
    switch (name.toLowerCase()) {
      case "bold": return !!s?.bold;
      case "italic": return !!s?.italic;
      case "underline": return !!s?.underline;
      case "strikethrough": return !!s?.strike;
      case "size": return s?.fontSize ?? 11;
      case "name": return s?.fontFamily ?? "";
      case "color": return cssToBgr(s?.color);
      default: throw new VbaError(`Font.${name} is not supported by sheetedit`, 438);
    }
  }
  set(name: string, _args: VbaValue[], value: VbaValue): void {
    switch (name.toLowerCase()) {
      case "bold": this.range.restyle({ bold: truthy(value) }); return;
      case "italic": this.range.restyle({ italic: truthy(value) }); return;
      case "underline": this.range.restyle({ underline: truthy(value) }); return;
      case "strikethrough": this.range.restyle({ strike: truthy(value) }); return;
      case "size": this.range.restyle({ fontSize: toNumber(value) }); return;
      case "name": this.range.restyle({ fontFamily: toStr(value) }); return;
      case "color": this.range.restyle({ color: bgrToCss(toNumber(value)) }); return;
      default: throw new VbaError(`Font.${name} is not supported by sheetedit`, 438);
    }
  }
}

class InteriorObject implements VbaObject {
  readonly typeName = "Interior";
  constructor(private readonly range: RangeObject) {}
  get(name: string): VbaValue {
    if (name.toLowerCase() === "color") return cssToBgr(this.range.firstStyle()?.bg);
    throw new VbaError(`Interior.${name} is not supported by sheetedit`, 438);
  }
  set(name: string, _args: VbaValue[], value: VbaValue): void {
    if (name.toLowerCase() !== "color") throw new VbaError(`Interior.${name} is not supported by sheetedit`, 438);
    this.range.restyle({ bg: bgrToCss(toNumber(value)) });
  }
}

/** VBA's Boolean-ish: True is -1, so a plain truthiness test is not enough on its own. */
const truthy = (v: VbaValue): boolean => (typeof v === "boolean" ? v : toNumber(v) !== 0);

/**
 * Read an argument by name or by position. Real macros pass Sort and Find arguments by name
 * (Key1:=, Order1:=, LookIn:=) far more often than positionally, so both have to work.
 */
function namedArg(args: VbaValue[], names: (string | undefined)[] | undefined, name: string, position: number): VbaValue | undefined {
  if (names) {
    const at = names.findIndex((n) => n?.toLowerCase() === name.toLowerCase());
    if (at >= 0) return args[at];
    // A call that names ANY argument may still pass earlier ones positionally, so fall through.
    if (names[position] !== undefined) return undefined;
  }
  const v = args[position];
  return v === EMPTY ? undefined : v;
}

// ---------------------------------------------------------------------------

/**
 * A Range, which may cover several areas: SpecialCells and Union both produce one, and a macro
 * that hides "every blank row in A5:A44" depends on it.
 */
export class RangeObject implements VbaObject {
  readonly typeName = "Range";
  constructor(readonly host: ExcelHost, readonly sheetIndex: number, readonly areas: Rect[]) {}

  private get sheet(): Sheet {
    const s = this.host.wb.sheets[this.sheetIndex];
    if (!s) throw new VbaError("that sheet no longer exists", 9);
    return s;
  }
  get first(): Rect {
    const a = this.areas[0];
    if (!a) throw new VbaError("that range is empty", 1004);
    return a;
  }
  /**
   * A whole-column or whole-row range is clamped to what the sheet actually uses, so walking one
   * costs nothing. An explicitly written range is NOT clamped: `Range("A1:B2").Value = 7` has to
   * write all four cells even on an empty sheet.
   */
  clamped(rect: Rect): Rect {
    const sheet = this.sheet;
    return {
      r1: rect.r1, c1: rect.c1,
      r2: rect.r2 >= MAX_ROW ? Math.max(sheet.maxRow, rect.r1) : rect.r2,
      c2: rect.c2 >= MAX_COL ? Math.max(sheet.maxCol, rect.c1) : rect.c2,
    };
  }
  private *cells(): Generator<{ r: number; c: number }> {
    for (const area of this.areas) {
      const a = this.clamped(area);
      for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) yield { r, c };
    }
  }

  firstStyle(): CellStyle | undefined {
    return getCell(this.sheet, this.first.r1, this.first.c1)?.cellStyle;
  }

  /** Styling goes through the format's own writer, so a macro's change reaches the saved file. */
  restyle(change: StyleChange): void {
    const sheet = this.sheet;
    for (const { r, c } of this.cells()) {
      this.checkWritable(r, c);
      this.host.onBeforeWrite?.(this.sheetIndex, r, c);
      const cell = ensureCell(sheet, r, c);
      if (this.host.wb.kind === "ods") setOdsCellStyle(this.host.wb, sheet, cell, change);
      else setXlsxCellStyle(this.host.wb, sheet, cell, change);
      cell.edited = true;
    }
  }

  /** Excel raises rather than silently skipping when a macro writes to a protected cell. */
  private checkWritable(r: number, c: number): void {
    if (!canEditCell(this.sheet, r, c)) {
      throw new VbaError(`${this.sheet.name}!${a1(r, c, false)} is locked on a protected sheet`, 1004);
    }
  }

  private write(r: number, c: number, v: VbaValue): void {
    this.checkWritable(r, c);
    this.host.onBeforeWrite?.(this.sheetIndex, r, c);
    if (v === EMPTY) { setCellInput(this.sheet, r, c, ""); return; }
    if (typeof v === "number") {
      const cell = ensureCell(this.sheet, r, c);
      if (cell.formula != null) cell.fDirty = true;
      cell.formula = undefined;
      cell.odfFormula = undefined;
      cell.value = numToStr(v);
      cell.kind = "n";
      cell.richRuns = undefined;
      cell.phonetic = undefined;
      cell.edited = true;
      return;
    }
    if (typeof v === "boolean") {
      const cell = ensureCell(this.sheet, r, c);
      if (cell.formula != null) cell.fDirty = true;
      cell.formula = undefined;
      cell.odfFormula = undefined;
      cell.value = v ? "TRUE" : "FALSE";
      cell.kind = "b";
      cell.edited = true;
      return;
    }
    // A string goes through the same path as typing into the grid, so "=SUM(...)" becomes a
    // formula and a typed date or percentage behaves as it does everywhere else in sheetedit.
    setCellInput(this.sheet, r, c, toStr(v));
  }

  get(name: string, args: VbaValue[], argNames?: (string | undefined)[]): VbaValue {
    const lower = name.toLowerCase();
    switch (lower) {
      case "item":
        // Item(n) hands back a Range, which a macro then reads or writes.
        return args.length ? this.subRange(args) : this;
      case "": case "value": case "value2": {
        if (args.length) return this.subRange(args);
        const a = this.first;
        if (a.r1 === a.r2 && a.c1 === a.c2) return cellValue(getCell(this.sheet, a.r1, a.c1));
        // A multi-cell Value is a 1-based two-dimensional array, which is how macros read a block.
        return this.valuesArray();
      }
      case "text": {
        const cell = getCell(this.sheet, this.first.r1, this.first.c1);
        return cell ? cell.display ?? cell.value : "";
      }
      case "formula": case "formular1c1": {
        if (lower === "formular1c1") throw new VbaError("FormulaR1C1 is not supported by sheetedit", 438);
        const cell = getCell(this.sheet, this.first.r1, this.first.c1);
        return cell?.formula != null ? `=${cell.formula}` : toStr(cellValue(cell));
      }
      case "row": return this.first.r1;
      case "column": return this.first.c1;
      case "count": return this.areas.reduce((n, a) => {
        const x = this.clamped(a);
        return n + (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1);
      }, 0);
      case "areas": return new AreasObject(this.host, this.sheetIndex, this.areas);
      case "address": {
        const rowAbs = args[0] === undefined ? true : truthy(args[0]);
        const colAbs = args[1] === undefined ? true : truthy(args[1]);
        return this.areas.map((a) => {
          const one = `${colAbs ? "$" : ""}${colToLetters(a.c1)}${rowAbs ? "$" : ""}${a.r1}`;
          if (a.r1 === a.r2 && a.c1 === a.c2) return one;
          return `${one}:${colAbs ? "$" : ""}${colToLetters(a.c2)}${rowAbs ? "$" : ""}${a.r2}`;
        }).join(",");
      }
      case "cells": {
        if (!args.length) return this;
        const a = this.first;
        const r = a.r1 + toNumber(args[0]!) - 1;
        const c = a.c1 + (args[1] !== undefined ? toNumber(args[1]) - 1 : 0);
        return new RangeObject(this.host, this.sheetIndex, [{ r1: r, c1: c, r2: r, c2: c }]);
      }
      case "offset": {
        const dr = args[0] !== undefined ? toNumber(args[0]) : 0;
        const dc = args[1] !== undefined ? toNumber(args[1]) : 0;
        return new RangeObject(this.host, this.sheetIndex,
          this.areas.map((a) => ({ r1: a.r1 + dr, c1: a.c1 + dc, r2: a.r2 + dr, c2: a.c2 + dc })));
      }
      case "resize": {
        const a = this.first;
        const rows = args[0] !== undefined ? toNumber(args[0]) : a.r2 - a.r1 + 1;
        const cols = args[1] !== undefined ? toNumber(args[1]) : a.c2 - a.c1 + 1;
        return new RangeObject(this.host, this.sheetIndex,
          [{ r1: a.r1, c1: a.c1, r2: a.r1 + rows - 1, c2: a.c1 + cols - 1 }]);
      }
      case "rows": {
        if (!args.length) return new LinesObject(this, "row");
        const a = this.first;
        const r = a.r1 + toNumber(args[0]!) - 1;
        return new RangeObject(this.host, this.sheetIndex, [{ r1: r, c1: a.c1, r2: r, c2: a.c2 }]);
      }
      case "columns": {
        if (!args.length) return new LinesObject(this, "col");
        const a = this.first;
        const c = a.c1 + toNumber(args[0]!) - 1;
        return new RangeObject(this.host, this.sheetIndex, [{ r1: a.r1, c1: c, r2: a.r2, c2: c }]);
      }
      case "currentregion": {
        // The block around the anchor, grown until a wholly empty row / column stops it. That is
        // what Excel means by it, and it is how a macro finds a table it did not lay out itself.
        const a = this.first;
        const used = { r2: Math.max(1, this.sheet.maxRow), c2: Math.max(1, this.sheet.maxCol) };
        const filled = (r: number, c: number): boolean => {
          const v = getCell(this.sheet, r, c)?.value;
          return v != null && v !== "";
        };
        const rowEmpty = (r: number, c1: number, c2: number): boolean => {
          for (let c = c1; c <= c2; c++) if (filled(r, c)) return false;
          return true;
        };
        const colEmpty = (c: number, r1: number, r2: number): boolean => {
          for (let r = r1; r <= r2; r++) if (filled(r, c)) return false;
          return true;
        };
        let r1 = a.r1, r2 = a.r2, c1 = a.c1, c2 = a.c2;
        for (let grew = true; grew; ) {
          grew = false;
          while (r1 > 1 && !rowEmpty(r1 - 1, c1, c2)) { r1--; grew = true; }
          while (r2 < used.r2 && !rowEmpty(r2 + 1, c1, c2)) { r2++; grew = true; }
          while (c1 > 1 && !colEmpty(c1 - 1, r1, r2)) { c1--; grew = true; }
          while (c2 < used.c2 && !colEmpty(c2 + 1, r1, r2)) { c2++; grew = true; }
        }
        return new RangeObject(this.host, this.sheetIndex, [{ r1, c1, r2, c2 }]);
      }
      case "autofit": {
        // APPROXIMATION, and a stated one: real AutoFit measures the rendered text, which needs the
        // font metrics of a page this may not have. The width is the longest displayed value in the
        // column, in the workbook's own character units, which lands within a character or two.
        const mdw = this.sheet.maxDigitWidth ?? 7;
        for (const a of this.areas) {
          const rows = a.c2 >= MAX_COL; // a full-row range autofits heights, which we leave alone
          if (rows) continue;
          for (let c = a.c1; c <= Math.min(a.c2, MAX_COL); c++) {
            let widest = 0;
            const last = Math.min(a.r2, Math.max(1, this.sheet.maxRow));
            for (let r = a.r1; r <= last; r++) {
              const cell = getCell(this.sheet, r, c);
              const text = cell ? cellDisplay(cell) : "";
              if (text.length > widest) widest = text.length;
            }
            (this.sheet.colWidths ??= new Map()).set(c, Math.max(mdw * 2, Math.round((widest + 1) * mdw) + 5));
          }
        }
        this.host.onStructureChange?.();
        return EMPTY;
      }
      case "entirerow":
        return new RangeObject(this.host, this.sheetIndex,
          this.areas.map((a) => ({ r1: a.r1, c1: 1, r2: a.r2, c2: MAX_COL })));
      case "entirecolumn":
        return new RangeObject(this.host, this.sheetIndex,
          this.areas.map((a) => ({ r1: 1, c1: a.c1, r2: MAX_ROW, c2: a.c2 })));
      case "hidden": {
        const a = this.first;
        return a.c2 >= MAX_COL
          ? !!this.sheet.hiddenRows?.has(a.r1)
          : !!this.sheet.hiddenCols?.has(a.c1);
      }
      case "numberformat": {
        const cell = getCell(this.sheet, this.first.r1, this.first.c1);
        return cell?.numFmt != null ? String(cell.numFmt) : "General";
      }
      case "font": return new FontObject(this);
      case "interior": return new InteriorObject(this);
      case "worksheet": return new WorksheetObject(this.host, this.sheetIndex);
      case "specialcells": return this.specialCells(args);
      case "clearcontents": this.clearContents(); return EMPTY;
      case "clearformats": this.clearFormats(); return EMPTY;
      // Clear takes the contents and the formatting together, as Excel does.
      case "clear": this.clearContents(); this.clearFormats(); return EMPTY;
      case "select": {
        this.host.selection = { ...this.first };
        this.host.activeCell = { r: this.first.r1, c: this.first.c1 };
        this.host.onStructureChange?.();
        return EMPTY;
      }
      case "merge": case "unmerge": return this.setMerged(lower === "merge");
      case "sort": return this.sort(args, argNames);
      case "autofilter": return this.autoFilter(args, argNames);
      case "find": return this.find(args, argNames);
      case "copy": case "cut": return this.copy(lower === "cut", args);
      case "pastespecial": case "paste": return this.paste();
      case "replace": return this.replace(args, argNames);
      case "removeduplicates": return this.removeDuplicates(args, argNames);
      case "texttocolumns": return this.textToColumns(args, argNames);
      case "advancedfilter": return this.advancedFilter(args, argNames);
      default:
        throw new VbaError(`Range.${name} is not supported by sheetedit`, 438);
    }
  }

  set(name: string, args: VbaValue[], value: VbaValue): void {
    const lower = name.toLowerCase();
    switch (lower) {
      case "": case "value": case "value2": case "item": {
        const target = args.length ? this.subRange(args) : this;
        target.writeAll(value);
        return;
      }
      case "formula": {
        const text = toStr(value);
        for (const { r, c } of this.cells()) {
          this.checkWritable(r, c);
          this.host.onBeforeWrite?.(this.sheetIndex, r, c);
          setCellInput(this.sheet, r, c, text.startsWith("=") ? text : `=${text}`);
        }
        return;
      }
      case "numberformat": {
        const fmt = toStr(value);
        for (const { r, c } of this.cells()) {
          this.checkWritable(r, c);
          this.host.onBeforeWrite?.(this.sheetIndex, r, c);
          const cell = ensureCell(this.sheet, r, c);
          cell.numFmt = fmt === "General" ? undefined : fmt;
          cell.numFmtDirty = true;
          cell.edited = true;
        }
        return;
      }
      case "hidden": {
        const hide = truthy(value);
        const sheet = this.sheet;
        for (const a of this.areas) {
          if (a.c2 >= MAX_COL) {
            sheet.hiddenRows ??= new Set();
            for (let r = a.r1; r <= a.r2; r++) hide ? sheet.hiddenRows.add(r) : sheet.hiddenRows.delete(r);
          } else {
            sheet.hiddenCols ??= new Set();
            for (let c = a.c1; c <= a.c2; c++) hide ? sheet.hiddenCols.add(c) : sheet.hiddenCols.delete(c);
          }
        }
        sheet.outlineDirty = true;
        this.host.onStructureChange?.();
        return;
      }
      case "formular1c1":
        throw new VbaError("FormulaR1C1 is not supported by sheetedit", 438);
      default:
        throw new VbaError(`Range.${name} cannot be assigned by sheetedit`, 438);
    }
  }

  defaultValue(): VbaValue { return this.get("value", []); }

  /** For Each over a Range walks its cells, one at a time, as Excel does. */
  enumerate(): VbaValue[] {
    const out: VbaValue[] = [];
    for (const { r, c } of this.cells()) {
      out.push(new RangeObject(this.host, this.sheetIndex, [{ r1: r, c1: c, r2: r, c2: c }]));
    }
    return out;
  }

  private subRange(args: VbaValue[]): RangeObject {
    const a = this.first;
    if (args.length >= 2) {
      const r = a.r1 + toNumber(args[0]!) - 1;
      const c = a.c1 + toNumber(args[1]!) - 1;
      return new RangeObject(this.host, this.sheetIndex, [{ r1: r, c1: c, r2: r, c2: c }]);
    }
    // A single index walks the range left to right then down, which is how Cells(n) counts.
    const width = a.c2 - a.c1 + 1;
    const n = toNumber(args[0]!) - 1;
    const r = a.r1 + Math.floor(n / width);
    const c = a.c1 + (n % width);
    return new RangeObject(this.host, this.sheetIndex, [{ r1: r, c1: c, r2: r, c2: c }]);
  }

  private valuesArray(): VbaValue {
    const a = this.clamped(this.first);
    const rows = a.r2 - a.r1 + 1, cols = a.c2 - a.c1 + 1;
    const arr = new VbaArray([1, 1], [rows, cols]);
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) arr.set([r + 1, c + 1], cellValue(getCell(this.sheet, a.r1 + r, a.c1 + c)));
    }
    return arr;
  }

  private writeAll(value: VbaValue): void {
    const arr = value instanceof VbaArray ? value : null;
    if (!arr) {
      for (const { r, c } of this.cells()) this.write(r, c, value);
      return;
    }
    // A two-dimensional array fills the range corner to corner, which is the fast way macros
    // write a block back after computing it. A ONE-dimensional array is a ROW, not a column:
    // Array("a","b","c") into A1:C1 fills across, which is what Excel does and what every macro
    // writing a header row expects.
    const a = this.first;
    const oneD = arr.lower.length === 1;
    const rows = oneD ? 1 : arr.upper[0]! - arr.lower[0]! + 1;
    const cols = oneD ? arr.upper[0]! - arr.lower[0]! + 1 : arr.upper[1]! - arr.lower[1]! + 1;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const idx = oneD ? [arr.lower[0]! + c] : [arr.lower[0]! + r, arr.lower[1]! + c];
        this.write(a.r1 + r, a.c1 + c, arr.get(idx));
      }
    }
  }

  /** Take every cell back to the workbook's default formatting. */
  private clearFormats(): void {
    const sheet = this.sheet;
    for (const { r, c } of this.cells()) {
      const cell = getCell(sheet, r, c);
      if (!cell || (!cell.style && !cell.cellStyle)) continue;
      this.checkWritable(r, c);
      this.host.onBeforeWrite?.(this.sheetIndex, r, c);
      if (this.host.wb.kind === "ods") clearOdsCellStyle(cell);
      else clearXlsxCellStyle(cell);
    }
  }

  private clearContents(): void {
    for (const { r, c } of this.cells()) {
      if (!getCell(this.sheet, r, c)) continue;
      this.checkWritable(r, c);
      this.host.onBeforeWrite?.(this.sheetIndex, r, c);
      setCellInput(this.sheet, r, c, "");
    }
  }

  private setMerged(merge: boolean): VbaValue {
    const sheet = this.sheet;
    sheet.merges ??= [];
    const a = this.clamped(this.first);
    const same = (m: Rect): boolean => m.r1 === a.r1 && m.c1 === a.c1 && m.r2 === a.r2 && m.c2 === a.c2;
    sheet.merges = sheet.merges.filter((m) => !same(m));
    if (merge && (a.r1 !== a.r2 || a.c1 !== a.c2)) sheet.merges.push(a);
    this.host.onStructureChange?.();
    return EMPTY;
  }

  /**
   * Range.AdvancedFilter. The criteria range's first row names columns of this range; each
   * following row is one set of conditions ANDed together, and the rows are ORed with each other,
   * which is exactly how Excel reads it. Action is xlFilterInPlace (1), which hides the rows that
   * do not match, or xlFilterCopy (2), which writes the matches at CopyToRange.
   */
  private advancedFilter(args: VbaValue[], names: (string | undefined)[] | undefined): VbaValue {
    const action = toNumber(namedArg(args, names, "Action", 0) ?? 1);
    const criteriaArg = namedArg(args, names, "CriteriaRange", 1);
    const unique = truthy(namedArg(args, names, "Unique", 3) ?? false);
    const rect = this.clamped(this.first);
    if (rect.r2 <= rect.r1) throw new VbaError("AdvancedFilter needs a range with a header row", 1004);

    // Which column each header name refers to, so a criteria header can be matched to it.
    const columnOf = new Map<string, number>();
    for (let c = rect.c1; c <= rect.c2; c++) columnOf.set(sortText(this.sheet, rect.r1, c).toLowerCase(), c);

    let tests: { col: number; op: string; operand: string }[][] = [];
    if (criteriaArg instanceof RangeObject) {
      const cr = criteriaArg.clamped(criteriaArg.first);
      const crSheet = this.host.wb.sheets[criteriaArg.sheetIndex];
      if (!crSheet) throw new VbaError("that sheet no longer exists", 9);
      for (let r = cr.r1 + 1; r <= cr.r2; r++) {
        const row: { col: number; op: string; operand: string }[] = [];
        for (let c = cr.c1; c <= cr.c2; c++) {
          const raw = sortText(crSheet, r, c);
          if (!raw) continue; // a blank criterion cell places no condition
          const header = sortText(crSheet, cr.r1, c).toLowerCase();
          const col = columnOf.get(header);
          if (col === undefined) throw new VbaError(`the criteria range names a column "${sortText(crSheet, cr.r1, c)}" the range does not have`, 1004);
          const m = /^([<>]=?|<>|=)?(.*)$/.exec(raw)!;
          row.push({ col, op: m[1] ?? "=", operand: m[2] ?? "" });
        }
        // A criteria row with nothing in it matches everything, as it does in Excel.
        tests.push(row);
      }
    }
    if (!tests.length) tests = [[]]; // no criteria range: every row qualifies

    const matches = (r: number): boolean =>
      tests.some((row) => row.every((t) => matchesCriterion(sortText(this.sheet, r, t.col), t.op, t.operand)));

    const seen = new Set<string>();
    const kept: number[] = [];
    for (let r = rect.r1 + 1; r <= rect.r2; r++) {
      if (!matches(r)) continue;
      if (unique) {
        const key = [];
        for (let c = rect.c1; c <= rect.c2; c++) key.push(sortText(this.sheet, r, c));
        const sig = key.join("\u0000");
        if (seen.has(sig)) continue;
        seen.add(sig);
      }
      kept.push(r);
    }

    if (action === 2) {
      const to = namedArg(args, names, "CopyToRange", 2);
      if (!(to instanceof RangeObject)) throw new VbaError("xlFilterCopy needs a CopyToRange", 1004);
      const dest = to.first;
      const target = this.host.wb.sheets[to.sheetIndex];
      if (!target) throw new VbaError("that sheet no longer exists", 9);
      const write = (dr: number, from: number): void => {
        for (let c = rect.c1; c <= rect.c2; c++) {
          const r = dest.r1 + dr, dc = dest.c1 + (c - rect.c1);
          if (!canEditCell(target, r, dc)) throw new VbaError(`${target.name}!${a1(r, dc, false)} is locked on a protected sheet`, 1004);
          this.host.onBeforeWrite?.(to.sheetIndex, r, dc);
          const cell = getCell(this.sheet, from, c);
          setCellInput(target, r, dc, cell?.formula != null ? `=${cell.formula}` : cell?.value ?? "");
        }
      };
      write(0, rect.r1); // the header goes across too
      kept.forEach((r, i) => write(i + 1, r));
      return EMPTY;
    }

    // In place: hide what did not qualify, through the same set the grid's own filter uses.
    const hidden = new Set<number>();
    for (let r = rect.r1 + 1; r <= rect.r2; r++) if (!kept.includes(r)) hidden.add(r);
    const sheet = this.sheet;
    if (this.host.wb.kind === "xlsx") {
      for (let r = rect.r1 + 1; r <= rect.r2; r++) {
        const was = sheet.filterHidden?.has(r) ?? false;
        if (hidden.has(r) !== was) setXlsxRowHidden(sheet, r, hidden.has(r));
      }
    }
    sheet.filterHidden = hidden.size ? hidden : undefined;
    if (this.host.wb.kind === "ods") sheet.odsDirty = true;
    this.host.onStructureChange?.();
    return EMPTY;
  }

  /**
   * Range.Replace, which rewrites matching text in place. Returns True when anything changed, as
   * Excel does, so `If Not rng.Replace(...) Then` behaves.
   */
  private replace(args: VbaValue[], names: (string | undefined)[] | undefined): VbaValue {
    const what = namedArg(args, names, "What", 0);
    if (what === undefined) throw new VbaError("Replace needs something to look for", 1004);
    const needle = toStr(what);
    if (!needle) throw new VbaError("Replace needs something to look for", 1004);
    const with_ = toStr(namedArg(args, names, "Replacement", 1) ?? "");
    const whole = toNumber(namedArg(args, names, "LookAt", 2) ?? 2) === 1; // xlWhole = 1
    const caseSensitive = truthy(namedArg(args, names, "MatchCase", 4) ?? false);
    let changed = false;
    for (const { r, c } of this.cells()) {
      const cell = getCell(this.sheet, r, c);
      // A formula's text is what Excel replaces in, and the result stays a formula.
      const text = cell?.formula != null ? `=${cell.formula}` : cell?.value ?? "";
      if (!text) continue;
      const next = whole
        ? (equalsText(text, needle, caseSensitive) ? with_ : text)
        : replaceAll(text, needle, with_, caseSensitive);
      if (next === text) continue;
      this.checkWritable(r, c);
      this.host.onBeforeWrite?.(this.sheetIndex, r, c);
      setCellInput(this.sheet, r, c, next);
      changed = true;
    }
    return changed;
  }

  /**
   * Range.RemoveDuplicates. Rows whose key columns repeat an earlier row are dropped and the rest
   * pulled up, which is what Excel does: the range keeps its size and empties out at the bottom.
   */
  private removeDuplicates(args: VbaValue[], names: (string | undefined)[] | undefined): VbaValue {
    const rect = this.clamped(this.first);
    const header = toNumber(namedArg(args, names, "Header", 1) ?? 2) === 1; // xlYes = 1
    const columnsArg = namedArg(args, names, "Columns", 0);
    // Columns are 1-based WITHIN the range. Absent means every column takes part.
    const given: VbaValue[] = columnsArg === undefined ? []
      : columnsArg instanceof VbaArray ? columnsArg.values() : [columnsArg];
    const offsets = given.length
      ? given.map((v) => toNumber(v) - 1)
      : Array.from({ length: rect.c2 - rect.c1 + 1 }, (_v, i) => i);
    const cols = offsets.map((o) => rect.c1 + o);
    if (cols.some((c) => c < rect.c1 || c > rect.c2)) throw new VbaError("a RemoveDuplicates column is outside the range", 1004);

    const r0 = rect.r1 + (header ? 1 : 0);
    const seen = new Set<string>();
    const keep: string[][] = [];
    for (let r = r0; r <= rect.r2; r++) {
      const key = cols.map((c) => sortText(this.sheet, r, c)).join("\u0000");
      if (seen.has(key)) continue;
      seen.add(key);
      const row: string[] = [];
      for (let c = rect.c1; c <= rect.c2; c++) {
        const cell = getCell(this.sheet, r, c);
        row.push(cell?.formula != null ? `=${cell.formula}` : cell?.value ?? "");
      }
      keep.push(row);
    }
    for (let r = r0; r <= rect.r2; r++) {
      const row = keep[r - r0];
      for (let c = rect.c1; c <= rect.c2; c++) {
        this.checkWritable(r, c);
        this.host.onBeforeWrite?.(this.sheetIndex, r, c);
        setCellInput(this.sheet, r, c, row ? row[c - rect.c1] ?? "" : "");
      }
    }
    return EMPTY;
  }

  /**
   * Range.TextToColumns, delimited only. The fixed-width form needs a FieldInfo list of column
   * starts that has no meaning without Excel's own text measuring, so it refuses rather than
   * splitting somewhere else.
   */
  private textToColumns(args: VbaValue[], names: (string | undefined)[] | undefined): VbaValue {
    // xlDelimited = 1, xlFixedWidth = 2.
    const kind = toNumber(namedArg(args, names, "DataType", 1) ?? 1);
    if (kind === 2) throw new VbaError("TextToColumns cannot split on fixed widths in sheetedit", 438);
    const flag = (n: string, at: number): boolean => truthy(namedArg(args, names, n, at) ?? false);
    const delimiters: string[] = [];
    if (flag("Tab", 3)) delimiters.push("\t");
    if (flag("Semicolon", 4)) delimiters.push(";");
    if (flag("Comma", 5)) delimiters.push(",");
    if (flag("Space", 6)) delimiters.push(" ");
    const other = namedArg(args, names, "OtherChar", 8);
    if (flag("Other", 7) && other !== undefined) delimiters.push(toStr(other));
    if (!delimiters.length) throw new VbaError("TextToColumns needs at least one delimiter", 1004);

    const destArg = namedArg(args, names, "Destination", 0);
    const dest = destArg instanceof RangeObject ? destArg.first : this.first;
    const destSheet = destArg instanceof RangeObject ? destArg.sheetIndex : this.sheetIndex;
    const a = this.clamped(this.first);
    // Excel splits the FIRST column of the range, row by row.
    const rows: string[][] = [];
    for (let r = a.r1; r <= a.r2; r++) rows.push(splitOn(sortText(this.sheet, r, a.c1), delimiters));
    const target = this.host.wb.sheets[destSheet];
    if (!target) throw new VbaError("that sheet no longer exists", 9);
    rows.forEach((parts, i) => parts.forEach((text, j) => {
      const r = dest.r1 + i, c = dest.c1 + j;
      if (!canEditCell(target, r, c)) throw new VbaError(`${target.name}!${a1(r, c, false)} is locked on a protected sheet`, 1004);
      this.host.onBeforeWrite?.(destSheet, r, c);
      setCellInput(target, r, c, text);
    }));
    return EMPTY;
  }

  /**
   * Range.Sort. Key1/Key2 name a column, either as a Range inside this one or as a reference
   * string; Order1/Order2 give the direction; Header says whether the first row is a title. The
   * ordering itself is range-ops', the same one the grid's filter menu uses.
   */
  private sort(args: VbaValue[], names: (string | undefined)[] | undefined): VbaValue {
    const rect = this.clamped(this.first);
    const keyCol = (v: VbaValue | undefined): number | undefined => {
      if (v === undefined || v === EMPTY) return undefined;
      if (v instanceof RangeObject) return v.areas[0]!.c1;
      const refs = parseRef(this.host.wb, this.sheet, toStr(v));
      return refs[0]?.c1;
    };
    const asc = (v: VbaValue | undefined): boolean => v === undefined || toNumber(v) !== 2; // xlDescending = 2
    const keys: SortKey[] = [];
    const k1 = keyCol(namedArg(args, names, "Key1", 0));
    if (k1 !== undefined) keys.push({ col: k1, ascending: asc(namedArg(args, names, "Order1", 1)) });
    const k2 = keyCol(namedArg(args, names, "Key2", 2));
    if (k2 !== undefined) keys.push({ col: k2, ascending: asc(namedArg(args, names, "Order2", 4)) });
    if (!keys.length) throw new VbaError("Sort needs a key column", 1004);
    // xlYes = 1, xlNo = 2, xlGuess = 0. Excel's default for Range.Sort is xlNo.
    const header = toNumber(namedArg(args, names, "Header", 5) ?? 2) === 1;
    const rows = sortedPositions(rect, header);
    for (const p of rows) {
      this.checkWritable(p.r, p.c);
      this.host.onBeforeWrite?.(this.sheetIndex, p.r, p.c);
    }
    sortRange(this.sheet, rect, keys, header);
    return EMPTY;
  }

  /**
   * Range.AutoFilter. With no arguments it toggles the filter over this range. With Field and
   * Criteria1 it filters that column. Only the criteria forms that map onto sheetedit's own value
   * filter are accepted; the rest refuse rather than filtering by something else.
   */
  private autoFilter(args: VbaValue[], names: (string | undefined)[] | undefined): VbaValue {
    const sheet = this.sheet;
    const field = namedArg(args, names, "Field", 0);
    const criteria = namedArg(args, names, "Criteria1", 1);
    if (field === undefined) {
      // Toggle, as Excel does when the call carries no field.
      const rect = this.clamped(this.first);
      sheet.autoFilter = sheet.autoFilter ? undefined : rect;
      if (!sheet.autoFilter) { sheet.filters = undefined; sheet.filterHidden = undefined; }
      this.applyFilter();
      return EMPTY;
    }
    const af = sheet.autoFilter ?? this.clamped(this.first);
    sheet.autoFilter = af;
    const col = af.c1 + toNumber(field) - 1;
    if (col < af.c1 || col > af.c2) throw new VbaError(`AutoFilter field ${toNumber(field)} is outside the range`, 1004);
    if (criteria === undefined) {
      // No criteria means "show all" for that column.
      sheet.filters?.delete(col);
      if (!sheet.filters?.size) sheet.filters = undefined;
      this.applyFilter();
      return EMPTY;
    }
    const wanted = toStr(criteria);
    const m = /^([<>]=?|<>|=)?(.*)$/.exec(wanted)!;
    const op = m[1] ?? "=";
    const operand = m[2] ?? "";
    const keep = new Set<string>();
    for (let r = af.r1 + 1; r <= af.r2; r++) {
      const text = sortText(sheet, r, col);
      if (matchesCriterion(text, op, operand)) keep.add(text);
    }
    sheet.filters ??= new Map();
    sheet.filters.set(col, keep);
    this.applyFilter();
    return EMPTY;
  }

  private applyFilter(): void {
    const sheet = this.sheet;
    const hidden = filterHiddenRows(sheet);
    const af = sheet.autoFilter;
    // filterHidden is separate from hiddenRows, and only hiddenRows goes through the outline
    // writer, so an xlsx has to be told row by row or the filtering is lost on save.
    if (af && this.host.wb.kind === "xlsx") {
      for (let r = af.r1 + 1; r <= af.r2; r++) {
        const was = sheet.filterHidden?.has(r) ?? false;
        if (hidden.has(r) !== was) setXlsxRowHidden(sheet, r, hidden.has(r));
      }
    }
    sheet.filterHidden = hidden.size ? hidden : undefined;
    if (this.host.wb.kind === "ods") sheet.odsDirty = true;
    this.host.onStructureChange?.();
  }

  /**
   * Range.Find, which macros use constantly. Returns the first matching cell as a Range, or
   * Nothing when there is none, which is exactly what `If Not c Is Nothing` tests.
   */
  private find(args: VbaValue[], names: (string | undefined)[] | undefined): VbaValue {
    const what = namedArg(args, names, "What", 0);
    if (what === undefined) throw new VbaError("Find needs something to look for", 1004);
    const needle = toStr(what);
    // xlPart = 2 (the default for Find), xlWhole = 1.
    const whole = toNumber(namedArg(args, names, "LookAt", 4) ?? 2) === 1;
    // xlValues = -4163 looks at the displayed value; xlFormulas = -4123 at what the cell holds.
    const inFormulas = toNumber(namedArg(args, names, "LookIn", 3) ?? -4163) === -4123;
    const caseSensitive = truthy(namedArg(args, names, "MatchCase", 6) ?? false);
    const norm = (s: string): string => (caseSensitive ? s : s.toLowerCase());
    const target = norm(needle);
    for (const { r, c } of this.cells()) {
      const cell = getCell(this.sheet, r, c);
      const text = inFormulas
        ? (cell?.formula != null ? `=${cell.formula}` : cell?.value ?? "")
        : sortText(this.sheet, r, c);
      const hay = norm(text);
      if (whole ? hay === target : hay.includes(target)) {
        return new RangeObject(this.host, this.sheetIndex, [{ r1: r, c1: c, r2: r, c2: c }]);
      }
    }
    return NOTHING;
  }

  /** Copy / Cut: the values go on the host's clipboard. A Cut also remembers where from. */
  private copy(cut: boolean, args: VbaValue[]): VbaValue {
    const rect = this.clamped(this.first);
    const values: string[][] = [];
    for (let r = rect.r1; r <= rect.r2; r++) {
      const row: string[] = [];
      for (let c = rect.c1; c <= rect.c2; c++) {
        const cell = getCell(this.sheet, r, c);
        row.push(cell?.formula != null ? `=${cell.formula}` : cell?.value ?? "");
      }
      values.push(row);
    }
    this.host.clipboard = { values, ...(cut ? { cut: { sheetIndex: this.sheetIndex, rect } } : {}) };
    // `Range("A1:B2").Copy Range("D1")` pastes straight away, which is the common one-liner.
    const dest = args[0];
    if (dest instanceof RangeObject) dest.paste();
    return EMPTY;
  }

  /**
   * Paste whatever Copy or Cut put down, at this range's top-left corner. Values and formulas
   * only: formats are not carried, and a PasteSpecial asking for them says so rather than
   * pasting the values and calling it done.
   */
  paste(): VbaValue {
    const clip = this.host.clipboard;
    if (!clip) throw new VbaError("there is nothing to paste", 1004);
    const a = this.first;
    clip.values.forEach((row, dr) => row.forEach((text, dc) => {
      const r = a.r1 + dr, c = a.c1 + dc;
      this.checkWritable(r, c);
      this.host.onBeforeWrite?.(this.sheetIndex, r, c);
      setCellInput(this.sheet, r, c, text);
    }));
    if (clip.cut) {
      // A Cut clears the source once it lands, which is what makes it a move.
      const src = this.host.wb.sheets[clip.cut.sheetIndex];
      if (src) {
        for (let r = clip.cut.rect.r1; r <= clip.cut.rect.r2; r++) {
          for (let c = clip.cut.rect.c1; c <= clip.cut.rect.c2; c++) {
            this.host.onBeforeWrite?.(clip.cut.sheetIndex, r, c);
            setCellInput(src, r, c, "");
          }
        }
      }
      this.host.clipboard = undefined;
    }
    return EMPTY;
  }

  /**
   * SpecialCells, restricted to the kinds a range can answer without Excel's own scan: blanks,
   * constants and formulas. Anything else refuses rather than returning a plausible wrong set.
   */
  private specialCells(args: VbaValue[]): VbaValue {
    const kind = toNumber(args[0] ?? 4);
    const wanted = (cell: Cell | undefined): boolean => {
      switch (kind) {
        case 4: return !cell || cell.value === "" || cell.kind === "blank";     // xlCellTypeBlanks
        case 2: return !!cell && cell.formula == null && cell.value !== "";     // xlCellTypeConstants
        case -4123: return !!cell && cell.formula != null;                      // xlCellTypeFormulas
        default: throw new VbaError(`SpecialCells(${kind}) is not supported by sheetedit`, 438);
      }
    };
    const found: Rect[] = [];
    for (const { r, c } of this.cells()) {
      if (!wanted(getCell(this.sheet, r, c))) continue;
      const last = found[found.length - 1];
      // Contiguous cells in a row collapse into one area, which keeps the result small.
      if (last && last.r1 === r && last.r2 === r && last.c2 === c - 1) last.c2 = c;
      else found.push({ r1: r, c1: c, r2: r, c2: c });
    }
    if (!found.length) throw new VbaError("no cells were found", 1004);
    return new RangeObject(this.host, this.sheetIndex, found);
  }
}

/** Range.Areas: the pieces a multi-area range is made of. */
class AreasObject implements VbaObject {
  readonly typeName = "Areas";
  constructor(private readonly host: ExcelHost, private readonly sheetIndex: number, private readonly areas: Rect[]) {}
  get(name: string, args: VbaValue[]): VbaValue {
    const lower = name.toLowerCase();
    if (lower === "count") return this.areas.length;
    if (lower === "" || lower === "item") {
      const a = this.areas[toNumber(args[0] ?? 1) - 1];
      if (!a) throw new VbaError("subscript out of range", 9);
      return new RangeObject(this.host, this.sheetIndex, [a]);
    }
    throw new VbaError(`Areas.${name} is not supported by sheetedit`, 438);
  }
  enumerate(): VbaValue[] {
    return this.areas.map((a) => new RangeObject(this.host, this.sheetIndex, [a]));
  }
}

/** Range.Rows / Range.Columns with no index: a count, and something For Each can walk. */
class LinesObject implements VbaObject {
  readonly typeName = "Range";
  constructor(private readonly range: RangeObject, private readonly axis: "row" | "col") {}
  get(name: string, args: VbaValue[]): VbaValue {
    const lower = name.toLowerCase();
    if (lower === "count") {
      const areas = this.range.areas;
      return areas.reduce((n, a) => n + (this.axis === "row" ? a.r2 - a.r1 + 1 : a.c2 - a.c1 + 1), 0);
    }
    if (lower === "" || lower === "item") return this.range.get(this.axis === "row" ? "rows" : "columns", args);
    return this.range.get(name, args);
  }
  set(name: string, args: VbaValue[], value: VbaValue): void { this.range.set(name, args, value); }
  enumerate(): VbaValue[] {
    const out: VbaValue[] = [];
    for (const a of this.range.areas) {
      if (this.axis === "row") {
        for (let r = a.r1; r <= a.r2; r++) out.push(new RangeObject(this.range.host, this.range.sheetIndex, [{ r1: r, c1: a.c1, r2: r, c2: a.c2 }]));
      } else {
        for (let c = a.c1; c <= a.c2; c++) out.push(new RangeObject(this.range.host, this.range.sheetIndex, [{ r1: a.r1, c1: c, r2: a.r2, c2: c }]));
      }
    }
    return out;
  }
}

// ---------------------------------------------------------------------------

export class WorksheetObject implements VbaObject {
  readonly typeName = "Worksheet";
  constructor(private readonly host: ExcelHost, readonly index: number) {}

  private get sheet(): Sheet {
    const s = this.host.wb.sheets[this.index];
    if (!s) throw new VbaError("that sheet no longer exists", 9);
    return s;
  }
  private whole(): Rect { return { r1: 1, c1: 1, r2: Math.max(1, this.sheet.maxRow), c2: Math.max(1, this.sheet.maxCol) }; }
  /** Where a bare Paste lands: the selection, or A1 when nothing is selected. */
  private selectionOrA1(): RangeObject {
    const sel = this.host.selection ?? { r1: 1, c1: 1, r2: 1, c2: 1 };
    return new RangeObject(this.host, this.index, [sel]);
  }

  get(name: string, args: VbaValue[], argNames?: (string | undefined)[]): VbaValue {
    const lower = name.toLowerCase();
    switch (lower) {
      case "name": return this.sheet.name;
      case "index": return this.index + 1;
      case "listobjects": {
        const coll = new ListObjectsCollection(this.host, this.index);
        return args.length ? coll.get("item", args) : coll;
      }
      case "oleobjects": {
        const coll = new OleObjectsCollection(this.host, this.index);
        // OLEObjects("Name") reads as a call on the property, so an argument indexes it directly.
        return args.length ? coll.get("item", args) : coll;
      }
      case "range": {
        if (!args.length) throw new VbaError("Range needs a reference", 1004);
        if (args[0] instanceof RangeObject) {
          // Range(topLeft, bottomRight) makes the block the two corners bound.
          const a = (args[0] as RangeObject).areas[0]!;
          const b = (args[1] instanceof RangeObject ? (args[1] as RangeObject).areas[0] : a)!;
          return new RangeObject(this.host, this.index, [rectOf(a.r1, a.c1, b.r2, b.c2)]);
        }
        return new RangeObject(this.host, this.index, parseRef(this.host.wb, this.sheet, toStr(args[0]!)));
      }
      case "cells": {
        if (!args.length) return new RangeObject(this.host, this.index, [{ r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL }]);
        const r = toNumber(args[0]!);
        const c = args[1] !== undefined ? toNumber(args[1]) : 1;
        return new RangeObject(this.host, this.index, [{ r1: r, c1: c, r2: r, c2: c }]);
      }
      case "rows": {
        const all = new RangeObject(this.host, this.index, [{ r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL }]);
        return all.get("rows", args);
      }
      case "columns": {
        const all = new RangeObject(this.host, this.index, [{ r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL }]);
        return all.get("columns", args);
      }
      case "usedrange": return new RangeObject(this.host, this.index, [this.whole()]);
      case "activate": case "select":
        this.host.activeSheet = this.index;
        this.host.onStructureChange?.();
        return EMPTY;
      case "protect": case "unprotect": {
        // The password is not checked: sheetedit's protection is a UI guard, not a security
        // boundary, and pretending otherwise would be worse than saying so.
        const sheet = this.sheet;
        if (lower === "protect") sheet.protection = { ...(sheet.protection ?? {}), sheet: true };
        else if (sheet.protection) sheet.protection = { ...sheet.protection, sheet: false };
        this.host.onStructureChange?.();
        return EMPTY;
      }
      case "protectcontents": return !!this.sheet.protection?.sheet;
      case "calculate": return EMPTY; // the run recalculates once at the end
      case "paste": {
        const at = args[0] instanceof RangeObject ? (args[0] as RangeObject) : this.selectionOrA1();
        return at.paste();
      }
      case "delete": {
        if (!sheetsEditable(this.host.wb)) throw new VbaError("sheets cannot be deleted in this format", 1004);
        if (this.host.wb.sheets.length <= 1) throw new VbaError("a workbook needs at least one sheet", 1004);
        deleteSheet(this.host.wb, this.index);
        // The active sheet may have just been removed from under the run.
        this.host.activeSheet = Math.min(this.host.activeSheet, this.host.wb.sheets.length - 1);
        this.host.onStructureChange?.();
        return EMPTY;
      }
      // xlSheetVisible = -1, xlSheetHidden = 0, xlSheetVeryHidden = 2.
      case "visible": return this.sheet.visibility === "veryHidden" ? 2 : this.sheet.visibility ? 0 : -1;
      case "move": {
        const to = this.destinationIndex(args, argNames);
        if (to === undefined) throw new VbaError("Move needs a Before or After sheet", 1004);
        moveSheet(this.host.wb, this.index, to);
        this.host.activeSheet = to;
        this.host.onStructureChange?.();
        return EMPTY;
      }
      case "printout": {
        if (!this.host.print) throw new VbaError("printing is not available here", 1004);
        this.host.print(this.index);
        return EMPTY;
      }
      case "copy": {
        if (!sheetsEditable(this.host.wb)) throw new VbaError("sheets cannot be copied in this format", 1004);
        const to = this.destinationIndex(args, argNames);
        // Copy with neither Before nor After makes a NEW WORKBOOK in Excel, which a page cannot do.
        if (to === undefined) throw new VbaError("Copy needs a Before or After sheet: sheetedit cannot open a new workbook", 1004);
        const at = copySheet(this.host.wb, this.index);
        moveSheet(this.host.wb, at, Math.min(to, this.host.wb.sheets.length - 1));
        this.host.activeSheet = Math.min(to, this.host.wb.sheets.length - 1);
        this.host.onStructureChange?.();
        return EMPTY;
      }
      case "exportasfixedformat":
        // A browser can only reach a PDF through the print dialog, where the user chooses it.
        throw new VbaError("sheetedit cannot export a PDF from a macro; PrintOut opens the print dialog instead", 1004);
      default:
        throw new VbaError(`Worksheet.${name} is not supported by sheetedit`, 438);
    }
  }

  set(name: string, _args: VbaValue[], value: VbaValue): void {
    const lower = name.toLowerCase();
    if (lower === "name") {
      this.sheet.name = toStr(value);
      this.host.onStructureChange?.();
      return;
    }
    if (lower === "visible") {
      // True / False, or one of the xlSheetVisible / Hidden / VeryHidden constants.
      const n = typeof value === "boolean" ? (value ? -1 : 0) : toNumber(value);
      const next = n === 2 ? "veryHidden" : n === 0 ? "hidden" : undefined;
      try {
        setSheetVisibility(this.host.wb, this.index, next);
      } catch (e) {
        throw new VbaError((e as Error).message, 1004);
      }
      this.host.onStructureChange?.();
      return;
    }
    throw new VbaError(`Worksheet.${name} cannot be assigned by sheetedit`, 438);
  }

  /**
   * Where a Move lands. Before puts the sheet at that sheet's place, After just past it, but the
   * move takes it out of the list first: everything after its old position shifts down by one, so
   * moving a sheet FORWARD lands one lower than the naive index.
   */
  private destinationIndex(args: VbaValue[], names: (string | undefined)[] | undefined): number | undefined {
    const indexOf = (v: VbaValue | undefined): number | undefined =>
      v instanceof WorksheetObject ? v.index : undefined;
    const b = indexOf(namedArg(args, names, "Before", 0));
    const a = indexOf(namedArg(args, names, "After", 1));
    const target = b ?? a;
    if (target === undefined) return undefined;
    const shift = this.index < target ? 1 : 0;
    const at = (b !== undefined ? b : target + 1) - shift;
    return Math.max(0, Math.min(at, this.host.wb.sheets.length - 1));
  }
}

/** Worksheets / Sheets: indexable by position or by name, and walkable with For Each. */
class SheetsObject implements VbaObject {
  readonly typeName = "Sheets";
  constructor(private readonly host: ExcelHost) {}
  private byArg(arg: VbaValue): WorksheetObject {
    if (typeof arg === "number") {
      const i = Math.trunc(arg) - 1;
      if (!this.host.wb.sheets[i]) throw new VbaError("subscript out of range", 9);
      return new WorksheetObject(this.host, i);
    }
    const wanted = toStr(arg);
    const i = this.host.wb.sheets.findIndex((s) => s.name === wanted);
    if (i < 0) throw new VbaError(`there is no sheet called ${wanted}`, 9);
    return new WorksheetObject(this.host, i);
  }
  get(name: string, args: VbaValue[]): VbaValue {
    const lower = name.toLowerCase();
    if (lower === "" || lower === "item") {
      if (!args.length) throw new VbaError("Worksheets needs an index or a name", 9);
      return this.byArg(args[0]!);
    }
    if (lower === "count") return this.host.wb.sheets.length;
    if (lower === "add") {
      if (!sheetsEditable(this.host.wb)) throw new VbaError("sheets cannot be added in this format", 1004);
      // Excel returns the new sheet, and macros chain straight onto it: Worksheets.Add.Name = "x".
      const at = addSheet(this.host.wb);
      this.host.activeSheet = at;
      this.host.onStructureChange?.();
      return new WorksheetObject(this.host, at);
    }
    throw new VbaError(`Worksheets.${name} is not supported by sheetedit`, 438);
  }
  enumerate(): VbaValue[] {
    return this.host.wb.sheets.map((_s, i) => new WorksheetObject(this.host, i));
  }
}

class WorkbookObject implements VbaObject {
  readonly typeName = "Workbook";
  constructor(private readonly host: ExcelHost, private readonly fileName: string) {}
  get(name: string, args: VbaValue[]): VbaValue {
    const lower = name.toLowerCase();
    switch (lower) {
      case "name": return this.fileName;
      case "fullname": case "path":
        // A browser tab has no path, and inventing one would send a macro down a wrong branch.
        throw new VbaError("a workbook open in a browser has no path", 1004);
      case "worksheets": case "sheets": {
        const sheets = new SheetsObject(this.host);
        return args.length ? sheets.get("item", args) : sheets;
      }
      case "activesheet": return new WorksheetObject(this.host, this.host.activeSheet);
      case "save": case "saveas": case "savecopyas": case "close": case "printout":
        throw new VbaError(`Workbook.${name} is not something sheetedit lets a macro do; save from the toolbar instead`, 1004);
      default:
        throw new VbaError(`Workbook.${name} is not supported by sheetedit`, 438);
    }
  }
}

/**
 * Application.WorksheetFunction, delegating to sheetedit's own formula engine by building the
 * formula it would have written. That keeps one implementation of SUM rather than two that drift.
 */
class WorksheetFunctionObject implements VbaObject {
  readonly typeName = "WorksheetFunction";
  constructor(private readonly host: ExcelHost) {}
  get(name: string, args: VbaValue[]): VbaValue {
    const parts = args.map((a) => {
      if (a instanceof RangeObject) {
        const sheet = this.host.wb.sheets[a.sheetIndex];
        const rect = a.areas[0]!;
        const ref = rect.r1 === rect.r2 && rect.c1 === rect.c2
          ? a1(rect.r1, rect.c1) : `${a1(rect.r1, rect.c1)}:${a1(rect.r2, rect.c2)}`;
        return `'${(sheet?.name ?? "").replace(/'/g, "''")}'!${ref}`;
      }
      if (typeof a === "number") return numToStr(a);
      if (typeof a === "boolean") return a ? "TRUE" : "FALSE";
      if (a === EMPTY) return '""';
      return `"${toStr(a).replace(/"/g, '""')}"`;
    });
    const evaluator = makeFormulaEvaluator(this.host.wb);
    const sheetName = this.host.wb.sheets[this.host.activeSheet]?.name ?? "";
    const res = evaluator.at(`${name}(${parts.join(",")})`, 1, 1, 1, 1, sheetName);
    if (res == null) return EMPTY;
    if (typeof res === "number" || typeof res === "string" || typeof res === "boolean") return res;
    // A FormulaError arrives as an object; a macro should see it as an error, not as text.
    throw new VbaError(`WorksheetFunction.${name} returned an error`, 1004);
  }
}

class ApplicationObject implements VbaObject {
  readonly typeName = "Application";
  /** Excel's UI switches a macro flips. They change nothing here, but a macro must be able to. */
  private readonly flags = new Map<string, VbaValue>([
    ["screenupdating", true], ["displayalerts", true], ["enableevents", true],
    ["calculation", -4105], ["statusbar", false], ["cursor", -4143], ["displaystatusbar", true],
  ]);
  constructor(private readonly host: ExcelHost, private readonly globals: () => Map<string, VbaValue>) {}
  get(name: string, args: VbaValue[]): VbaValue {
    const lower = name.toLowerCase();
    if (this.flags.has(lower)) return this.flags.get(lower)!;
    switch (lower) {
      case "worksheetfunction": return new WorksheetFunctionObject(this.host);
      case "version": return "16.0";
      case "name": return "sheetedit";
      case "activesheet": case "activecell": case "selection": case "activeworkbook":
      case "thisworkbook": case "worksheets": case "sheets": case "range": case "cells": {
        const g = this.globals().get(lower.toUpperCase());
        if (g === undefined) throw new VbaError(`Application.${name} is not available`, 438);
        return args.length && typeof g === "object" && g !== null && "get" in g
          ? (g as VbaObject).get("", args) : g;
      }
      case "calculate": return EMPTY; // the run recalculates once at the end
      case "quit": case "shell": case "createobject": case "getobject": case "run":
      case "sendkeys": case "ontime": case "wait": case "executeexcel4macro":
      case "filedialog": case "getopenfilename": case "getsaveasfilename":
        throw new VbaError(`Application.${name} reaches outside the page, which sheetedit does not do`, 1004);
      default:
        throw new VbaError(`Application.${name} is not supported by sheetedit`, 438);
    }
  }
  set(name: string, _args: VbaValue[], value: VbaValue): void {
    const lower = name.toLowerCase();
    if (this.flags.has(lower)) { this.flags.set(lower, value); return; }
    throw new VbaError(`Application.${name} cannot be assigned by sheetedit`, 438);
  }
}

// ---------------------------------------------------------------------------

/**
 * Excel's named constants. Only the ones a macro is likely to name: an unknown xl* constant hits
 * the interpreter's refusal path, which is better than defaulting it to zero and taking a branch
 * its author never meant.
 */
export const XL_CONSTANTS: Record<string, number> = {
  XLUP: -4162, XLDOWN: -4121, XLTOLEFT: -4159, XLTORIGHT: -4161,
  XLCELLTYPEBLANKS: 4, XLCELLTYPECONSTANTS: 2, XLCELLTYPEFORMULAS: -4123,
  XLCELLTYPELASTCELL: 11, XLCELLTYPEVISIBLE: 12,
  XLCALCULATIONAUTOMATIC: -4105, XLCALCULATIONMANUAL: -4135,
  XLNONE: -4142, XLSOLID: 1, XLAUTOMATIC: -4105,
  XLVALUES: -4163, XLFORMULAS: -4123, XLWHOLE: 1, XLPART: 2,
  XLBYROWS: 1, XLBYCOLUMNS: 2, XLNEXT: 1, XLPREVIOUS: 2,
  XLASCENDING: 1, XLDESCENDING: 2, XLYES: 1, XLNO: 2, XLGUESS: 0,
  XLDELIMITED: 1, XLFIXEDWIDTH: 2,
  XLFILTERINPLACE: 1, XLFILTERCOPY: 2,
  XLEDGETOP: 8, XLEDGEBOTTOM: 9, XLEDGELEFT: 7, XLEDGERIGHT: 10,
  XLTHIN: 2, XLMEDIUM: -4138, XLTHICK: 4, XLCONTINUOUS: 1,
  XLLEFT: -4131, XLRIGHT: -4152, XLCENTER: -4108, XLTOP: -4160, XLBOTTOM: -4107,
  XLSHEETVISIBLE: -1, XLSHEETHIDDEN: 0, XLSHEETVERYHIDDEN: 2,
  // ListObjects.Add's source type and header flag.
  XLSRCRANGE: 1, XLSRCEXTERNAL: 0, XLSRCXML: 2, XLSRCQUERY: 3, XLSRCMODEL: 4,
};

/**
 * The names a macro sees. Pass the result as the interpreter's `globals`; anything not here stops
 * the run by name rather than evaluating to Empty.
 */
export function excelGlobals(host: ExcelHost, fileName = "workbook"): Map<string, VbaValue | LazyGlobal> {
  const map = new Map<string, VbaValue | LazyGlobal>();
  const self = (): Map<string, VbaValue> => map as Map<string, VbaValue>;
  const activeSheetObj = (): WorksheetObject => new WorksheetObject(host, host.activeSheet);

  // Range / Cells with no sheet qualifier mean the active sheet, which can change mid-run, so
  // these are thin objects that resolve it at the moment of the call.
  const bare = (member: string): VbaObject => ({
    typeName: "Range",
    get: (_n, args) => activeSheetObj().get(member, args),
  });

  map.set("APPLICATION", new ApplicationObject(host, self));
  map.set("WORKSHEETS", new SheetsObject(host));
  map.set("SHEETS", new SheetsObject(host));
  map.set("RANGE", bare("range"));
  map.set("CELLS", bare("cells"));
  map.set("ROWS", bare("rows"));
  map.set("COLUMNS", bare("columns"));
  map.set("ACTIVEWORKBOOK", new WorkbookObject(host, fileName));
  map.set("THISWORKBOOK", new WorkbookObject(host, fileName));
  // ActiveSheet is a PROPERTY: read fresh each time it is evaluated, so `Set ws = ActiveSheet`
  // captures the sheet active at that moment instead of following whatever becomes active later.
  map.set("ACTIVESHEET", { lazy: () => activeSheetObj() });
  map.set("SELECTION", {
    typeName: "Range",
    get: (n, a) => selectionRange(host).get(n, a),
    set: (n, a, v) => selectionRange(host).set(n, a, v),
    defaultValue: () => selectionRange(host).defaultValue(),
    enumerate: () => selectionRange(host).enumerate(),
  });
  map.set("ACTIVECELL", {
    typeName: "Range",
    get: (n, a) => activeCellRange(host).get(n, a),
    set: (n, a, v) => activeCellRange(host).set(n, a, v),
    defaultValue: () => activeCellRange(host).defaultValue(),
  });
  for (const [name, value] of Object.entries(XL_CONSTANTS)) map.set(name, value);
  return map;
}

function selectionRange(host: ExcelHost): RangeObject {
  const sel = host.selection ?? (host.activeCell
    ? { r1: host.activeCell.r, c1: host.activeCell.c, r2: host.activeCell.r, c2: host.activeCell.c }
    : { r1: 1, c1: 1, r2: 1, c2: 1 });
  return new RangeObject(host, host.activeSheet, [sel]);
}

function activeCellRange(host: ExcelHost): RangeObject {
  const a = host.activeCell ?? { r: host.selection?.r1 ?? 1, c: host.selection?.c1 ?? 1 };
  return new RangeObject(host, host.activeSheet, [{ r1: a.r, c1: a.c, r2: a.r, c2: a.c }]);
}

/** One AutoFilter criterion, in the forms Excel's Criteria1 string actually takes. */
function matchesCriterion(text: string, op: string, operand: string): boolean {
  const n = Number(text), m = Number(operand);
  const numeric = text !== "" && operand !== "" && !Number.isNaN(n) && !Number.isNaN(m);
  switch (op) {
    case "=": {
      // "=" with * or ? is a wildcard match, which is how Criteria1 spells "starts with".
      if (/[*?]/.test(operand)) {
        const re = new RegExp(`^${operand.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".")}$`, "i");
        return re.test(text);
      }
      return text.toLowerCase() === operand.toLowerCase();
    }
    case "<>": return text.toLowerCase() !== operand.toLowerCase();
    case ">": return numeric ? n > m : text > operand;
    case ">=": return numeric ? n >= m : text >= operand;
    case "<": return numeric ? n < m : text < operand;
    case "<=": return numeric ? n <= m : text <= operand;
    default: throw new VbaError(`AutoFilter criterion "${op}${operand}" is not supported by sheetedit`, 1004);
  }
}

// --- ListObjects: the tables on a sheet -------------------------------------------------------
// A macro that lays data out and then makes it a table is a common shape, and the table is what
// structured references in the rest of the workbook then name. Creating one writes the whole
// package Excel looks for (see createTable), not just a model entry.

/** One entry of Worksheet.ListObjects. */
class ListObjectObject implements VbaObject {
  readonly typeName = "ListObject";
  constructor(private readonly host: ExcelHost, private tbl: WorkbookTable) {}

  /** Re-read the table's own record, since a rename or a resize moves it. */
  private current(): WorkbookTable {
    const found = listWorkbookTables(this.host.wb).find((t) => t.path === this.tbl.path);
    if (found) this.tbl = found;
    return this.tbl;
  }

  get(name: string, _args: VbaValue[]): VbaValue {
    const t = this.current();
    switch (name.toLowerCase()) {
      case "name": case "displayname": return t.displayName;
      case "range": return new RangeObject(this.host, t.sheetIndex, [{ r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 }]);
      case "headerrowrange":
        if (!t.headerRows) throw new VbaError("that table has no header row", 1004);
        return new RangeObject(this.host, t.sheetIndex, [{ r1: t.r1, c1: t.c1, r2: t.r1, c2: t.c2 }]);
      case "databodyrange": {
        const r1 = t.r1 + t.headerRows;
        if (t.r2 < r1) return NOTHING; // an empty table has no body, as in Excel
        return new RangeObject(this.host, t.sheetIndex, [{ r1, c1: t.c1, r2: t.r2, c2: t.c2 }]);
      }
      case "tablestyle": return tableStyle(this.host.wb, t);
      case "showheaders": return t.headerRows > 0;
      case "parent": return new WorksheetObject(this.host, t.sheetIndex);
    }
    throw new VbaError(`${this.typeName}.${name} is not supported by sheetedit`, 438);
  }

  set(name: string, _args: VbaValue[], value: VbaValue): void {
    const t = this.current();
    switch (name.toLowerCase()) {
      case "name": case "displayname":
        renameTable(this.host.wb, t, toStr(value));
        this.host.onStructureChange?.();
        return;
      case "tablestyle":
        tableStyle(this.host.wb, t, toStr(value));
        this.host.onStructureChange?.();
        return;
    }
    throw new VbaError(`${this.typeName}.${name} cannot be set by sheetedit`, 438);
  }
}

/** Worksheet.ListObjects: the sheet's tables, indexable by name or 1-based position. */
class ListObjectsCollection implements VbaObject {
  readonly typeName = "ListObjects";
  constructor(private readonly host: ExcelHost, private readonly sheetIndex: number) {}

  private list(): WorkbookTable[] {
    return listWorkbookTables(this.host.wb).filter((t) => t.sheetIndex === this.sheetIndex);
  }
  enumerate(): VbaValue[] { return this.list().map((t) => new ListObjectObject(this.host, t)); }

  get(name: string, args: VbaValue[]): VbaValue {
    const lower = name.toLowerCase();
    if (lower === "count") return this.list().length;
    if (lower === "item" || lower === "") {
      const key = args[0];
      if (key === undefined) throw new VbaError("ListObjects needs a name or an index", 1004);
      const list = this.list();
      if (typeof key === "number") {
        const t = list[Math.trunc(key) - 1];
        if (!t) throw new VbaError("no such table", 9);
        return new ListObjectObject(this.host, t);
      }
      const want = toStr(key).toLowerCase();
      const t = list.find((x) => x.displayName.toLowerCase() === want || x.name.toLowerCase() === want);
      if (!t) throw new VbaError(`no table named "${toStr(key)}"`, 9);
      return new ListObjectObject(this.host, t);
    }
    if (lower === "add") {
      // ListObjects.Add(SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination)
      const sourceType = args[0] !== undefined && args[0] !== EMPTY ? toNumber(args[0]) : XL_CONSTANTS.XLSRCRANGE!;
      if (sourceType !== XL_CONSTANTS.XLSRCRANGE) throw new VbaError("ListObjects.Add can only build a table from a range", 1004);
      const src = args[1];
      if (!(src instanceof RangeObject)) throw new VbaError("ListObjects.Add needs a Range as its source", 1004);
      const a = src.areas[0]!;
      // xlYes = 1, xlNo = 2, xlGuess = 0. Excel's own default here is xlNo.
      const flag = args[3] !== undefined && args[3] !== EMPTY ? toNumber(args[3]) : XL_CONSTANTS.XLNO!;
      const hasHeaders = flag === XL_CONSTANTS.XLYES || (flag === XL_CONSTANTS.XLGUESS && a.r2 > a.r1);
      const tbl = createTable(this.host.wb, src.sheetIndex, { r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 }, { hasHeaders });
      this.host.onStructureChange?.();
      return new ListObjectObject(this.host, tbl);
    }
    throw new VbaError(`${this.typeName}.${name} is not supported by sheetedit`, 438);
  }
}

// --- OLEObjects: the ActiveX controls on a sheet ---------------------------------------------
// A macro reaches an ActiveX control through Worksheet.OLEObjects("Name"), and its state through
// that object's .Object. We read those controls and can write their persisted value, so the
// members a real macro uses are modelled: everything else refuses, as ever.

/** The Forms 2.0 class name a control reports, which is what `TypeName` on it returns. */
const CONTROL_TYPE_NAME: Record<SheetControl["kind"], string> = {
  dropdown: "ComboBox", list: "ListBox", checkbox: "CheckBox", radio: "OptionButton",
  button: "CommandButton", label: "Label", spin: "SpinButton", scroll: "ScrollBar",
  groupBox: "Frame", textbox: "TextBox", toggle: "ToggleButton", image: "Image",
};

/** The control behind an OLEObject: what `.Object` returns. */
class OleControlObject implements VbaObject {
  readonly typeName: string;
  constructor(private readonly host: ExcelHost, private readonly sheetIndex: number, private readonly ctl: SheetControl) {
    // TypeName(cmb.Object) = "ComboBox" is how a macro tells one control from another, so the
    // reported name has to be the control's own class, not a wrapper's.
    this.typeName = CONTROL_TYPE_NAME[ctl.kind] ?? "Control";
  }

  private items(): string[] {
    return this.host.controlItems?.(this.ctl) ?? [];
  }
  private changed(): void {
    this.ctl.dirty = true;
    // Excel keeps a control's linked cell in step with its value, so a macro that only sets the
    // control still moves the cell. The write goes through a Range so it lands in the undo step.
    const ref = this.ctl.linkedCell?.replace(/\$/g, "").split("!").pop();
    const at = ref ? parseA1Ref(ref) : null;
    if (at) new RangeObject(this.host, this.sheetIndex, [{ r1: at.row, c1: at.col, r2: at.row, c2: at.col }]).set("value", [], this.ctl.activeXValue ?? "");
    this.host.onControlChange?.(this.sheetIndex, this.ctl);
  }
  /** The 0-based selected index, or -1. A list control's value IS its text, so the two agree. */
  private listIndex(): number {
    const items = this.items();
    const v = this.ctl.activeXValue ?? "";
    const i = items.indexOf(v);
    return i;
  }

  defaultValue(): VbaValue { return this.get("value", []); }

  get(name: string, args: VbaValue[]): VbaValue {
    switch (name.toLowerCase()) {
      case "value": case "text": return this.ctl.activeXValue ?? "";
      case "caption": return this.ctl.label ?? "";
      case "listcount": return this.items().length;
      case "listindex": return this.listIndex();
      case "list": {
        const items = this.items();
        if (!args.length) throw new VbaError("List needs an index", 1004);
        const i = toNumber(args[0]!);
        if (i < 0 || i >= items.length) throw new VbaError("List index out of range", 9);
        return items[i]!;
      }
      case "enabled": return this.ctl.visuals?.enabled ?? true;
      case "locked": return this.ctl.visuals?.locked ?? false;
      case "visible": return true;
      case "multiline": return this.ctl.visuals?.multiLine ?? false;
      case "maxlength": return this.ctl.visuals?.maxLength ?? 0;
      case "columncount": return 1;
      case "name": return this.ctl.name;
      case "font": {
        const f = this.ctl.visuals?.font ?? {};
        return {
          typeName: "Font",
          get: (n: string): VbaValue => {
            switch (n.toLowerCase()) {
              case "name": return f.name ?? "";
              case "size": return f.sizePt ?? 0;
              case "bold": return !!f.bold;
              case "italic": return !!f.italic;
              case "underline": return !!f.underline;
              case "strikethrough": return !!f.strike;
            }
            throw new VbaError(`Font.${n} is not supported by sheetedit`, 438);
          },
        };
      }
    }
    throw new VbaError(`${this.typeName}.${name} is not supported by sheetedit`, 438);
  }

  set(name: string, _args: VbaValue[], value: VbaValue): void {
    switch (name.toLowerCase()) {
      case "value": case "text": this.ctl.activeXValue = toStr(value); this.changed(); return;
      case "listindex": {
        const items = this.items();
        const i = toNumber(value);
        if (i < 0 || i >= items.length) throw new VbaError("ListIndex out of range", 380);
        this.ctl.activeXValue = items[i]!;
        this.changed();
        return;
      }
    }
    throw new VbaError(`${this.typeName}.${name} cannot be set by sheetedit`, 438);
  }
}

/** One entry of Worksheet.OLEObjects. */
class OleObjectObject implements VbaObject {
  readonly typeName = "OLEObject";
  constructor(private readonly host: ExcelHost, private readonly sheetIndex: number, private readonly ctl: SheetControl, private readonly index: number) {}

  get(name: string, _args: VbaValue[]): VbaValue {
    switch (name.toLowerCase()) {
      case "object": return new OleControlObject(this.host, this.sheetIndex, this.ctl);
      case "name": return this.ctl.name;
      case "index": return this.index;
      case "linkedcell": return this.ctl.linkedCell ?? "";
      case "listfillrange": return this.ctl.sourceRange ?? "";
      case "topleftcell": {
        const a = this.ctl.anchor;
        if (!a) throw new VbaError("that control has no anchor", 1004);
        return new RangeObject(this.host, this.sheetIndex, [{ r1: a.fromRow, c1: a.fromCol, r2: a.fromRow, c2: a.fromCol }]);
      }
      case "value": return this.ctl.activeXValue ?? "";
    }
    throw new VbaError(`${this.typeName}.${name} is not supported by sheetedit`, 438);
  }

  set(name: string, _args: VbaValue[], value: VbaValue): void {
    if (name.toLowerCase() === "value") { new OleControlObject(this.host, this.sheetIndex, this.ctl).set("value", [], value); return; }
    throw new VbaError(`${this.typeName}.${name} cannot be set by sheetedit`, 438);
  }
}

/** Worksheet.OLEObjects: indexable by name or by 1-based position, and enumerable. */
class OleObjectsCollection implements VbaObject {
  readonly typeName = "OLEObjects";
  constructor(private readonly host: ExcelHost, private readonly sheetIndex: number) {}

  private list(): SheetControl[] {
    return (this.host.wb.sheets[this.sheetIndex]?.controls ?? []).filter((c) => c.activeX);
  }
  enumerate(): VbaValue[] {
    return this.list().map((c, i) => new OleObjectObject(this.host, this.sheetIndex, c, i + 1));
  }
  get(name: string, args: VbaValue[]): VbaValue {
    const lower = name.toLowerCase();
    if (lower === "count") return this.list().length;
    if (lower === "item" || lower === "") {
      const list = this.list();
      const key = args[0];
      if (key === undefined) throw new VbaError("OLEObjects needs a name or an index", 1004);
      if (typeof key === "number") {
        const c = list[Math.trunc(key) - 1];
        if (!c) throw new VbaError("no such OLE object", 9);
        return new OleObjectObject(this.host, this.sheetIndex, c, Math.trunc(key));
      }
      const want = toStr(key).toLowerCase();
      const i = list.findIndex((c) => c.name.toLowerCase() === want);
      if (i < 0) throw new VbaError(`no OLE object named "${toStr(key)}"`, 9);
      return new OleObjectObject(this.host, this.sheetIndex, list[i]!, i + 1);
    }
    throw new VbaError(`${this.typeName}.${name} is not supported by sheetedit`, 438);
  }
}

const equalsText = (a: string, b: string, caseSensitive: boolean): boolean =>
  caseSensitive ? a === b : a.toLowerCase() === b.toLowerCase();

/** Replace every occurrence, honouring case-insensitivity without a regex over user text. */
function replaceAll(text: string, needle: string, with_: string, caseSensitive: boolean): string {
  if (caseSensitive) return text.split(needle).join(with_);
  const hay = text.toLowerCase(), find = needle.toLowerCase();
  let out = "", at = 0;
  for (;;) {
    const i = hay.indexOf(find, at);
    if (i < 0) return out + text.slice(at);
    out += text.slice(at, i) + with_;
    at = i + needle.length;
  }
}

/** Split on any of several delimiters, which is how TextToColumns takes its flags. */
function splitOn(text: string, delimiters: string[]): string[] {
  let parts = [text];
  for (const d of delimiters) {
    if (!d) continue;
    parts = parts.flatMap((p) => p.split(d));
  }
  return parts;
}
