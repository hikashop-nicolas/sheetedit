import type { Sheet, Workbook } from "./model";

// A sheet read from a file can be left unparsed until something first needs it. Its parsed fields
// are accessors that load the sheet on first touch, so every caller keeps reading plain fields.

/** The fields only a parsed worksheet can fill. Everything else on a sheet is known up front. */
const LOADED_FIELDS = [
  "cells", "maxRow", "maxCol", "doc", "sheetData", "codeName", "hideGridLines",
  "colWidths", "maxDigitWidth", "rowHeights", "defaultRowHeight", "defaultColWidth",
  "hiddenRows", "hiddenCols", "rowOutline", "colOutline", "rowCollapsed", "colCollapsed",
  "summaryBelow", "summaryRight", "merges", "validations", "condFormats", "charts", "images",
  "shapes", "controls", "sparklines", "freeze", "paneSplit", "printSetup", "protection", "autoFilter",
] as const;

interface Pending {
  wb: Workbook;
  load: () => void;
  /** Whether the unparsed sheet may hold a formula stored without its result. */
  mayBeUncomputed: () => boolean;
}

const pending = new WeakMap<Sheet, Pending>();
const listeners = new WeakMap<Workbook, ((sheet: Sheet) => void)[]>();

/** Leave a sheet unparsed until one of its parsed fields is first read or written. */
export function deferSheet(wb: Workbook, sheet: Sheet, load: () => void, mayBeUncomputed: () => boolean): void {
  pending.set(sheet, { wb, load, mayBeUncomputed });
  const bag = sheet as unknown as Record<string, unknown>;
  for (const field of LOADED_FIELDS) {
    Object.defineProperty(sheet, field, {
      configurable: true,
      enumerable: true,
      get: () => { loadSheet(sheet); return bag[field]; },
      set: (v: unknown) => { loadSheet(sheet); bag[field] = v; },
    });
  }
}

export const isSheetLoaded = (sheet: Sheet): boolean => !pending.has(sheet);

/** Parse a deferred sheet now. A sheet already parsed is left alone. */
export function loadSheet(sheet: Sheet): void {
  const p = pending.get(sheet);
  if (!p) return;
  pending.delete(sheet);
  const bag = sheet as unknown as Record<string, unknown>;
  for (const field of LOADED_FIELDS) delete bag[field];
  sheet.cells = new Map();
  sheet.maxRow = 0;
  sheet.maxCol = 0;
  p.load();
  for (const fn of listeners.get(p.wb) ?? []) fn(sheet);
}

export const loadAllSheets = (wb: Workbook): void => wb.sheets.forEach(loadSheet);

/** Run fn on each sheet as it gets parsed from now on. */
export function onSheetLoaded(wb: Workbook, fn: (sheet: Sheet) => void): void {
  listeners.set(wb, [...(listeners.get(wb) ?? []), fn]);
}

/** For an unparsed sheet, whether it may carry formulas without results; false once parsed. */
export const deferredMayBeUncomputed = (sheet: Sheet): boolean => pending.get(sheet)?.mayBeUncomputed() ?? false;
