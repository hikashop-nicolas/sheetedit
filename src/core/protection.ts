import { getCell, type Sheet, type Workbook } from "./model";

// ---------------------------------------------------------------------------
// Sheet / workbook protection
// ---------------------------------------------------------------------------
// Protection is an *editing-intent* mechanism, not security: both formats store at most a hash of
// the password, and any tool that can read the file can drop the flag. sheetedit honours it so you
// cannot trash a protected sheet by accident, and says so plainly in the UI.
//
// Two layers decide whether a cell can be edited:
//   1. the sheet carries protection with sheet=true, and
//   2. the cell's style says it is locked (the default in both formats).
// Unlocking a range is therefore the way to leave input cells editable on a protected sheet.

/**
 * The per-action flags, with the ECMA-376 default for each. Every flag means "this action is
 * BLOCKED while the sheet is protected", which is the sense of the xlsx CT_SheetProtection
 * attribute of the same name. ODF states the opposite (its loext flags are permissions), so the
 * ods adapter inverts on the way in and out.
 */
export const SHEET_LOCK_DEFAULTS = {
  objects: false,
  scenarios: false,
  formatCells: true,
  formatColumns: true,
  formatRows: true,
  insertColumns: true,
  insertRows: true,
  insertHyperlinks: true,
  deleteColumns: true,
  deleteRows: true,
  selectLockedCells: false,
  selectUnlockedCells: false,
  sort: true,
  autoFilter: true,
  pivotTables: true,
} as const;

export type SheetLock = keyof typeof SHEET_LOCK_DEFAULTS;

export const SHEET_LOCKS = Object.keys(SHEET_LOCK_DEFAULTS) as SheetLock[];

/** A password hash read from a file. Preserved verbatim; sheetedit never authors or verifies one. */
export interface ProtectionPassword {
  /** Legacy 16-bit hash (xlsx @password / ODF short key). */
  legacy?: string;
  /** Modern hash: the ODF protection-key, or the xlsx hashValue + its parameters. */
  hash?: string;
  algorithmName?: string;
  saltValue?: string;
  spinCount?: string;
}

export interface SheetProtection {
  /** The sheet is protected (xlsx sheetProtection@sheet, ODF table:protected). */
  sheet: boolean;
  /** Per-action blocks; a flag absent here takes its SHEET_LOCK_DEFAULTS value. */
  locks?: Partial<Record<SheetLock, boolean>>;
  /** The file's password hash, kept so re-saving a protected sheet does not weaken it. */
  password?: ProtectionPassword;
}

export interface WorkbookProtection {
  /** Adding / removing / renaming / reordering sheets is blocked. */
  structure?: boolean;
  /** The window layout is locked (xlsx only; recorded for round-trip, not enforced here). */
  windows?: boolean;
  password?: ProtectionPassword;
}

/** Whether `flag`'s action is blocked right now: only while the sheet is actually protected. */
export function isBlocked(sheet: Sheet | undefined, flag: SheetLock): boolean {
  const prot = sheet?.protection;
  if (!prot?.sheet) return false;
  return prot.locks?.[flag] ?? SHEET_LOCK_DEFAULTS[flag];
}

/** Whether the sheet is protected at all. */
export const isProtected = (sheet: Sheet | undefined): boolean => !!sheet?.protection?.sheet;

/**
 * Whether one cell is locked *as a cell*, ignoring whether the sheet is currently protected.
 * Both formats default to locked, so only an explicit unlock (xlsx `<protection locked="0"/>`,
 * ODF `style:cell-protect="none"`) makes a cell editable under protection.
 */
export function isCellLocked(sheet: Sheet, row: number, col: number): boolean {
  return !getCell(sheet, row, col)?.cellStyle?.unlocked;
}

/** Whether a cell's content may be changed right now. */
export function canEditCell(sheet: Sheet | undefined, row: number, col: number): boolean {
  if (!sheet || !isProtected(sheet)) return true;
  return !isCellLocked(sheet, row, col);
}

/** Whether every cell in a rectangle may be changed (1-based, inclusive). */
export function canEditRange(sheet: Sheet | undefined, range: { r1: number; c1: number; r2: number; c2: number }): boolean {
  if (!sheet || !isProtected(sheet)) return true;
  // A blank cell inherits the sheet default (locked), so scan the rectangle rather than the
  // sparse cell map: an unlocked *range* is what makes a protected sheet usable.
  for (let r = range.r1; r <= range.r2; r++)
    for (let c = range.c1; c <= range.c2; c++) if (isCellLocked(sheet, r, c)) return false;
  return true;
}

/** Whether the workbook's sheet set is locked (add / delete / rename / reorder). */
export const isStructureLocked = (wb: Workbook): boolean => !!wb.protection?.structure;

/** True when this protection carries a password hash, so the UI can warn before dropping it. */
export const hasPassword = (p: { password?: ProtectionPassword } | undefined): boolean => {
  const pw = p?.password;
  return !!(pw?.legacy || pw?.hash);
};
