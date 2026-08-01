import type { Sheet } from "./model";

// The settings a sheet carries beside its cells: what is frozen, what is protected, what is
// validated, how it prints, which rows are hidden or grouped.
//
// Shared per field group rather than as one blob per sheet. A blob is simpler and would
// mean two people changing different settings on the same sheet lose one of the two
// changes for no reason: freezing a pane has nothing to do with a print area. Ten groups is
// a coarse enough grain to stay simple and a fine enough one to stop that happening.
//
// Several of these are Maps and Sets, which JSON does not carry, so each group states its
// own wire form rather than relying on a generic clone.

export type SettingsGroup =
  | "merges"
  | "validations"
  | "condFormats"
  | "freeze"
  | "printSetup"
  | "protection"
  | "autoFilter"
  | "outline"
  | "sizes"
  | "visibility";

export const SETTINGS_GROUPS: readonly SettingsGroup[] = [
  "merges",
  "validations",
  "condFormats",
  "freeze",
  "printSetup",
  "protection",
  "autoFilter",
  "outline",
  "sizes",
  "visibility",
];

const mapToPairs = (m: Map<number, number> | undefined): [number, number][] => (m ? [...m] : []);
const pairsToMap = (pairs: unknown): Map<number, number> | undefined => {
  if (!Array.isArray(pairs) || !pairs.length) return undefined;
  return new Map(pairs as [number, number][]);
};
const setToList = (s: Set<number> | undefined): number[] => (s ? [...s] : []);
const listToSet = (list: unknown): Set<number> | undefined => {
  if (!Array.isArray(list) || !list.length) return undefined;
  return new Set(list as number[]);
};

/** One group's value, as it travels. Undefined means the sheet has none of that group. */
export function readGroup(sheet: Sheet, group: SettingsGroup): unknown {
  switch (group) {
    case "merges":
      return sheet.merges;
    case "validations":
      return sheet.validations;
    case "condFormats":
      return sheet.condFormats;
    case "freeze":
      return sheet.freeze ? { ...sheet.freeze, paneSplit: sheet.paneSplit ?? false } : undefined;
    case "printSetup":
      return sheet.printSetup;
    case "protection":
      return sheet.protection;
    case "autoFilter":
      return sheet.autoFilter;
    case "outline":
      return {
        summaryBelow: sheet.summaryBelow,
        summaryRight: sheet.summaryRight,
        rowOutline: mapToPairs(sheet.rowOutline),
        colOutline: mapToPairs(sheet.colOutline),
        rowCollapsed: setToList(sheet.rowCollapsed),
        colCollapsed: setToList(sheet.colCollapsed),
      };
    case "sizes":
      return { colWidths: mapToPairs(sheet.colWidths), rowHeights: mapToPairs(sheet.rowHeights) };
    case "visibility":
      return { hiddenRows: setToList(sheet.hiddenRows), hiddenCols: setToList(sheet.hiddenCols) };
  }
}

/** Put a group's value back, and say whether it changed anything the file must record. */
export function writeGroup(sheet: Sheet, group: SettingsGroup, value: unknown): void {
  switch (group) {
    case "merges":
      sheet.merges = value as Sheet["merges"];
      return;
    case "validations":
      sheet.validations = value as Sheet["validations"];
      return;
    case "condFormats":
      sheet.condFormats = value as Sheet["condFormats"];
      return;
    case "freeze": {
      const v = value as { rows: number; cols: number; paneSplit?: boolean } | undefined;
      sheet.freeze = v ? { rows: v.rows, cols: v.cols } : undefined;
      sheet.paneSplit = v?.paneSplit ?? false;
      sheet.freezeDirty = true;
      return;
    }
    case "printSetup":
      sheet.printSetup = value as Sheet["printSetup"];
      sheet.printDirty = true;
      return;
    case "protection":
      sheet.protection = value as Sheet["protection"];
      sheet.protectionDirty = true;
      return;
    case "autoFilter":
      sheet.autoFilter = value as Sheet["autoFilter"];
      return;
    case "outline": {
      const v = (value ?? {}) as Record<string, unknown>;
      sheet.summaryBelow = v.summaryBelow as boolean | undefined;
      sheet.summaryRight = v.summaryRight as boolean | undefined;
      sheet.rowOutline = pairsToMap(v.rowOutline);
      sheet.colOutline = pairsToMap(v.colOutline);
      sheet.rowCollapsed = listToSet(v.rowCollapsed);
      sheet.colCollapsed = listToSet(v.colCollapsed);
      sheet.outlineDirty = true;
      return;
    }
    case "sizes": {
      const v = (value ?? {}) as Record<string, unknown>;
      sheet.colWidths = pairsToMap(v.colWidths);
      sheet.rowHeights = pairsToMap(v.rowHeights);
      sheet.layoutDirty = true;
      return;
    }
    case "visibility": {
      const v = (value ?? {}) as Record<string, unknown>;
      sheet.hiddenRows = listToSet(v.hiddenRows);
      sheet.hiddenCols = listToSet(v.hiddenCols);
      sheet.layoutDirty = true;
      return;
    }
  }
}
