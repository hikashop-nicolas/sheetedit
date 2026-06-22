export type CellKind = "n" | "s" | "b" | "e" | "blank";
/** Resolved visual formatting for a cell (read from the file's style pools). */
export interface CellStyle {
    bold?: boolean;
    italic?: boolean;
    color?: string;
    bg?: string;
    align?: "left" | "center" | "right";
    /** Border presence + CSS colour per side. */
    borders?: {
        top?: string;
        right?: string;
        bottom?: string;
        left?: string;
    };
}
export interface Cell {
    row: number;
    col: number;
    /** Canonical/editable value: the literal, or the cached result for a formula cell. */
    value: string;
    /** Serialization hint for `value`. */
    kind: CellKind;
    /** Formatted text for the grid (number format applied). Falls back to `value`. */
    display?: string;
    /** Number format for this cell: an xlsx format code, a built-in numFmtId, or none. */
    numFmt?: string | number;
    /** Formula text in A1 syntax, without the leading "=". Undefined if not a formula. */
    formula?: string;
    /** xlsx: the <c> element in the worksheet DOM (for surgical edits). */
    el?: Element;
    /** ods: the original <table:table-cell> element (cloned verbatim if untouched). */
    odfFormula?: string;
    /** xlsx @s style index / ods @table:style-name, preserved across edits. */
    style?: string;
    /** Resolved visual formatting (fonts/fills/borders/alignment) for the grid. */
    cellStyle?: CellStyle;
    /** User changed the value/formula (forces regeneration). */
    edited?: boolean;
    /** Recalc changed the cached value. */
    recomputed?: boolean;
}
export interface Sheet {
    name: string;
    cells: Map<string, Cell>;
    maxRow: number;
    maxCol: number;
    /** 1-based column -> width in px (from the file's <cols>), when specified. */
    colWidths?: Map<number, number>;
    /** 1-based row -> height in px (from the file's <row ht>), when specified. */
    rowHeights?: Map<number, number>;
    /** Merged ranges (1-based, inclusive); the top-left cell holds the value. */
    merges?: {
        r1: number;
        c1: number;
        r2: number;
        c2: number;
    }[];
    doc?: Document;
    sheetData?: Element;
    path?: string;
    /** Column widths or row heights changed and the sheet XML must be re-serialized. */
    layoutDirty?: boolean;
    tableEl?: Element;
}
export interface Workbook {
    kind: "xlsx" | "ods";
    sheets: Sheet[];
    files: Record<string, Uint8Array>;
    contentDoc?: Document;
    contentPath?: string;
    stylesDoc?: Document;
    stylesDirty?: boolean;
}
/** A style change to apply to a cell (only the set fields change). */
export interface StyleChange {
    bold?: boolean;
    italic?: boolean;
    color?: string;
    bg?: string;
    align?: "left" | "center" | "right";
    border?: boolean;
    /** Per-side borders to set; each specified side is turned on/off, others kept. */
    borderSides?: {
        top?: boolean;
        right?: boolean;
        bottom?: boolean;
        left?: boolean;
    };
}
export declare function colToLetters(col: number): string;
export declare function setXlsxColWidth(sheet: Sheet, col: number, px: number): void;
export declare function setXlsxRowHeight(sheet: Sheet, row: number, px: number): void;
export declare function setXlsxMerge(sheet: Sheet, r1: number, c1: number, r2: number, c2: number, merge: boolean): void;
/**
 * Apply a style change to a cell, managing the xlsx style pools: derive a new font /
 * fill / border from the cell's current format plus the change, find-or-create each in
 * styles.xml, find-or-create the combined <xf>, and point the cell at it.
 */
export declare function setXlsxCellStyle(wb: Workbook, sheet: Sheet, cell: Cell, change: StyleChange): void;
/** ODF formula (`of:=[.A1]+[.B1]`) -> A1 (`A1+B1`). */
export declare function odfToA1(odf: string): string;
/** A1 (`A1+B1`) -> ODF formula (`of:=[.A1]+[.B1]`). Used only for user-typed formulas. */
export declare function a1ToOdf(a1: string): string;
/** Recompute every formula cell's cached value, in dependency order. */
export declare function recalc(wb: Workbook): void;
export declare function readWorkbook(bytes: Uint8Array): Workbook;
export declare function writeWorkbook(wb: Workbook): Uint8Array;
/** Commit a raw grid edit (the text a user typed) into the model. */
export declare function setCellInput(sheet: Sheet, row: number, col: number, raw: string): void;
export interface SheetEditorOptions {
    onChange?: () => void;
}
export interface SheetEditor {
    getBytes(): Promise<Uint8Array>;
    isDirty(): boolean;
    destroy(): void;
}
export declare function createSheetEditor(container: HTMLElement, bytes: Uint8Array, options?: SheetEditorOptions): SheetEditor;
