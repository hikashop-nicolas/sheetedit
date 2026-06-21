export type CellKind = "n" | "s" | "b" | "e" | "blank";
export interface Cell {
    row: number;
    col: number;
    /** Display value: the literal, or the cached result for a formula cell. */
    value: string;
    /** Serialization hint for `value`. */
    kind: CellKind;
    /** Formula text in A1 syntax, without the leading "=". Undefined if not a formula. */
    formula?: string;
    /** xlsx: the <c> element in the worksheet DOM (for surgical edits). */
    el?: Element;
    /** ods: the original <table:table-cell> element (cloned verbatim if untouched). */
    odfFormula?: string;
    /** xlsx @s style index / ods @table:style-name, preserved across edits. */
    style?: string;
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
    doc?: Document;
    sheetData?: Element;
    path?: string;
    tableEl?: Element;
}
export interface Workbook {
    kind: "xlsx" | "ods";
    sheets: Sheet[];
    files: Record<string, Uint8Array>;
    contentDoc?: Document;
    contentPath?: string;
}
export declare function colToLetters(col: number): string;
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
