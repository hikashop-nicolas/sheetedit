import FormulaParser from "fast-formula-parser";
import { strFromU8, strToU8 } from "fflate";
// ---------------------------------------------------------------------------
// Model
// ---------------------------------------------------------------------------

export type CellKind = "n" | "s" | "b" | "e" | "blank";

/** Resolved visual formatting for a cell (read from the file's style pools). */
export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** Underline flavour as a CSS text-decoration-style (double/dotted/dashed/wavy); absent =
      plain solid. Preserved from the file so re-styling a cell keeps its flavour. */
  underlineStyle?: string;
  strike?: boolean;
  /** Font size in points (only when it differs from the workbook default). */
  fontSize?: number;
  /** Font family name (only when it differs from the workbook default). */
  fontFamily?: string;
  color?: string; // CSS text colour
  bg?: string; // CSS fill colour
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
  /** Wrap text within the cell. */
  wrap?: boolean;
  /** Border presence + CSS colour per side. */
  borders?: { top?: string; right?: string; bottom?: string; left?: string };
  /** Theme colour references kept alongside the resolved colours, so switching the workbook's
      theme recolours the cell instead of baking in the palette it was read with. */
  colorRef?: import("./theme").ThemeColorRef;
  bgRef?: import("./theme").ThemeColorRef;
  borderRefs?: { top?: import("./theme").ThemeColorRef; right?: import("./theme").ThemeColorRef; bottom?: import("./theme").ThemeColorRef; left?: import("./theme").ThemeColorRef };
  /** The font follows the theme's heading (major) or body (minor) typeface. */
  fontScheme?: "major" | "minor";
  /** The cell is explicitly unlocked (xlsx `<protection locked="0"/>`, ODF cell-protect "none").
      Both formats default to locked, so absent = locked; this only matters once the sheet is
      protected, when locked cells become read-only. */
  unlocked?: boolean;
  /** The formula is hidden from the formula bar while the sheet is protected. Preserved for a
      faithful round-trip; sheetedit renders the value either way. */
  formulaHidden?: boolean;
}

/** A list-type data validation (dropdown): the ranges it covers and its allowed values,
    either inline (`values`) or a range reference resolved at render time (`rangeRef`). Stored on
    the sheet (not per cell) so a whole-column validation costs nothing. */
export interface DataValidation {
  ranges: { r1: number; c1: number; r2: number; c2: number }[];
  /** list dropdown: the inline allowed values (or a range via rangeRef). */
  values?: string[];
  rangeRef?: string;
  allowBlank?: boolean;
  /** Validation kind. Absent / "list" is the dropdown; the others constrain the entered value. */
  type?: "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom";
  /** Comparison for the non-list, non-custom types. */
  operator?: "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
  /** Operand(s): the bound(s) for a comparison, or the expression for a custom rule. */
  formula1?: string;
  formula2?: string;
}

/** A slicer: the interactive filter buttons Excel attaches to a pivot table. The selection lives in
    the slicer *cache* part (one `<x14:i x s>` per source item); the slicer part holds the view
    (caption, columns) and the drawing holds its anchor. Toggling an item re-filters the linked
    pivot(s) and the new selection is written back into the cache part. */
export interface SheetSlicer {
  /** Slicer view name (unique in the workbook) and the cache it reads. */
  name: string;
  cache: string;
  caption?: string;
  columnCount?: number;
  /** The pivot-cache field this slicer filters, by name (matches a source column header). */
  sourceName: string;
  /** Names of the pivot tables this slicer drives. */
  pivotTables: string[];
  /** What the slicer filters. "olap" is rendered read-only: we model no OLAP source to filter. */
  kind?: "pivot" | "table" | "olap";
  /** Built-in or custom slicer style name (SlicerStyleLight1 .. SlicerStyleDark6, or a custom one). */
  style?: string;
  /** Table slicers (x15:tableSlicerCache): the bound table's grid range and the column inside it. */
  table?: { sheetIndex: number; r1: number; c1: number; r2: number; c2: number; headerRows: number; col: number };
  /** Every item, in cache order, with its label and selected state. */
  items: { x: number; label: string; selected: boolean }[];
  anchor?: import("./chart-model").ChartAnchor;
  /** Part paths, so a selection change can be written back. */
  slicerPath?: string;
  cachePath?: string;
  /** Selection changed in the UI -> the cache part is rewritten on save. */
  dirty?: boolean;
}

/** A timeline: Excel's date-range filter for a pivot. The view part holds the caption and the
    granularity (level); the cache part holds the selected range and the overall bounds. Changing the
    range re-filters the linked pivot(s) and is written back into the cache part. */
export interface SheetTimeline {
  name: string;
  cache: string;
  caption?: string;
  /** Granularity: 0=years 1=quarters 2=months 3=days (CT_Timeline @level). */
  level?: string;
  /** The pivot cache field (a date column) this timeline filters. */
  sourceName: string;
  pivotTables: string[];
  /** ISO datetimes: the selected range, and the full range the data spans. */
  startDate?: string;
  endDate?: string;
  boundStart?: string;
  boundEnd?: string;
  anchor?: import("./chart-model").ChartAnchor;
  timelinePath?: string;
  cachePath?: string;
  dirty?: boolean;
}

/** A data-validation authoring spec (a DataValidation without its ranges), shared by both writers. */
export type DvSpec = { type?: DataValidation["type"]; operator?: DataValidation["operator"]; formula1?: string; formula2?: string; values?: string[]; rangeRef?: string; allowBlank?: boolean };

/** A differential format (dxf) a conditional-formatting rule applies when it matches. */
export interface CfDxf { bg?: string; color?: string; bold?: boolean; italic?: boolean; }

/** One conditional-formatting rule (a subset of the xlsx cfRule types). */
export interface CfRule {
  priority: number;
  stopIfTrue?: boolean;
  type: string; // cellIs, containsText, notContainsText, beginsWith, endsWith, top10, aboveAverage, duplicateValues, uniqueValues, colorScale, dataBar
  operator?: string;
  formulas?: string[];
  text?: string;
  dxf?: CfDxf;
  rank?: number; // top10
  percent?: boolean;
  bottom?: boolean;
  aboveAverage?: boolean; // aboveAverage direction
  equalAverage?: boolean;
  colorScale?: { cfvo: { type: string; val?: number }[]; colors: string[] };
  dataBar?: { color: string; min: { type: string; val?: number }; max: { type: string; val?: number } };
  iconSet?: { set: string; cfvo: { type: string; val?: number; gte?: boolean }[]; reverse?: boolean };
  timePeriod?: string; // timePeriod rule: today, yesterday, last7Days, thisMonth, ...
}

export interface CondFormat { ranges: { r1: number; c1: number; r2: number; c2: number }[]; rules: CfRule[]; }

/** A field that appears in a data-field slot (Sum of Sales, ...) with its aggregation. */
export interface PivotDataField { name: string; func?: string; }

/** A pivot table / data-pilot detected in a workbook. Read-only: the output is rendered as the
    ordinary cells it materialises; this describes the definition so the UI can surface it and so
    an edit to the source range can flag the cache for refresh. Format-agnostic. */
export interface PivotTableInfo {
  name: string;
  /** Source data location, when it is a worksheet range (sheet name + 1-based inclusive range). */
  sourceSheet?: string;
  sourceRange?: { r1: number; c1: number; r2: number; c2: number };
  /** Output range on the host sheet (1-based inclusive), when known. */
  targetRange?: { r1: number; c1: number; r2: number; c2: number };
  rowFields: string[];
  colFields: string[];
  pageFields: string[];
  dataFields: PivotDataField[];
  /** Runtime (set when authored in-app): the spec that built this pivot, so it can be edited and
      refreshed. Absent for pivots read from a file. */
  authorSpec?: import("./pivot").PivotSpec;
  /** xlsx part paths for this pivot (its pivotTable + backing cache definition), for edit/delete. */
  part?: string;
  cachePart?: string;
  /** Host sheet name (where the output lives), so edit/refresh can find it. */
  hostSheet?: string;
}

/** xlsx-only: a pivot cache backed by a worksheet range, tracked so a source-data edit can set
    refreshOnLoad on its definition part (Excel then recomputes the pivot when the file opens). */
export interface PivotCacheRef {
  part: string; // e.g. "xl/pivotCache/pivotCacheDefinition1.xml"
  sourceSheet: string;
  source: { r1: number; c1: number; r2: number; c2: number };
  refreshFlagged?: boolean; // set once we have written refreshOnLoad, to avoid re-serialising
}

/** A phonetic-guide (furigana) run: the reading shown above base[sb..eb). */
export interface Phonetic {
  sb: number; // start base-char offset (inclusive)
  eb: number; // end base-char offset (exclusive)
  reading: string;
}

export interface Cell {
  row: number; // 1-based
  col: number; // 1-based
  /** Canonical/editable value: the literal, or the cached result for a formula cell. */
  value: string;
  /** Phonetic guide (furigana) runs for a Japanese string cell, shown as ruby in the grid;
      the value stays the base text. Preserved untouched in the file's shared/inline string. */
  phonetic?: Phonetic[];
  /** Rich text: per-run styling for a multi-format string cell. Rendered as styled spans; the
      value stays the concatenated text. Preserved via the untouched shared/inline string. */
  richRuns?: TextRun[];
  /** Serialization hint for `value`. */
  kind: CellKind;
  /** Formatted text for the grid (number format applied). Falls back to `value`. */
  display?: string;
  /** Number format for this cell: an xlsx format code, a built-in numFmtId, or none. */
  numFmt?: string | number;
  /** Formula text in A1 syntax, without the leading "=". Undefined if not a formula. */
  formula?: string;
  /** Legacy array formula: the A1 spill range this formula fills (xlsx <f t="array" ref>,
      ODF matrix span). Only the top-left cell carries the formula + this ref. */
  arrayRef?: string;
  /** This cell is filled by another cell's array formula (a spill result), not its own. */
  spill?: boolean;
  /** Transient (never saved): for a dynamic-array anchor, the bottom-right cell of the
      range its result last spilled into, so a later recalc can clear a stale/shrunk spill. */
  dynSpill?: { r: number; c: number };
  /** xlsx: the <c> element in the worksheet DOM (for surgical edits). */
  el?: Element;
  /** ods: the original <table:table-cell> element (cloned verbatim if untouched). */
  odfFormula?: string;
  /** Comments / notes on this cell (xlsx legacy comments + threaded comments), newest last.
      Shown as a marker with a hover popover; preserved via the untouched comment parts. */
  comments?: { author?: string; text: string }[];
  /** Hyperlink on this cell (xlsx <hyperlinks>): an external URL, or an internal
      "Sheet!A1" / defined-name target. Rendered as a clickable link; preserved via the
      untouched sheet XML (sheetedit reads and follows it, it does not rewrite it). */
  link?: { href: string; internal?: boolean; tip?: string };
  /** xlsx @s style index / ods @table:style-name, preserved across edits. */
  style?: string;
  /** Resolved visual formatting (fonts/fills/borders/alignment) for the grid. */
  cellStyle?: CellStyle;
  /** User changed the value/formula (forces regeneration). */
  edited?: boolean;
  /** ods: the hyperlink / note / data-validation was authored on this cell (patched on write,
      preserving the rest of the original cell element instead of rebuilding it from scratch). */
  linkDirty?: boolean;
  commentsDirty?: boolean;
  dvDirty?: boolean;
  /** Recalc changed the cached value. */
  recomputed?: boolean;
  /** xlsx: shared-formula group (@si) this cell belongs to. */
  sharedSi?: string;
  /** The formula itself was added/changed/removed (not just its cached value). */
  fDirty?: boolean;
  /** Transient (never saved): the last recalc could not honestly evaluate this
      formula. "name" = unknown function (typo), "eval" = parse/evaluation
      failure, "circular" = the cell takes part in a reference cycle. */
  calcFailed?: "name" | "eval" | "circular";
  /** numFmt changed (typed date / picker) and must be persisted to styles on save. */
  numFmtDirty?: boolean;
  /** ods: original office:value-type ("date", "time", "percentage", "currency"),
      preserved so an edit keeps the cell's type. */
  odsValueType?: string;
  /** ods: office:currency code ("EUR", ...) for currency cells. */
  odsCurrency?: string;
  /** ods: the table:content-validation-name this cell references, re-emitted on write. */
  odsValidationName?: string;
}

export interface Sheet {
  name: string;
  /** The VBA code name (xlsx `<sheetPr codeName>`), which is how a sheet's macro module is named
      and the only way to tie a Worksheet_Change handler to the sheet it belongs to. */
  codeName?: string;
  /** Tab visibility (xlsx `<sheet state>`, ODF `table:display`). Absent = visible. Excel's
      "very hidden" cannot be undone from the UI, only by a macro, which is the point of it. */
  visibility?: "hidden" | "veryHidden";
  /** The visibility changed in the UI -> the workbook part is rewritten on save. */
  visibilityDirty?: boolean;
  cells: Map<string, Cell>;
  maxRow: number; // 1-based extent of used cells (0 = empty)
  maxCol: number;
  /** 1-based column -> width in px (from the file's <cols>), when specified. */
  colWidths?: Map<number, number>;
  /** Pixels per character unit for this sheet's column widths: the workbook normal font's
      maximum digit width. Absent = Calibri 11's 7. */
  maxDigitWidth?: number;
  /** 1-based row -> height in px (from the file's <row ht>), when specified. */
  rowHeights?: Map<number, number>;
  /** The sheet's own defaults for rows/columns with no explicit size, in px (xlsx
      <sheetFormatPr defaultRowHeight/defaultColWidth>). Absent = the grid's own. */
  defaultRowHeight?: number;
  defaultColWidth?: number;
  /** 1-based rows / columns hidden in the file (<row hidden> / <col hidden> / ODF
      table:visibility). Rendered collapsed (zero size); preserved on save. */
  hiddenRows?: Set<number>;
  hiddenCols?: Set<number>;
  /** Outline (grouping) levels, 1-based index -> depth 1..7. A row/column absent from the map is at
      level 0, i.e. in no group. Rendered as collapsible bars in the gutter. */
  rowOutline?: Map<number, number>;
  colOutline?: Map<number, number>;
  /** The rows/columns carrying a collapsed group's summary marker (xlsx @collapsed). A group is
      shown collapsed when its members are hidden, which is what the file actually records. */
  rowCollapsed?: Set<number>;
  colCollapsed?: Set<number>;
  /** Where a group's summary line sits, from <outlinePr>. Excel defaults both to true (below /
      right); the toggle button is drawn on the summary line. */
  summaryBelow?: boolean;
  summaryRight?: boolean;
  /** Outline state changed in the UI -> the row/col attributes are rewritten on save. */
  outlineDirty?: boolean;
  /** Merged ranges (1-based, inclusive); the top-left cell holds the value. */
  merges?: { r1: number; c1: number; r2: number; c2: number }[];
  /** List-type data validations (dropdowns), matched to cells by range at render time. */
  validations?: DataValidation[];
  /** Conditional-formatting blocks (rules + the ranges they cover). */
  condFormats?: CondFormat[];
  /** Charts anchored on this sheet (rendered as an overlay; created/edited ones are written back). */
  charts?: import("./chart-model").ChartModel[];
  /** Embedded pictures anchored on this sheet (rendered on an overlay; move/resize is written back
      for xlsx, otherwise preserved on save). */
  images?: SheetImage[];
  /** Drawing shapes (rect / ellipse / line / ...) anchored on this sheet; rendered as an SVG overlay,
      authored + written back. */
  shapes?: SheetShape[];
  /** Slicers (interactive pivot filters) on this sheet; rendered as a clickable overlay. */
  slicers?: SheetSlicer[];
  /** Timelines (date-range pivot filters) on this sheet; rendered as a range overlay. */
  timelines?: SheetTimeline[];
  /** Form controls (checkboxes, dropdowns, spinners, buttons) over this sheet; rendered as an
      interactive overlay, with a state change written back into their parts. */
  controls?: SheetControl[];
  /** Sparklines (in-cell mini charts) hosted on this sheet; rendered in the host cell, preserved. */
  sparklines?: { type: "line" | "column" | "stacked"; color: string; negColor?: string; host: { r: number; c: number }; dataRef: string }[];
  /** Frozen panes: count of frozen leading rows / columns (from the file's <pane> /
      view settings). Rendered as sticky; preserved on save via the untouched view XML. */
  freeze?: { rows: number; cols: number };
  /** The freeze changed in the UI -> the view settings are rewritten on save. */
  freezeDirty?: boolean;
  /** The pane boundary is a draggable SPLIT rather than a frozen one. Both render the same way
      here (the leading pane stays put); the difference is what the file records. */
  paneSplit?: boolean;
  /** Page setup for printing (paper, margins, header/footer, print area, page breaks). Not used to
      render the grid, but shown as a print-area / page-break overlay and written back per format. */
  printSetup?: import("./print").PrintSetup;
  /** The print setup changed in the UI -> it is rewritten on save. */
  printDirty?: boolean;
  /** Sheet protection: whether the sheet is protected and which actions it blocks. Enforced in the
      grid (locked cells become read-only) and written back per format. */
  protection?: import("./protection").SheetProtection;
  /** The protection changed in the UI -> it is rewritten on save. */
  protectionDirty?: boolean;
  /** Autofilter range (1-based inclusive; first row = header). Drives sort/filter UI. */
  autoFilter?: { r1: number; c1: number; r2: number; c2: number };
  /** Pivot tables / data-pilots whose output lands on this sheet (read-only; rendered as cells). */
  pivotTables?: PivotTableInfo[];
  /** Runtime: per-column (1-based) set of allowed display values; a column absent from the map
      passes everything. Recomputed into filterHidden. */
  filters?: Map<number, Set<string>>;
  /** Runtime: rows hidden by the active filter (in addition to file-hidden rows). */
  filterHidden?: Set<number>;
  // xlsx
  doc?: Document;
  sheetData?: Element;
  path?: string;
  /** Column widths or row heights changed and the sheet XML must be re-serialized. */
  layoutDirty?: boolean;
  // ods
  tableEl?: Element;
  /** ods: 1-based row -> its table:style-name, preserved/updated so writeOds can re-emit it. */
  odsRowStyles?: Map<number, string>;
  /** ods: original <table:table-row> per expanded row index, so row attributes
      (visibility, default-cell-style-name, ...) survive a rebuild. */
  odsRowEls?: Map<number, Element>;
  /** ods: repeated content rows beyond REPEAT_CAP, preserved verbatim as runs. */
  odsRowRuns?: { from: number; to: number; el: Element }[];
  /** ods: original covered-table-cell elements that carry content or attributes, by "r:c". */
  odsCoveredEls?: Map<string, Element>;
  /** ods: expanded row range of the original table:table-header-rows group. */
  odsHeaderRows?: { el: Element; from: number; to: number };
  /** ods: rows must be re-emitted (cell/merge/row-height edits); untouched sheets keep their XML. */
  odsDirty?: boolean;
  /** ods: this sheet's sparklines were authored/edited/removed, so writeOdsSparklines rebuilds its
      calcext:sparkline-groups from the model (untouched sheets keep their original group verbatim). */
  sparklinesDirty?: boolean;
  /** csv: each physical row's original text (without terminator) + terminator,
      for byte-exact re-emission of untouched rows (index 0 = row 1). `width` is
      the parsed field count, so trailing empty fields survive a rewrite. */
  csvRows?: { raw: string; terminator: string; dirty: boolean; width: number }[];
}

export interface Workbook {
  kind: "xlsx" | "ods" | "csv";
  sheets: Sheet[];
  files: Record<string, Uint8Array>;
  contentDoc?: Document; // ods
  contentPath?: string; // ods
  stylesDoc?: Document; // xlsx xl/styles.xml, kept for style writes
  stylesDirty?: boolean; // xlsx styles.xml changed and must be re-serialized
  /** csv: the sniffed (or hinted) field delimiter. */
  csvDelimiter?: string;
  /** Workbook-level defined names -> an A1 reference ("Sheet1!A1:A10", "A1"), so recalc
      can resolve named ranges in formulas. */
  definedNames?: Map<string, string>;
  /** xlsx: pivot caches backed by a worksheet range; a source edit flags one for refresh. */
  pivotCaches?: PivotCacheRef[];
  /** xlsx: user-defined slicer styles from styles.xml, keyed by name. */
  slicerStyles?: Map<string, import("../adapters/xlsx/slicer-style-read").SlicerStyleDef>;
  /** The workbook's theme: the palette its cells reference and the heading / body fonts. */
  theme?: import("./theme").WorkbookTheme;
  /** The theme changed in the UI -> xl/theme/theme1.xml is rewritten on save. */
  themeDirty?: boolean;
  /** The resolved style pool a freshly rendered cell reads (xlsx cellXfs), kept so a theme switch
      can re-resolve the styles of cells that have not been touched. */
  themeStyles?: (CellStyle | undefined)[];
  /** The VBA project, when the file carries one (.xlsm/.xlsb). Read-only: the source is shown so a
      user can see what a workbook does. The bin itself is preserved untouched on save. */
  vba?: import("./vba").VbaProject;
  /** Workbook-level protection (locked sheet set / window layout). */
  protection?: import("./protection").WorkbookProtection;
  /** The workbook protection changed in the UI -> it is rewritten on save. */
  protectionDirty?: boolean;
}

/** A style change to apply to a cell (only the set fields change). */
export interface StyleChange {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number; // points
  fontFamily?: string;
  color?: string; // CSS "#rrggbb" text colour
  bg?: string; // CSS "#rrggbb" fill colour
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
  wrap?: boolean;
  border?: boolean; // all-sides box border on/off (convenience)
  /** Per-side borders to set; each specified side is turned on/off, others kept. */
  borderSides?: { top?: boolean; right?: boolean; bottom?: boolean; left?: boolean };
  /** Lock (true) or unlock (false) the cell for sheet protection. */
  locked?: boolean;
}

/** A styled run within a rich-text (multi-format) cell string. */
export interface TextRun { text: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; size?: number; color?: string; font?: string; }

/** An embedded picture anchored on a sheet. Rendered on an overlay; a move/resize updates `anchor`
    and sets `dirty`, which the writer flushes back (xlsx: the drawing anchor; ods: the draw:frame). */
export interface SheetImage {
  anchor: import("./chart-model").ChartAnchor;
  dataUri: string;
  /** xlsx: the drawing part path + the anchor's index among the drawing's anchors, so a dirty
      move/resize can find and rewrite the exact anchor element on save. */
  drawingPath?: string;
  anchorIndex?: number;
  /** The media part path holding this picture's bytes (xlsx xl/media/..., ods Pictures/...), so a
      replace can swap them. */
  mediaPath?: string;
  /** ods: the draw:frame element (in the persistent contentDoc) and its anchor cell, so a move/
      resize can patch svg:x/y/width/height relative to that cell without re-parenting the frame. */
  odsFrameEl?: Element;
  odsAnchorCol?: number;
  odsAnchorRow?: number;
  /** ods: pixel offset (from the anchor cell) + size committed by a move/resize, applied on save. */
  odsFrame?: { x: number; y: number; w: number; h: number };
  /** Replacement bytes + extension picked in the UI; the writer swaps the media part on save. */
  replaceBytes?: Uint8Array;
  replaceExt?: string;
  /** Moved/resized/replaced in the UI -> written back; otherwise the parts stay verbatim. */
  dirty?: boolean;
}

/**
 * A form control: the checkboxes, dropdowns, spinners and buttons a workbook can put over the grid.
 * Their point is the LINKED CELL: a checkbox writes TRUE/FALSE into it, a dropdown writes the
 * 1-based index of the chosen item, a spinner writes its number. Formulas read that cell, so a
 * control is an input device for the sheet rather than decoration.
 *
 * xlsx keeps the state in a ctrlProps part and mirrors it in the legacy VML drawing that positions
 * the control; the VML is the only source for older files, so both are read and both written back.
 */
export interface SheetControl {
  /** What the control is. Anything unrecognised is rendered as a label so it is never invisible. */
  kind: "checkbox" | "radio" | "dropdown" | "list" | "spin" | "scroll" | "button" | "label" | "groupBox";
  name: string;
  /** The text drawn on the control (from the VML textbox). */
  label?: string;
  /** A1 reference of the cell this control drives ("$B$1"), when it is linked to one. */
  linkedCell?: string;
  /** The macro assigned to the control (VML `<x:FmlaMacro>`), without its workbook prefix. */
  macro?: string;
  /** An ActiveX control rather than a form control. Its state lives in a persisted binary this
      does not write, and its click handler lives in the sheet's own code module. */
  activeX?: boolean;
  /** ActiveX: the control's persisted value, as the file spells it. */
  activeXValue?: string;
  /** ActiveX: the part holding that value, so a change can be written back into it. */
  activeXBinPath?: string;
  /** List controls: the range the items come from, and the 1-based selection (0 = nothing). */
  sourceRange?: string;
  selected?: number;
  /** Checkbox / radio state. */
  checked?: boolean;
  /** Spinner / scrollbar value and bounds. */
  value?: number;
  min?: number;
  max?: number;
  inc?: number;
  anchor?: import("./chart-model").ChartAnchor;
  /** Part paths + shape id, so a state change can be written back into both places. */
  propsPath?: string;
  vmlPath?: string;
  shapeId?: string;
  /** The user changed the state -> the parts are rewritten on save. */
  dirty?: boolean;
}

/** A supported drawing shape geometry. Anything else read from a file maps to "rect" for rendering
    but keeps its original preset name so it round-trips. */
export type ShapeGeom = "rect" | "roundRect" | "ellipse" | "line" | "triangle" | "diamond" | "parallelogram" | "hexagon" | "pentagon" | "star" | "rightArrow";

/** A drawing shape anchored on a sheet. Rendered as an SVG box over the grid; authored + written
    back (xlsx xdr:sp in the drawing part, ods draw:rect/ellipse/line/custom-shape). */
export interface SheetShape {
  geom: ShapeGeom;
  /** The file's original preset geometry name (xlsx prst / ods class), preserved for a faithful
      round-trip of shapes whose geometry we render approximately as `geom`. */
  preset?: string;
  anchor: import("./chart-model").ChartAnchor;
  fill?: string;        // css hex, or undefined = no fill
  stroke?: string;      // outline css hex, or undefined = no outline
  strokeWidth?: number; // px (default 1)
  text?: string;
  textColor?: string;
  /** xlsx: the drawing part path + the anchor's index, to rewrite the exact anchor on save. */
  drawingPath?: string;
  anchorIndex?: number;
  /** ods: the draw shape element (in the persistent contentDoc) + its anchor cell, so a move/resize
      patches svg:x/y/width/height relative to that cell. */
  odsShapeEl?: Element;
  odsAnchorCol?: number;
  odsAnchorRow?: number;
  odsFrame?: { x: number; y: number; w: number; h: number };
  /** Created / moved / resized / restyled in the UI -> written back; otherwise preserved. */
  dirty?: boolean;
  /** A brand-new shape (authored in-app), so the writer emits it from scratch. */
  created?: boolean;
}

export const key = (row: number, col: number): string => `${row}:${col}`;

export function getCell(sheet: Sheet, row: number, col: number): Cell | undefined {
  return sheet.cells.get(key(row, col));
}

export function ensureCell(sheet: Sheet, row: number, col: number): Cell {
  let c = sheet.cells.get(key(row, col));
  if (!c) {
    c = { row, col, value: "", kind: "blank" };
    sheet.cells.set(key(row, col), c);
  }
  if (row > sheet.maxRow) sheet.maxRow = row;
  if (col > sheet.maxCol) sheet.maxCol = col;
  return c;
}

export function noteExtent(sheet: Sheet, row: number, col: number): void {
  if (row > sheet.maxRow) sheet.maxRow = row;
  if (col > sheet.maxCol) sheet.maxCol = col;
}

/** Typed value of a cell for the formula engine: number, boolean, string or null. */
export function typedValue(cell: Cell | undefined): number | boolean | string | null {
  if (!cell) return null;
  if (cell.value === "") return null;
  switch (cell.kind) {
    case "n": {
      const n = Number(cell.value);
      return Number.isFinite(n) ? n : null;
    }
    case "b":
      return cell.value === "TRUE" || cell.value === "1" || cell.value === "true";
    case "blank":
      return null;
    default:
      return cell.value;
  }
}

export const isNumeric = (s: string): boolean => {
  const t = s.trim();
  return t !== "" && /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t);
};

export function numToStr(n: number): string {
  if (Number.isInteger(n)) return String(n);
  return String(parseFloat(n.toPrecision(15)));
}

/** Apply a number format (code or built-in id) to a numeric value via SSF. */
export function formatNumber(fmt: string | number, value: string): string | undefined {
  const n = Number(value);
  if (!Number.isFinite(n)) return undefined;
  try {
    return FormulaParser.SSF.format(fmt, n);
  } catch {
    return undefined;
  }
}

/** Text shown in the grid for a cell: the formatted display, else the raw value. */
export const cellDisplay = (cell: Cell | undefined): string => (cell ? cell.display ?? cell.value : "");

// ---------------------------------------------------------------------------
// A1 helpers
// ---------------------------------------------------------------------------

export function colToLetters(col: number): string {
  let s = "";
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

export function lettersToCol(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function parseA1Ref(ref: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
  if (!m) return null;
  return { col: lettersToCol(m[1]!), row: Number(m[2]) };
}

export const MAX_COL = 16384; // XFD
export const MAX_ROW = 1048576;
export const REF_TOKEN = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\dA-Za-z_(])/g;

export function shiftRefsInChunk(chunk: string, dr: number, dc: number): string {
  return chunk.replace(REF_TOKEN, (m, cAbs: string, letters: string, rAbs: string, digits: string, off: number) => {
    const prev = off > 0 ? chunk[off - 1]! : "";
    if (/[A-Za-z0-9_$.[\]]/.test(prev)) return m; // inside a longer name (LOG10, tax2026, Table1[...])
    const col0 = lettersToCol(letters.toUpperCase());
    const row0 = Number(digits);
    if (col0 > MAX_COL || row0 > MAX_ROW) return m; // not a real cell ref (defined name)
    const col = cAbs ? col0 : col0 + dc;
    const row = rAbs ? row0 : row0 + dr;
    if (col < 1 || row < 1 || col > MAX_COL || row > MAX_ROW) return "#REF!";
    return cAbs + colToLetters(col) + rAbs + String(row);
  });
}

/**
 * Shift the relative A1 references in a formula by (dr, dc): the translation a
 * shared-formula child applies to its master's formula. Skips string literals
 * and quoted sheet names; absolute parts ($) stay put.
 */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  let out = "";
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i]!;
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === ch) {
          if (formula[j + 1] === ch) j += 2; // doubled quote = escaped
          else {
            j++;
            break;
          }
        } else j++;
      }
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    let next = formula.length;
    for (const q of ['"', "'"]) {
      const p = formula.indexOf(q, i);
      if (p !== -1 && p < next) next = p;
    }
    out += shiftRefsInChunk(formula.slice(i, next), dr, dc);
    i = next;
  }
  return out;
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

export const parseXml = (file: Uint8Array): Document => {
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("corrupt workbook XML");
  return doc;
};
// Optional parts (styles, rels): a broken one degrades instead of failing the open.
export const parseXmlOpt = (file: Uint8Array): Document | undefined => {
  try {
    return parseXml(file);
  } catch {
    return undefined;
  }
};

export function serializeXml(doc: Document): Uint8Array {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith("<?xml")) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return strToU8(s);
}

export const firstByLocal = (parent: Element, local: string): Element | undefined => {
  for (const ch of Array.from(parent.children)) if (ch.localName === local) return ch;
  return undefined;
};

export const removeByLocal = (parent: Element, local: string): void => {
  for (const ch of Array.from(parent.children)) if (ch.localName === local) parent.removeChild(ch);
};


// xlsx and ods style writers so both stay consistent; affected borders become black.
export function mergeCellStyle(cur: CellStyle, change: StyleChange): CellStyle {
  const sides = {
    top: !!cur.borders?.top,
    right: !!cur.borders?.right,
    bottom: !!cur.borders?.bottom,
    left: !!cur.borders?.left,
  };
  if (change.border !== undefined) sides.top = sides.right = sides.bottom = sides.left = change.border;
  if (change.borderSides) Object.assign(sides, change.borderSides);
  const any = sides.top || sides.right || sides.bottom || sides.left;
  return {
    bold: change.bold ?? cur.bold,
    italic: change.italic ?? cur.italic,
    underline: change.underline ?? cur.underline,
    // Keep the file's underline flavour unless the underline is explicitly toggled off.
    underlineStyle: change.underline === false ? undefined : cur.underlineStyle,
    strike: change.strike ?? cur.strike,
    fontSize: change.fontSize ?? cur.fontSize,
    fontFamily: change.fontFamily ?? cur.fontFamily,
    color: change.color ?? cur.color,
    bg: change.bg ?? cur.bg,
    align: change.align ?? cur.align,
    valign: change.valign ?? cur.valign,
    wrap: change.wrap ?? cur.wrap,
    // `locked` is stated the way the formats state it; the model keeps the explicit-unlock flag.
    unlocked: change.locked === undefined ? cur.unlocked : !change.locked,
    formulaHidden: cur.formulaHidden,
    borders: any
      ? {
          top: sides.top ? "#000000" : undefined,
          right: sides.right ? "#000000" : undefined,
          bottom: sides.bottom ? "#000000" : undefined,
          left: sides.left ? "#000000" : undefined,
        }
      : undefined,
  };
}
