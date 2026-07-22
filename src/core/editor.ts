import { t } from "./i18n";
import { hasTimeFmt, isDateFmt, isTimeOnlyFmt, serialToEditText } from "./dates";
import { computeFill, type FillSource } from "./fill";
import { createFormulaBar } from "./ui/formulabar";
import { setupFormulaAssist } from "./ui/formula-assist";
import { setupQueryPanel } from "./ui/query-panel";
import { setupQueryEditor } from "./ui/pq-editor";
import { applyQueryResult, loadResultToNewSheet, listWorkbookTables, tableForQuery, touchedPositions, workbookHasQueries } from "../adapters/xlsx/tables";
import { setupFindBar } from "./ui/findbar";
import { buildToolbar, tbIcon } from "./ui/toolbar";
import { setupFloatBar } from "./ui/floatbar";
import { UndoHistory, applyFields, snapFields, type CellFields, type UndoCellChange } from "./history";
import type { Cell, CellStyle, DataValidation, Phonetic, Sheet, StyleChange, Workbook } from "./model";
import { cellDisplay, colToLetters, ensureCell, getCell, key, parseA1Ref } from "./model";
import { setOdsCellNumFmt, setOdsCellStyle, setOdsColWidth, setOdsMerge, setOdsRowHeight } from "../adapters/ods";
import { recalc } from "./recalc";
import { csvToXlsx, writeCsv } from "../adapters/csv";
import { applyLineOp, syncXlsxMerges, type LineOp } from "./structure";
import { addSheet, renameSheet, deleteSheet, moveSheet, sheetsEditable } from "./sheet-ops";
import { computeCondVisuals, type CfVisual } from "../adapters/xlsx/condformat";
import { readWorkbook, setCellInput, writeWorkbookAsync } from "./workbook";
import { unzipAsync } from "./zip";
import { setXlsxCellNumFmt, setXlsxCellStyle, setXlsxColWidth, setXlsxMerge, setXlsxRowHeight } from "../adapters/xlsx";
// ---------------------------------------------------------------------------
// Editor
// ---------------------------------------------------------------------------

export interface SheetEditorOptions {
  onChange?: () => void;
  /** Route plain-text input explicitly (the host knows the .csv/.tsv extension). */
  formatHint?: "csv" | "tsv";
  /** The document's file name; used to name a converted workbook. */
  fileName?: string;
  /** csv mode: receives the "Convert to XLSX" result. Without it, a download starts. */
  onConvert?: (bytes: Uint8Array, name: string) => void;
}
export interface SheetEditor {
  getBytes(): Promise<Uint8Array>;
  /** The serialized text for text-backed workbooks (csv/tsv); null otherwise. */
  getText(): string | null;
  isDirty(): boolean;
  /** Reset the dirty flag after the host has persisted getBytes()/getText(). */
  markClean(): void;
  /** Read a cell's value by A1 reference on the active sheet ("" if empty or out of range). */
  getCellValue(ref: string): string;
  /** Set a cell's value by A1 reference on the active sheet (recalculates and marks dirty). */
  setCellValue(ref: string, value: string): void;
  destroy(): void;
}

export const ROWS_MIN = 24;
export const COLS_MIN = 12;
/** No longer cap rendering (both axes are virtualized); kept for API compatibility. */
export const ROWS_CAP = 5000;
export const ROW_H = 24; // uniform virtual row height (px) unless the sheet overrides a row
export const COL_W = 96; // uniform virtual column width (px) unless the sheet overrides a column
export const OVERSCAN = 15; // rows rendered beyond the viewport on each side
export const OVERSCAN_COLS = 4; // columns rendered beyond the viewport on each side
export const COLS_CAP = 256;
export const ROW_CHUNK = 20; // rows added per "+ Row" click
export const COL_CHUNK = 6; // columns added per "+ Col" click
export const STYLE_ID = "sheetedit-style";

export function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-wrap { display:flex; flex-direction:column; height:100%; background:var(--sheetedit-bg, #1f2227); color:var(--sheetedit-text, #e6e6e6); font:13px system-ui, sans-serif; }
    .sheetedit-toolbar { display:flex; flex-wrap:nowrap; overflow:hidden; align-items:center; gap:5px; padding:5px 8px; background:var(--sheetedit-chrome, #2b2f36); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
    .sheetedit-btn {
      font:inherit; font-size:13px; background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57);
      border-radius:6px; padding:4px 9px; cursor:pointer; min-width:32px; line-height:1.1;
    }
    .sheetedit-btn:hover { background:var(--sheetedit-btn-hover, #454b54); }
    .sheetedit-btn:focus-visible { outline:2px solid var(--sheetedit-accent, #6e7bff); outline-offset:1px; }
    .sheetedit-tb-sep { width:1px; align-self:stretch; background:var(--sheetedit-btn-border, #4a4f57); margin:1px 3px; }
    .sheetedit-color { width:30px; height:28px; padding:0; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px; background:var(--sheetedit-btn, #3a3f47); cursor:pointer; }
    .sheetedit-tb-select {
      font:inherit; font-size:13px; background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57);
      border-radius:6px; padding:4px 4px; cursor:pointer; max-width:64px; height:28px;
    }
    .sheetedit-tb-select:focus-visible { outline:2px solid var(--sheetedit-accent, #6e7bff); outline-offset:1px; }
    .sheetedit-btn svg { display:block; width:16px; height:16px; }
    .sheetedit-table th.colhead, .sheetedit-table th.rownum, .sheetedit-table th.corner { cursor:pointer; }
    .sheetedit-table th.colhead:hover, .sheetedit-table th.rownum:hover, .sheetedit-table th.corner:hover { background:#e3e3e8; }
    .sheetedit-table td.sheetedit-sel input { background:rgba(110,123,255,0.18); }
    /* Furigana: the ruby overlay shows base + reading; the input's own text is hidden (but
       its selection/fill background still shows through) until the cell is focused for editing. */
    .sheetedit-table td.has-ruby { position:relative; }
    .sheetedit-table td.has-ruby:not(:focus-within) input { color:transparent !important; }
    .sheetedit-table td.has-ruby .sheetedit-ruby {
      position:absolute; inset:0; display:flex; align-items:center; padding:0 6px;
      pointer-events:none; overflow:hidden; white-space:nowrap; line-height:1.05;
    }
    .sheetedit-table td.has-ruby .sheetedit-ruby rt { font-size:0.6em; line-height:1; user-select:none; }
    .sheetedit-table td.has-ruby:focus-within .sheetedit-ruby { display:none; }
    /* Wrap: a wrapping display overlay; the single-line input's text is hidden until focus. */
    .sheetedit-table td.has-wrap { position:relative; }
    .sheetedit-table td.has-wrap:not(:focus-within) input { color:transparent !important; }
    .sheetedit-table td.has-wrap .sheetedit-cellwrap {
      position:absolute; inset:0; padding:3px 8px; white-space:pre-wrap; word-break:break-word;
      overflow:hidden; pointer-events:none; line-height:1.3; color:#1a1a1a;
    }
    .sheetedit-table td.has-wrap:focus-within .sheetedit-cellwrap { display:none; }
    /* Hyperlink cells: link-coloured underlined text and a small open button top-right. */
    .sheetedit-table td.has-link { position:relative; }
    .sheetedit-table td.has-link input:not(:focus) { color:#2563eb; text-decoration:underline; }
    .sheetedit-linkbtn { position:absolute; top:1px; right:1px; z-index:3; display:inline-flex; align-items:center; justify-content:center; width:15px; height:15px; padding:0; border:0; border-radius:3px; background:transparent; color:#2563eb; cursor:pointer; opacity:.8; }
    .sheetedit-linkbtn:hover { opacity:1; background:rgba(37,99,235,0.14); }
    /* Data-validation dropdowns: a caret on the cell's right edge (shown on hover/select), and a
       red outline when the value is not in the allowed list. */
    .sheetedit-table td.has-dv { position:relative; }
    .sheetedit-dvbtn { position:absolute; top:0; right:0; bottom:0; z-index:3; visibility:hidden; display:inline-flex; align-items:center; justify-content:center; width:17px; padding:0; border:0; border-left:1px solid #d4d4d8; background:#eef0f4; color:#555; cursor:pointer; }
    .sheetedit-dvbtn:hover { background:#e3e6ec; color:#111; }
    .sheetedit-table td.has-dv:hover .sheetedit-dvbtn, .sheetedit-table td.has-dv:focus-within .sheetedit-dvbtn, .sheetedit-table td.has-dv.sheetedit-sel .sheetedit-dvbtn { visibility:visible; }
    .sheetedit-table td.sheetedit-dv-invalid { box-shadow: inset 0 0 0 2px #e0533d; }
    .sheetedit-dvmenu { min-width:120px; max-height:240px; overflow-y:auto; }
    /* Conditional-formatting data bar: a proportional bar behind the cell text. */
    .sheetedit-table td.has-cfbar { position:relative; }
    .sheetedit-cfbar { position:absolute; left:1px; top:2px; bottom:2px; z-index:0; border-radius:1px; opacity:.85; pointer-events:none; }
    .sheetedit-table td.has-cfbar input { position:relative; z-index:1; background:transparent; }
    /* Comment marker (corner triangle) + hover popover. */
    .sheetedit-table td.has-comment { position:relative; }
    .sheetedit-commark { position:absolute; top:0; right:0; z-index:2; width:0; height:0; border-top:6px solid #d9534f; border-left:6px solid transparent; pointer-events:none; }
    .sheetedit-compop { position:fixed; z-index:40; max-width:260px; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e7eaf0); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 8px 24px rgba(0,0,0,.45); padding:8px 10px; font:12px/1.45 system-ui,sans-serif; }
    .sheetedit-comitem + .sheetedit-comitem { margin-top:7px; padding-top:7px; border-top:1px solid var(--sheetedit-border, #1c1f24); }
    .sheetedit-comauthor { font-weight:600; margin-bottom:2px; }
    .sheetedit-comtext { white-space:pre-wrap; color:var(--sheetedit-muted, #cfd3da); }
    .sheetedit-furi-pop { min-width:180px; gap:6px; }
    .sheetedit-furi-input { font:inherit; font-size:13px; padding:6px 8px; border-radius:5px; border:1px solid var(--sheetedit-btn-border,#4a4f57); background:var(--sheetedit-btn,#3a3f47); color:var(--sheetedit-text,#e6e6e6); }
    .sheetedit-furi-row { display:flex; gap:4px; }
    .sheetedit-furi-row .sheetedit-pop-item { flex:1; text-align:center; }
    .sheetedit-pop { position:fixed; z-index:30; background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:8px; padding:4px; box-shadow:0 6px 18px rgba(0,0,0,0.45); display:flex; flex-direction:column; min-width:130px; }
    .sheetedit-pop-item { font:inherit; font-size:13px; text-align:left; background:transparent; color:var(--sheetedit-text, #e6e6e6); border:0; border-radius:5px; padding:7px 11px; cursor:pointer; }
    .sheetedit-pop-item:hover { background:var(--sheetedit-btn, #3a3f47); }
    /* The grid is a light canvas (like a real spreadsheet) so the file's fills and
       font colours render faithfully and stay readable; the chrome stays dark. */
    .sheetedit-grid { flex:1; min-height:0; overflow:auto; background:#e9e9ec; }
    table.sheetedit-table { border-collapse:collapse; table-layout:fixed; font:13px/1.3 ui-sans-serif, system-ui, sans-serif; }
    .sheetedit-table th, .sheetedit-table td { padding:0; margin:0; }
    .sheetedit-table th { border:1px solid #d4d4d8; }
    /* Cell gridlines as box-shadows (not borders) so the file's own borders sit flush
       against their neighbours and touch, like a real spreadsheet. */
    .sheetedit-table td { background:#fff; box-shadow: inset -1px -1px 0 0 #e3e3e6; }
    .sheetedit-table th {
      position:sticky; top:0; z-index:8; background:#f1f1f4; color:#555; font-weight:600;
      padding:3px 8px; text-align:center; user-select:none;
    }
    .sheetedit-table th.corner { left:0; z-index:9; }
    .sheetedit-table th.rownum { position:sticky; left:0; z-index:5; top:auto; text-align:right; background:#f1f1f4; }
    /* Frozen-pane cells: opaque so scrolled content does not show through the sticky cell. */
    .sheetedit-table td.frz { background:#fff; }
    /* Vertical alignment: table cells align their inline content natively (flex/grid would
       drop display:table-cell and break the row height). The input is inline-block. */
    .sheetedit-table td.va-top { vertical-align:top; }
    .sheetedit-table td.va-bottom { vertical-align:bottom; }
    .sheetedit-table td.va-top > input, .sheetedit-table td.va-bottom > input { display:inline-block; }
    /* Resize grips: a thin strip on the header border, wide enough to grab on touch. */
    .sheetedit-colgrip { position:absolute; top:0; right:-4px; width:9px; height:100%; cursor:col-resize; z-index:4; touch-action:none; }
    .sheetedit-rowgrip { position:absolute; left:0; bottom:-4px; width:100%; height:9px; cursor:row-resize; z-index:4; touch-action:none; }
    .sheetedit-colgrip:hover { box-shadow:inset -2px 0 0 0 var(--sheetedit-accent, #6e7bff); }
    .sheetedit-rowgrip:hover { box-shadow:inset 0 -2px 0 0 var(--sheetedit-accent, #6e7bff); }
    .sheetedit-table input {
      border:0; background:transparent; color:#1a1a1a; font:inherit; padding:3px 8px;
      width:100%; box-sizing:border-box; outline:none;
    }
    .sheetedit-table td.num input { text-align:right; font-variant-numeric:tabular-nums; }
    .sheetedit-table td.sheetedit-fillsrc { position:relative; }
    .sheetedit-fillhandle {
      position:absolute; right:-4px; bottom:-4px; width:8px; height:8px; z-index:5;
      background:var(--sheetedit-accent, #6e7bff); border:1px solid #fff; cursor:crosshair; touch-action:none;
    }
    .sheetedit-table td.sheetedit-fillprev input { background:rgba(110,123,255,0.10); }
    .sheetedit-table td.sheetedit-fillprev { box-shadow: inset 0 0 0 1px var(--sheetedit-accent, #6e7bff); }
    .sheetedit-findbar { display:flex; align-items:center; gap:6px; padding:5px 8px; background:var(--sheetedit-chrome2, #23262c); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
    .sheetedit-findbar[hidden] { display:none; }
    .sheetedit-findbar input { flex:0 1 180px; background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); font:13px system-ui,sans-serif; padding:4px 8px; }
    .sheetedit-findcount { color:var(--sheetedit-muted, #aab2bf); font:12px system-ui,sans-serif; min-width:64px; }
    .sheetedit-table td.sheetedit-calcerr { position:relative; }
    .sheetedit-table td.sheetedit-calcerr::after {
      content:""; position:absolute; top:0; right:0; z-index:1; pointer-events:none;
      border:5px solid transparent; border-top-color:#d33d3d; border-right-color:#d33d3d;
    }
    .sheetedit-table input:focus { box-shadow:inset 0 0 0 2px var(--sheetedit-accent, #6e7bff); background:#eef0ff; }
    .sheetedit-tb-slot { display:inline-flex; align-items:center; gap:5px; }
    .sheetedit-tb-groupmenu { position:absolute; z-index:30; display:flex; align-items:center; gap:5px; padding:6px 8px; background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
    .sheetedit-fxbar { display:flex; align-items:center; gap:6px; padding:4px 8px; background:var(--sheetedit-chrome2, #23262c); border-bottom:1px solid var(--sheetedit-border, #1c1f24); position:relative; }
    .sheetedit-fxref { min-width:52px; text-align:center; font:12px/1.6 ui-monospace,monospace; color:var(--sheetedit-muted, #aab2bf); background:var(--sheetedit-chrome, #2b2f36); border-radius:5px; padding:2px 6px; }
    .sheetedit-fxbtns { position:relative; display:inline-flex; gap:2px; }
    .sheetedit-fxsum { font-weight:700; }
    .sheetedit-fxmenu { position:absolute; top:100%; left:0; z-index:30; display:flex; flex-direction:column; background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); padding:4px; }
    .sheetedit-fxmenu-item { background:none; border:0; color:var(--sheetedit-text, #e7eaf0); text-align:left; padding:5px 12px; font:12px system-ui,sans-serif; cursor:pointer; border-radius:5px; }
    .sheetedit-fxmenu-item:hover { background:var(--sheetedit-btn, #3a4047); }
    .sheetedit-fxinput { flex:1; min-width:60px; background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); font:13px ui-monospace,monospace; padding:4px 8px; }
    .sheetedit-fxbar.is-picking .sheetedit-fxinput { border-color:var(--sheetedit-accent, #4f8ef7); }
    .sheetedit-fxmenu[hidden], .sheetedit-tb-groupmenu[hidden] { display:none; }
    .sheetedit-fxassist { display:inline-flex; align-items:center; justify-content:center; }
    .sheetedit-fxa-pop { position:absolute; z-index:40; width:min(380px,92%); box-sizing:border-box; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e7eaf0); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 10px 34px rgba(0,0,0,.5); padding:12px; display:flex; flex-direction:column; gap:8px; font:13px/1.4 system-ui,sans-serif; }
    .sheetedit-fxa-pop[hidden] { display:none; }
    .sheetedit-fxa-title { font-weight:600; font-size:14px; }
    .sheetedit-fxa-desc { width:100%; box-sizing:border-box; resize:vertical; padding:7px; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:6px; background:var(--sheetedit-chrome2, #23262c); color:var(--sheetedit-text, #e7eaf0); font:inherit; }
    .sheetedit-fxa-progress { color:var(--sheetedit-muted, #aab2bf); font-size:12px; min-height:15px; }
    .sheetedit-fxa-rlabel { display:flex; flex-direction:column; gap:4px; color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
    .sheetedit-fxa-result { width:100%; box-sizing:border-box; padding:6px 8px; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:6px; background:var(--sheetedit-border, #1c1f24); color:var(--sheetedit-text, #e7eaf0); font:13px ui-monospace,monospace; }
    .sheetedit-fxa-actions { display:flex; gap:8px; justify-content:flex-end; }
    .sheetedit-fxa-btn { padding:6px 14px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:7px; background:var(--sheetedit-chrome2, #23262c); color:var(--sheetedit-text, #e7eaf0); font:inherit; cursor:pointer; }
    .sheetedit-fxa-btn:hover:not(:disabled) { border-color:var(--sheetedit-accent, #6e7bff); }
    .sheetedit-fxa-btn.is-primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:#fff; }
    .sheetedit-fxa-btn:disabled { opacity:.45; cursor:default; }
    .sheetedit-qp-pop { position:absolute; z-index:40; width:min(460px,94%); box-sizing:border-box; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e7eaf0); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 10px 34px rgba(0,0,0,.5); padding:12px; display:flex; flex-direction:column; gap:8px; font:13px/1.4 system-ui,sans-serif; max-height:70vh; }
    .sheetedit-qp-pop[hidden] { display:none; }
    .sheetedit-qp-title { font-weight:600; font-size:14px; }
    .sheetedit-qp-body { display:flex; flex-direction:column; gap:8px; overflow:auto; min-height:0; }
    .sheetedit-qp-row { border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; padding:8px; display:flex; flex-direction:column; gap:5px; }
    .sheetedit-qp-rowhead { display:flex; align-items:center; gap:8px; }
    .sheetedit-qp-name { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:600; }
    .sheetedit-qp-btn { padding:4px 11px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px; background:var(--sheetedit-chrome2, #23262c); color:var(--sheetedit-text, #e7eaf0); font:inherit; cursor:pointer; }
    .sheetedit-qp-btn:hover:not(:disabled) { border-color:var(--sheetedit-accent, #6e7bff); }
    .sheetedit-qp-btn:disabled { opacity:.45; cursor:default; }
    .sheetedit-qp-status { color:var(--sheetedit-muted, #aab2bf); font-size:12px; min-height:14px; }
    .sheetedit-qp-m { margin:0; max-height:180px; overflow:auto; padding:7px 9px; border-radius:6px; background:var(--sheetedit-border, #1c1f24); color:var(--sheetedit-text, #e7eaf0); font:12px/1.5 ui-monospace,monospace; white-space:pre-wrap; }
    .sheetedit-qp-note { color:var(--sheetedit-muted, #aab2bf); font-size:11px; }
    .sheetedit-qp-attach { display:inline-block; margin:4px 0; font-size:11px; color:var(--sheetedit-accent, #6cf); cursor:pointer; }
    .sheetedit-qp-medit { display:flex; flex-direction:column; gap:4px; margin-top:4px; }
    .sheetedit-qp-medit button { align-self:flex-start; }
    /* Token/field colours follow the app's light/dark theme. Light is the default; the two
       blocks below flip to a dark palette on prefers-color-scheme:dark or [data-theme=dark]
       (mirroring how the host drives its own theme). Both bg and tokens move together so the
       field is always readable, standalone or embedded. */
    .sheetedit-qp-mwrap {
      --se-code-bg:#ffffff; --se-code-border:#c9ccd4; --se-code-fg:#24292e; --se-code-caret:#111;
      --se-kw:#0000c8; --se-fn:#795e26; --se-str:#a31515; --se-num:#098658; --se-com:#2e8b57; --se-op:#555; --se-id:#001080;
      position:relative; height:170px; border:1px solid var(--se-code-border); border-radius:4px; overflow:hidden; background:var(--se-code-bg);
    }
    .sheetedit-qp-mhl, .sheetedit-qp-medit textarea.sheetedit-qp-m {
      position:absolute; inset:0; margin:0; box-sizing:border-box; padding:6px; border:0;
      font-family:ui-monospace,Menlo,Consolas,monospace; font-size:11px; line-height:1.5;
      white-space:pre; overflow:auto; tab-size:4; -moz-tab-size:4;
    }
    .sheetedit-qp-mhl { color:var(--se-code-fg); pointer-events:none; z-index:0; }
    .sheetedit-qp-mhl code { font:inherit; white-space:inherit; }
    .sheetedit-qp-medit textarea.sheetedit-qp-m { z-index:1; background:transparent; color:transparent; caret-color:var(--se-code-caret); resize:none; outline:none; }
    .sheetedit-qp-mhl .mtok-kw { color:var(--se-kw); font-weight:600; }
    .sheetedit-qp-mhl .mtok-fn { color:var(--se-fn); }
    .sheetedit-qp-mhl .mtok-str { color:var(--se-str); }
    .sheetedit-qp-mhl .mtok-num { color:var(--se-num); }
    .sheetedit-qp-mhl .mtok-com { color:var(--se-com); font-style:italic; }
    .sheetedit-qp-mhl .mtok-op { color:var(--se-op); }
    .sheetedit-qp-mhl .mtok-id { color:var(--se-id); }
    @media (prefers-color-scheme: dark) {
      :root:not([data-theme="light"]) .sheetedit-qp-mwrap {
        --se-code-bg:#1e2228; --se-code-border:#3a3f4b; --se-code-fg:#d6dae0; --se-code-caret:#e6e6e6;
        --se-kw:#7aa2ff; --se-fn:#d7ba7d; --se-str:#ce9178; --se-num:#b5cea8; --se-com:#7fb37f; --se-op:#b0b6c0; --se-id:#9cdcfe;
      }
    }
    :root[data-theme="dark"] .sheetedit-qp-mwrap {
      --se-code-bg:#1e2228; --se-code-border:#3a3f4b; --se-code-fg:#d6dae0; --se-code-caret:#e6e6e6;
      --se-kw:#7aa2ff; --se-fn:#d7ba7d; --se-str:#ce9178; --se-num:#b5cea8; --se-com:#7fb37f; --se-op:#b0b6c0; --se-id:#9cdcfe;
    }
    .sheetedit-fmtmenu { flex-direction:column; align-items:stretch; gap:2px; }
    .sheetedit-fmtmenu .sheetedit-btn { text-align:left; justify-content:flex-start; }
    .sheetedit-floatbar { position:fixed; z-index:40; display:flex; align-items:center; gap:2px; padding:4px 6px; background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
    .sheetedit-floatbar[hidden] { display:none; }
    .sheetedit-error { background:#7a2b2b; color:#ffd7d7; padding:10px 14px; font:13px/1.5 system-ui,sans-serif; }
    .sheetedit-tabs { display:flex; align-items:center; gap:2px; padding:5px 8px; background:var(--sheetedit-chrome, #2b2f36); border-top:1px solid var(--sheetedit-border, #1c1f24); overflow-x:auto; }
    .sheetedit-tab {
      font:inherit; background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-muted, #cfd3da); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-bottom:none;
      border-radius:5px 5px 0 0; padding:4px 12px; cursor:pointer; white-space:nowrap;
    }
    .sheetedit-tab[aria-selected="true"] { background:var(--sheetedit-accent, #6e7bff); color:#fff; border-color:var(--sheetedit-accent, #6e7bff); }
    .sheetedit-tab:focus-visible { outline:2px solid #fff; outline-offset:1px; }
    .sheetedit-tab-rename { font:inherit; width:9ch; min-width:60px; box-sizing:border-box; background:var(--sheetedit-bg, #1f2227); color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-accent, #6e7bff); border-radius:4px; padding:2px 5px; }
    .sheetedit-tab-add { display:inline-flex; align-items:center; justify-content:center; flex:none; width:26px; height:26px; margin-left:4px; padding:0; cursor:pointer; background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-muted, #cfd3da); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:5px; }
    .sheetedit-tab-add:hover { background:var(--sheetedit-btn-hover, #454b54); color:var(--sheetedit-text, #e6e6e6); }
    .sheetedit-tabmenu { bottom:40px; }
  `;
  document.head.appendChild(s);
}

export function createSheetEditor(
  container: HTMLElement,
  bytes: Uint8Array,
  options: SheetEditorOptions = {},
  preunzipped?: Record<string, Uint8Array>,
): SheetEditor {
  const original = bytes.slice();
  let dirty = false;
  injectStyles();

  let wb: Workbook;
  try {
    wb = readWorkbook(bytes, { formatHint: options.formatHint }, preunzipped);
  } catch (e) {
    // A file that cannot be opened must never lead to a blank editable grid
    // overwriting it: show the reason and return the original bytes on save.
    console.warn("[sheetedit] failed to open workbook", e);
    const errWrap = document.createElement("div");
    errWrap.className = "sheetedit-wrap";
    const banner = document.createElement("div");
    banner.className = "sheetedit-error";
    banner.setAttribute("role", "alert");
    banner.textContent = t("openFailed");
    errWrap.appendChild(banner);
    container.appendChild(errWrap);
    return {
      getBytes: () => Promise.resolve(original.slice()),
      getText: () => null,
      isDirty: () => false,
      markClean: () => undefined,
      getCellValue: () => "",
      setCellValue: () => undefined,
      destroy() {
        errWrap.remove();
      },
    };
  }
  // Trust the file's cached results on open (like Excel/LibreOffice); recalc only runs
  // after an edit. Recomputing on load would overwrite valid cached values whose inputs
  // are blank in this session (e.g. a DATEDIF age before a birthdate is entered).

  const wrap = document.createElement("div");
  wrap.className = "sheetedit-wrap";
  const toolbar = document.createElement("div");
  toolbar.className = "sheetedit-toolbar";
  const gridScroll = document.createElement("div");
  gridScroll.className = "sheetedit-grid";
  const tabs = document.createElement("div");
  tabs.className = "sheetedit-tabs";
  tabs.setAttribute("role", "tablist");
  tabs.setAttribute("aria-label", t("sheets"));
  // --- formula bar + range picking state -------------------------------------
  let activeCell: { r: number; c: number } | null = null;
  let barGrab = false; // pointerdown on the bar: skip the cell's blur-commit
  let skipFocusValue = false; // one focus event keeps the input's pending text
  let pickCb: ((ref: string) => void) | null = null;
  const refName = (r: number, c: number) => `${colToLetters(c)}${r}`;
  const rangeRef = (g: { r1: number; c1: number; r2: number; c2: number }) =>
    g.r1 === g.r2 && g.c1 === g.c2 ? refName(g.r1, g.c1) : `${refName(g.r1, g.c1)}:${refName(g.r2, g.c2)}`;
  const rawOf = (r: number, c: number): string => {
    const live = getCell(wb.sheets[active]!, r, c);
    if (!live) return "";
    if (live.formula != null) return "=" + live.formula;
    // Date cells edit as a date, not as the underlying serial number.
    if (live.kind === "n" && isDateFmt(live.numFmt)) {
      const serial = Number(live.value);
      if (isTimeOnlyFmt(live.numFmt) && serial >= 0 && serial < 1) {
        const secs = Math.round(serial * 86400);
        const p2 = (n: number) => String(n).padStart(2, "0");
        return `${p2(Math.floor(secs / 3600))}:${p2(Math.floor((secs % 3600) / 60))}:${p2(secs % 60)}`;
      }
      const text = serialToEditText(serial, hasTimeFmt(live.numFmt));
      if (text != null) return text;
    }
    return live.value;
  };
  const fxbar = createFormulaBar({
    onInput: (v) => {
      if (!activeCell) return;
      const ip = inputs.get(key(activeCell.r, activeCell.c));
      if (ip) ip.value = v;
    },
    onEnter: (v) => {
      if (!activeCell) return;
      commitValue(activeCell.r, activeCell.c, v);
      fxbar.setValue(rawOf(activeCell.r, activeCell.c));
    },
    onEscape: () => {
      if (pickCb) {
        pickCb = null;
        fxbar.setHint(null);
        return;
      }
      if (!activeCell) return;
      fxbar.setValue(rawOf(activeCell.r, activeCell.c));
      const ip = inputs.get(key(activeCell.r, activeCell.c));
      if (ip) ip.value = displayValue(wb.sheets[active]!, activeCell.r, activeCell.c);
    },
    onFn: (fn) => applyFn(fn),
    onAssist: (anchor) => formulaAssist?.open(anchor),
  });
  // Wired up once commitValue/context helpers below exist (the button only fires on click).
  let formulaAssist: { open(anchor: HTMLElement): void } | null = null;
  fxbar.input.addEventListener("pointerdown", () => {
    barGrab = true;
  });
  fxbar.input.addEventListener("focus", () => {
    setTimeout(() => (barGrab = false), 0);
  });
  if (wb.kind === "csv") wrap.append(toolbar, fxbar.el, gridScroll);
  else wrap.append(toolbar, fxbar.el, gridScroll, tabs);
  container.appendChild(wrap);

  let active = 0;
  let condVisuals = new Map<string, CfVisual>(); // conditional-format visuals for the active sheet, per render
  let inputs = new Map<string, HTMLInputElement>();
  let tds = new Map<string, HTMLElement>();
  // Extra rows/columns the user added beyond the sheet's used extent (per active sheet).
  let extraRows = 0;
  let extraCols = 0;
  // Selection rectangle (1-based, inclusive) and the anchor for shift-extend.
  let sel: { r1: number; c1: number; r2: number; c2: number } | null = null;
  let anchor: { r: number; c: number } | null = null;
  // Grid extent of the last render (whole-line selections are made against it).
  let renderedRows = 0;
  let renderedCols = 0;

  // Fill handle: a corner grip on the selection; dragging extends values,
  // series, patterns and relative formulas (core/fill.ts) in one undo step.
  const fillHandle = document.createElement("div");
  fillHandle.className = "sheetedit-fillhandle";
  let fillPreview: HTMLElement[] = [];
  const clearFillPreview = () => {
    for (const td of fillPreview) td.classList.remove("sheetedit-fillprev");
    fillPreview = [];
  };
  const placeFillHandle = () => {
    fillHandle.remove();
    if (!sel) return;
    const td = tds.get(key(sel.r2, sel.c2));
    if (!td) return;
    td.classList.add("sheetedit-fillsrc");
    td.appendChild(fillHandle);
  };
  const applyFill = (er: number, ec: number) => {
    if (!sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const { r1, c1, r2, c2 } = sel;
    let axis: "row" | "col";
    let dir: 1 | -1;
    let count: number;
    if (er > r2) [axis, dir, count] = ["row", 1, er - r2];
    else if (er < r1) [axis, dir, count] = ["row", -1, r1 - er];
    else if (ec > c2) [axis, dir, count] = ["col", 1, ec - c2];
    else if (ec < c1) [axis, dir, count] = ["col", -1, c1 - ec];
    else return;
    const positions: { r: number; c: number; raw: string }[] = [];
    if (axis === "row") {
      for (let c = c1; c <= c2; c++) {
        const source: FillSource[] = [];
        for (let r = r1; r <= r2; r++) {
          const cell = getCell(sheet, r, c);
          source.push({ value: cell?.value ?? "", formula: cell?.formula, kind: cell?.kind ?? "blank" });
        }
        const vals = computeFill(source, count, dir, "row");
        vals.forEach((raw, n) => positions.push({ r: dir === 1 ? r2 + 1 + n : r1 - 1 - n, c, raw }));
      }
    } else {
      for (let r = r1; r <= r2; r++) {
        const source: FillSource[] = [];
        for (let c = c1; c <= c2; c++) {
          const cell = getCell(sheet, r, c);
          source.push({ value: cell?.value ?? "", formula: cell?.formula, kind: cell?.kind ?? "blank" });
        }
        const vals = computeFill(source, count, dir, "col");
        vals.forEach((raw, n) => positions.push({ r, c: dir === 1 ? c2 + 1 + n : c1 - 1 - n, raw }));
      }
    }
    recordCells(positions.map((p2) => ({ r: p2.r, c: p2.c })), () => {
      for (const p2 of positions) setCellInput(sheet, p2.r, p2.c, p2.raw);
    });
    // Extend the selection over the filled range, like Excel.
    if (axis === "row") sel = { r1: Math.min(r1, dir === 1 ? r1 : r1 - count), c1, r2: Math.max(r2, dir === 1 ? r2 + count : r2), c2 };
    else sel = { r1, c1: Math.min(c1, dir === 1 ? c1 : c1 - count), r2, c2: Math.max(c2, dir === 1 ? c2 + count : c2) };
    recalc(wb);
    mark();
    renderGrid();
  };
  fillHandle.addEventListener("pointerdown", (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (!sel) return;
    let last: { r: number; c: number } | null = null;
    const onMove = (ev: PointerEvent) => {
      const cell = cellAtPoint(ev.clientX, ev.clientY);
      if (!cell || (last && cell.r === last.r && cell.c === last.c) || !sel) return;
      last = cell;
      clearFillPreview();
      // Preview the rectangle that the drop would fill.
      const { r1, c1, r2, c2 } = sel;
      let range: { fr: number; to: number; horizontal: boolean } | null = null;
      if (cell.r > r2) range = { fr: r2 + 1, to: cell.r, horizontal: false };
      else if (cell.r < r1) range = { fr: cell.r, to: r1 - 1, horizontal: false };
      else if (cell.c > c2) range = { fr: c2 + 1, to: cell.c, horizontal: true };
      else if (cell.c < c1) range = { fr: cell.c, to: c1 - 1, horizontal: true };
      if (!range) return;
      for (let i = range.fr; i <= range.to; i++) {
        if (range.horizontal) for (let r = r1; r <= r2; r++) tds.get(key(r, i))?.classList.add("sheetedit-fillprev");
        else for (let c = c1; c <= c2; c++) tds.get(key(i, c))?.classList.add("sheetedit-fillprev");
      }
      fillPreview = [...document.querySelectorAll<HTMLElement>("td.sheetedit-fillprev")];
    };
    const onUp = (ev: PointerEvent) => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      clearFillPreview();
      const cell = cellAtPoint(ev.clientX, ev.clientY);
      if (cell) applyFill(cell.r, cell.c);
    };
    window.addEventListener("pointermove", onMove, { passive: false });
    window.addEventListener("pointerup", onUp);
  });

  const paintSel = () => {
    // Iterate the rendered window's cells (bounded), not the selection rectangle
    // (which can span a million virtual rows after a select-all).
    for (const [k, td] of tds) {
      if (!sel) {
        td.classList.remove("sheetedit-sel");
        continue;
      }
      const [r, c] = k.split(":").map(Number);
      td.classList.toggle("sheetedit-sel", r! >= sel.r1 && r! <= sel.r2 && c! >= sel.c1 && c! <= sel.c2);
      td.classList.remove("sheetedit-fillsrc");
    }
    placeFillHandle();
  };
  const setSel = (r1: number, c1: number, r2: number, c2: number) => {
    sel = { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
    paintSel();
  };
  const selectCell = (r: number, c: number, extend: boolean) => {
    if (extend && anchor) setSel(anchor.r, anchor.c, r, c);
    else {
      anchor = { r, c };
      setSel(r, c, r, c);
    }
  };

  // Rectangular range selection by dragging: mouse drag, or touch long-press then drag.
  // A plain tap still focuses a cell for editing; header/corner taps select a whole line.
  const cellAtPoint = (x: number, y: number): { r: number; c: number } | null => {
    const el = document.elementFromPoint(x, y) as HTMLElement | null;
    const td = el?.closest("td") as HTMLElement | null;
    const rc = td?.dataset.rc;
    if (!rc) return null;
    const [r, c] = rc.split(":").map(Number);
    return { r, c };
  };
  let dragAnchor: { r: number; c: number } | null = null;
  let dragActive = false;
  let justDragged = false;
  let lpTimer: number | null = null;
  let lpStart: { x: number; y: number } | null = null;
  gridScroll.addEventListener("pointerdown", (e) => {
    if (resizing) return;
    const cell = cellAtPoint(e.clientX, e.clientY);
    if (!cell) return;
    if (e.pointerType === "touch") {
      lpStart = { x: e.clientX, y: e.clientY };
      lpTimer = window.setTimeout(() => {
        lpTimer = null;
        dragActive = true;
        dragAnchor = cell;
        anchor = cell;
        (document.activeElement as HTMLElement | null)?.blur?.();
        setSel(cell.r, cell.c, cell.r, cell.c);
      }, 250);
    } else {
      if (pickCb) {
        // Range-pick mode: this gesture only selects the range for the pending
        // formula; preventDefault keeps the edited input focused.
        e.preventDefault();
        dragActive = true;
        anchor = cell;
        dragAnchor = cell;
        setSel(cell.r, cell.c, cell.r, cell.c);
        return;
      }
      dragAnchor = cell; // mouse: a click still edits; a drag selects
    }
  });
  gridScroll.addEventListener(
    "pointermove",
    (e) => {
      if (e.pointerType === "touch") {
        if (lpTimer != null && lpStart) {
          if (Math.abs(e.clientX - lpStart.x) > 8 || Math.abs(e.clientY - lpStart.y) > 8) {
            clearTimeout(lpTimer); // moved before the long-press fired: it is a scroll
            lpTimer = null;
          }
          return;
        }
        if (!dragActive || !dragAnchor) return;
        const cell = cellAtPoint(e.clientX, e.clientY);
        if (cell) {
          e.preventDefault();
          setSel(dragAnchor.r, dragAnchor.c, cell.r, cell.c);
        }
      } else {
        if (!dragAnchor || e.buttons === 0) return;
        const cell = cellAtPoint(e.clientX, e.clientY);
        if (!cell || (cell.r === dragAnchor.r && cell.c === dragAnchor.c && !dragActive)) return;
        if (!dragActive) {
          dragActive = true;
          anchor = dragAnchor;
          (document.activeElement as HTMLElement | null)?.blur?.();
        }
        e.preventDefault();
        setSel(dragAnchor.r, dragAnchor.c, cell.r, cell.c);
      }
    },
    { passive: false },
  );
  const endDrag = () => {
    if (pickCb && dragActive && sel) {
      // Deliver the picked range to the pending formula insertion. No justDragged
      // here: the pick's pointerdown was prevented, so no trailing tap can follow.
      const cb = pickCb;
      pickCb = null;
      const ref = rangeRef(sel);
      dragAnchor = null;
      dragActive = false;
      cb(ref);
      return;
    }
    if (lpTimer != null) {
      clearTimeout(lpTimer);
      lpTimer = null;
    }
    dragAnchor = null;
    if (dragActive) {
      justDragged = true; // swallow the trailing tap so it does not enter edit mode
      setTimeout(() => (justDragged = false), 350);
    }
    dragActive = false;
  };
  window.addEventListener("pointerup", endDrag);
  window.addEventListener("pointercancel", endDrag);

  // Column-width / row-height drag from the header borders. `resizing` suppresses the
  // cell drag-select while a resize is in progress.
  let resizing = false;
  const startColResize = (e: PointerEvent, col: number, colEl: HTMLElement, startW: number) => {
    e.preventDefault();
    e.stopPropagation();
    resizing = true;
    const x0 = e.clientX;
    const onMove = (ev: PointerEvent) => {
      const w = Math.max(24, Math.round(startW + (ev.clientX - x0)));
      colEl.style.width = `${w}px`;
    };
    const onUp = (ev: PointerEvent) => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      const w = Math.max(24, Math.round(startW + (ev.clientX - x0)));
      const sheet = wb.sheets[active];
      if (sheet && wb.kind === "xlsx") {
        setXlsxColWidth(sheet, col, w);
        mark();
      } else if (sheet && wb.kind === "ods") {
        setOdsColWidth(wb, sheet, col, w);
        mark();
      } else if (sheet) {
        (sheet.colWidths ??= new Map()).set(col, w);
      }
      renderGrid();
      setTimeout(() => (resizing = false), 0);
    };
    window.addEventListener("pointermove", onMove, { passive: false });
    window.addEventListener("pointerup", onUp);
  };
  const startRowResize = (e: PointerEvent, row: number, rowEl: HTMLElement, startH: number) => {
    e.preventDefault();
    e.stopPropagation();
    resizing = true;
    const y0 = e.clientY;
    const onMove = (ev: PointerEvent) => {
      const h = Math.max(16, Math.round(startH + (ev.clientY - y0)));
      rowEl.style.height = `${h}px`;
    };
    const onUp = (ev: PointerEvent) => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      const h = Math.max(16, Math.round(startH + (ev.clientY - y0)));
      const sheet = wb.sheets[active];
      if (sheet && wb.kind === "xlsx") {
        setXlsxRowHeight(sheet, row, h);
        mark();
      } else if (sheet && wb.kind === "ods") {
        setOdsRowHeight(wb, sheet, row, h);
        mark();
      } else if (sheet) {
        (sheet.rowHeights ??= new Map()).set(row, h);
      }
      renderGrid();
      setTimeout(() => (resizing = false), 0);
    };
    window.addEventListener("pointermove", onMove, { passive: false });
    window.addEventListener("pointerup", onUp);
  };

  // Style of the selection's top-left cell (used to toggle bold/italic/borders).
  const curStyle = () => (sel ? getCell(wb.sheets[active]!, sel.r1, sel.c1)?.cellStyle : undefined);
  // Apply a style change to every cell in the selection (xlsx only), then re-render.
  // Apply a style change to one cell using the active format's style writer.
  const setCellStyle = (sheet: Sheet, cell: Cell, change: StyleChange) => {
    if (wb.kind === "ods") setOdsCellStyle(wb, sheet, cell, change);
    else setXlsxCellStyle(wb, sheet, cell, change);
  };

  // Cells of the current selection, clamped to the sheet's used extent so styling a whole
  // column / select-all enumerates only real cells (not millions of empty ones); a high cap
  // still guards a pathological explicit range.
  const selPositions = (sheet: Sheet): { r: number; c: number }[] => {
    const out: { r: number; c: number }[] = [];
    if (!sel) return out;
    const r2 = Math.min(sel.r2, Math.max(sel.r1, sheet.maxRow));
    const c2 = Math.min(sel.c2, Math.max(sel.c1, sheet.maxCol));
    for (let r = sel.r1; r <= r2 && out.length < 200000; r++)
      for (let c = sel.c1; c <= c2 && out.length < 200000; c++) out.push({ r, c });
    return out;
  };

  const applyStyle = (change: StyleChange) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const positions = selPositions(sheet);
    recordCells(positions, () => {
      for (const pos of positions) setCellStyle(sheet, ensureCell(sheet, pos.r, pos.c), change);
    });
    mark();
    // Wrap toggles or wrap cells change row heights, so full re-render; everything else patches
    // the rendered cells in place, keeping focus and scroll.
    if (change.wrap !== undefined || positions.some((p) => getCell(sheet, p.r, p.c)?.cellStyle?.wrap)) renderGrid();
    else patchStyle(positions);
  };

  // Apply a number format preset to the selection (General when fmt is undefined).
  const applyNumFmt = (fmt: string | number | undefined, currency?: string) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const positions = selPositions(sheet);
    recordCells(positions, () => {
      for (const pos of positions) {
        const cell = ensureCell(sheet, pos.r, pos.c);
        if (wb.kind === "ods") setOdsCellNumFmt(wb, sheet, cell, fmt, currency);
        else setXlsxCellNumFmt(wb, sheet, cell, fmt);
      }
    });
    mark();
    renderGrid();
  };

  // Furigana authoring: set (or clear, when the reading is empty) the phonetic guide on the
  // active text cell as a single whole-cell run. The write path emits it as rPh / text:ruby.
  const furiTarget = (): { sheet: Sheet; r: number; c: number; cell: Cell } | null => {
    if (wb.kind !== "xlsx" && wb.kind !== "ods") return null;
    const sheet = wb.sheets[active];
    const pos = activeCell ?? (sel ? { r: sel.r1, c: sel.c1 } : null);
    if (!sheet || !pos) return null;
    const cell = getCell(sheet, pos.r, pos.c);
    return cell && cell.value !== "" && cell.kind === "s" ? { sheet, r: pos.r, c: pos.c, cell } : null;
  };
  const applyFurigana = (reading: string): void => {
    const tgt = furiTarget();
    if (!tgt) return;
    recordCells([{ r: tgt.r, c: tgt.c }], () => {
      const rd = reading.trim();
      tgt.cell.phonetic = rd ? [{ sb: 0, eb: tgt.cell.value.length, reading: rd }] : undefined;
      tgt.cell.edited = true;
    });
    mark();
    renderGrid();
  };
  let furiPop: HTMLElement | null = null;
  const openFuriganaPopover = (btn: HTMLElement): void => {
    if (furiPop) {
      furiPop.remove();
      furiPop = null;
      return;
    }
    const tgt = furiTarget();
    if (!tgt) return; // furigana needs a non-empty text cell
    const pop = document.createElement("div");
    pop.className = "sheetedit-pop sheetedit-furi-pop";
    const input = document.createElement("input");
    input.type = "text";
    input.className = "sheetedit-furi-input";
    input.placeholder = t("furiReading");
    input.setAttribute("aria-label", t("furiReading"));
    input.value = tgt.cell.phonetic?.[0]?.reading ?? "";
    const close = () => {
      pop.remove();
      furiPop = null;
      document.removeEventListener("pointerdown", onOutside, true);
    };
    const onOutside = (e: Event) => {
      const n = e.target as Node;
      if (!pop.contains(n) && !btn.contains(n)) close();
    };
    const commit = () => {
      applyFurigana(input.value);
      close();
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        commit();
      } else if (e.key === "Escape") {
        e.preventDefault();
        close();
      }
    });
    const row = document.createElement("div");
    row.className = "sheetedit-furi-row";
    const mkBtn = (label: string, fn: () => void) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-pop-item";
      b.textContent = label;
      b.addEventListener("click", fn);
      return b;
    };
    row.append(
      mkBtn(t("furiSet"), commit),
      mkBtn(t("furiRemove"), () => {
        applyFurigana("");
        close();
      }),
    );
    pop.append(input, row);
    document.body.appendChild(pop);
    furiPop = pop;
    const r = btn.getBoundingClientRect();
    pop.style.left = `${Math.round(r.left)}px`;
    pop.style.top = `${Math.round(r.bottom + 4)}px`;
    input.focus();
    input.select();
    setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
  };

  // Apply a border mode across the selection, computing per-cell sides from each cell's
  // position in the rectangle (e.g. "outer" only borders the perimeter).
  type BorderMode = "all" | "outer" | "top" | "bottom" | "left" | "right" | "none";
  const applyBorder = (mode: BorderMode) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const { r1, c1 } = sel;
    const r2 = Math.min(sel.r2, Math.max(r1, sheet.maxRow)); // clamp to the used extent
    const c2 = Math.min(sel.c2, Math.max(c1, sheet.maxCol));
    const positions = selPositions(sheet);
    recordCells(positions, () => {
      for (const pos of positions) {
        const { r, c } = pos;
        let sides: StyleChange["borderSides"] = {};
        if (mode === "all") sides = { top: true, right: true, bottom: true, left: true };
        else if (mode === "none") sides = { top: false, right: false, bottom: false, left: false };
        else {
          if ((mode === "outer" || mode === "top") && r === r1) sides.top = true;
          if ((mode === "outer" || mode === "bottom") && r === r2) sides.bottom = true;
          if ((mode === "outer" || mode === "left") && c === c1) sides.left = true;
          if ((mode === "outer" || mode === "right") && c === c2) sides.right = true;
        }
        if (Object.keys(sides).length) setCellStyle(sheet, ensureCell(sheet, r, c), { borderSides: sides });
      }
    });
    mark();
    patchStyle(positions); // borders are style-only; patch in place, keeping focus and scroll
  };

  let borderPop: HTMLElement | null = null;
  const openBorderPopover = (btn: HTMLElement) => {
    if (borderPop) {
      borderPop.remove();
      borderPop = null;
      return;
    }
    const pop = document.createElement("div");
    pop.className = "sheetedit-pop";
    pop.setAttribute("role", "menu");
    const close = (refocus = false) => {
      pop.remove();
      borderPop = null;
      document.removeEventListener("pointerdown", onOutside, true);
      if (refocus) btn.focus();
    };
    const onOutside = (e: Event) => {
      const tgt = e.target as Node;
      if (!pop.contains(tgt) && !btn.contains(tgt)) close();
    };
    // Keyboard: Escape closes (returning focus to the trigger); arrows roam the items.
    pop.addEventListener("keydown", (e) => {
      const items = Array.from(pop.querySelectorAll<HTMLButtonElement>(".sheetedit-pop-item"));
      const i = items.indexOf(document.activeElement as HTMLButtonElement);
      if (e.key === "Escape") {
        e.preventDefault();
        close(true);
      } else if (e.key === "ArrowDown") {
        e.preventDefault();
        items[(i + 1) % items.length]?.focus();
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        items[(i - 1 + items.length) % items.length]?.focus();
      }
    });
    const opts: [string, BorderMode][] = [
      [t("borderAll"), "all"],
      [t("borderOuter"), "outer"],
      [t("borderTop"), "top"],
      [t("borderBottom"), "bottom"],
      [t("borderLeft"), "left"],
      [t("borderRight"), "right"],
      [t("borderNone"), "none"],
    ];
    for (const [label, mode] of opts) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-pop-item";
      b.setAttribute("role", "menuitem");
      b.textContent = label;
      b.addEventListener("click", () => {
        applyBorder(mode);
        close(true);
      });
      pop.appendChild(b);
    }
    document.body.appendChild(pop);
    borderPop = pop;
    const r = btn.getBoundingClientRect();
    pop.style.left = `${Math.round(r.left)}px`;
    pop.style.top = `${Math.round(r.bottom + 4)}px`;
    (pop.firstElementChild as HTMLElement | null)?.focus(); // move focus into the menu
    setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
  };

  // Merge the selection into one cell, or unmerge when the selection sits on an
  // existing merge. Merging a region first clears any merges it overlaps.
  const setMerge = (sheet: Sheet, mr1: number, mc1: number, mr2: number, mc2: number, on: boolean) => {
    if (wb.kind === "ods") setOdsMerge(sheet, mr1, mc1, mr2, mc2, on);
    else setXlsxMerge(sheet, mr1, mc1, mr2, mc2, on);
  };
  const toggleMerge = () => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const { r1, c1, r2, c2 } = sel;
    const merges = sheet.merges ?? [];
    const within = (m: { r1: number; c1: number; r2: number; c2: number }) =>
      r1 >= m.r1 && c1 >= m.c1 && r2 <= m.r2 && c2 <= m.c2;
    const intersects = (m: { r1: number; c1: number; r2: number; c2: number }) =>
      !(r2 < m.r1 || r1 > m.r2 || c2 < m.c1 || c1 > m.c2);
    const containing = merges.find(within); // selection inside (or equal to) a merge
    const ops: { r1: number; c1: number; r2: number; c2: number; on: boolean }[] = [];
    if (containing) {
      ops.push({ r1: containing.r1, c1: containing.c1, r2: containing.r2, c2: containing.c2, on: false });
    } else if (r1 !== r2 || c1 !== c2) {
      for (const m of merges.filter(intersects)) ops.push({ r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2, on: false });
      ops.push({ r1, c1, r2, c2, on: true });
    } else {
      return; // a single, unmerged cell: nothing to do
    }
    const positions: { r: number; c: number }[] = [];
    let np = 0;
    for (const o of ops)
      for (let r = o.r1; r <= o.r2 && np < 4000; r++) for (let c = o.c1; c <= o.c2 && np < 4000; c++, np++) positions.push({ r, c });
    const sh = sheet;
    recordCells(
      positions,
      () => {
        for (const o of ops) setMerge(sh, o.r1, o.c1, o.r2, o.c2, o.on);
      },
      {
        undoExtra: () => {
          for (const o of [...ops].reverse()) setMerge(sh, o.r1, o.c1, o.r2, o.c2, !o.on);
        },
        redoExtra: () => {
          for (const o of ops) setMerge(sh, o.r1, o.c1, o.r2, o.c2, o.on);
        },
      },
    );
    mark();
    renderGrid();
  };

  // csv mode: convert the grid into a real workbook and hand it to the host
  // (or download it standalone). Values, formulas and column widths carry over.
  const doConvert = () => {
    const bytes = csvToXlsx(wb);
    const name = (options.fileName ?? "sheet").replace(/\.(csv|tsv|txt)$/i, "") + ".xlsx";
    if (options.onConvert) {
      options.onConvert(bytes, name);
      return;
    }
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  };

  const toolbarHandle = buildToolbar({
    toolbar,
    wrap,
    styled: wb.kind === "xlsx" || wb.kind === "ods",
    convert: wb.kind === "csv" ? doConvert : null,
    onUndo: () => doUndo(),
    onRedo: () => doRedo(),
    addRows: () => {
      extraRows += ROW_CHUNK;
      renderGrid();
    },
    addCols: () => {
      extraCols += COL_CHUNK;
      renderGrid();
    },
    findReplace: () => findBar.toggle(),
    applyStyle,
    applyNumFmt,
    curStyle,
    openBorderPopover,
    toggleMerge,
    editFurigana: openFuriganaPopover,
  });

  // Power Query panel: only when the workbook carries a DataMashup payload. Refresh writes
  // the result through the model (undo-recorded when the target sheet is active) and the
  // table part's @ref is resized by applyQueryResult.
  if (wb.kind === "xlsx" && workbookHasQueries(wb.files)) {
    // Files attached for File.Contents, shared between the quick-refresh panel and the editor.
    const pqFiles: Record<string, Uint8Array> = {};
    const panel = setupQueryPanel({
      wrap,
      wb,
      attachedFiles: pqFiles,
      apply: (target, result, opts) => {
        const run = () => applyQueryResult(wb, target, result);
        // On-open auto-refresh is silent: no undo step and no dirty mark (it wasn't a user edit).
        if (!opts?.silent && target.sheetIndex === active) recordCells(touchedPositions(target, result), run);
        else run();
        recalc(wb);
        if (!opts?.silent) mark();
        renderGrid();
        return { rows: result.rows.length };
      },
      markEdited: () => mark(),
    });
    // "Refresh data when opening the file": auto-refresh the flagged queries now.
    void panel.runOnLoad();
    const QUERY_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><ellipse cx="8" cy="3.5" rx="5" ry="2"/><path d="M3 3.5v4c0 1.1 2.2 2 5 2 .7 0 1.4-.06 2-.16M3 7.5v4c0 1.1 2.2 2 5 2 .5 0 1-.03 1.4-.09"/><path d="M13 8.5v3M11.5 10l1.5 1.5L14.5 10"/></svg>`;
    const qBtn = tbIcon(QUERY_ICON, t("queries"), () => panel.open(qBtn));
    toolbar.append(qBtn);

    // Full Power Query editor (Applied Steps + live preview). Reads/writes Section1.m via qdeff.
    const editor = setupQueryEditor({
      wrap,
      wb,
      attachedFiles: pqFiles,
      save: (newM) => {
        void import("mlang/qdeff").then(({ writeWorkbookSectionM }) => {
          wb.files = writeWorkbookSectionM(wb.files, newM);
          mark();
        });
      },
      loadQuery: (name, result) => {
        // Load into the query's existing destination table if there is one, otherwise onto a
        // fresh sheet. Either way recalc, refresh the tabs and jump to where it landed.
        const existing = tableForQuery(listWorkbookTables(wb), name);
        const sheetIdx = existing ? existing.sheetIndex : loadResultToNewSheet(wb, name, result).sheetIndex;
        if (existing) applyQueryResult(wb, existing, result);
        recalc(wb);
        mark();
        if (sheetIdx === active) { renderTabs(); renderGrid(); }
        else switchSheet(sheetIdx);
        return { sheetName: wb.sheets[sheetIdx]?.name ?? name, rows: result.rows.length };
      },
    });
    const EDIT_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M11 2.5l2.5 2.5M8.5 5L11 2.5 13.5 5 5 13.5H2.5V11z"/></svg>`;
    const eBtn = tbIcon(EDIT_ICON, t("queryEdit"), () => {
      void import("mlang/qdeff").then(({ readWorkbookQueries }) => {
        const q = readWorkbookQueries(wb.files);
        if (q) editor.open(q.mashup.sectionM);
      });
    });
    toolbar.append(eBtn);
  }

  const mark = () => {
    if (!dirty) {
      dirty = true;
    }
    options.onChange?.();
  };

  // Find and replace: finds across every sheet, replaces on the active one.
  const findBar = setupFindBar({
    container: wrap,
    beforeEl: gridScroll,
    getWorkbook: () => wb,
    getActiveSheet: () => active,
    setActiveSheet: (i) => switchSheet(i),
    focusCell: (r, c) => focusCell(r, c),
    commitValue: (r, c, raw) => commitValue(r, c, raw),
    applyBatch: (changes) => {
      const sheet = wb.sheets[active];
      if (!sheet || !changes.length) return;
      recordCells(changes.map((ch) => ({ r: ch.r, c: ch.c })), () => {
        for (const ch of changes) setCellInput(sheet, ch.r, ch.c, ch.raw);
      });
      recalc(wb);
      mark();
      renderGrid();
    },
  });
  wrap.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f") {
      e.preventDefault();
      findBar.show();
    }
  });

  const displayValue = (sheet: Sheet, r: number, c: number): string => cellDisplay(getCell(sheet, r, c));

  const refreshDisplays = (sheet: Sheet, except?: HTMLInputElement) => {
    for (const [k, input] of inputs) {
      if (input === except) continue;
      const [r, c] = k.split(":").map(Number);
      input.value = displayValue(sheet, r!, c!);
      const cell = getCell(sheet, r!, c!);
      input.parentElement?.classList.toggle("num", cell?.kind === "n");
      input.parentElement?.classList.toggle("sheetedit-calcerr", !!cell?.calcFailed);
      if (cell?.calcFailed) input.title = t(cell.calcFailed === "circular" ? "calcCircular" : cell.calcFailed === "name" ? "calcName" : "calcEval");
      else input.removeAttribute("title");
    }
  };

  // --- undo / redo ------------------------------------------------------------
  const history = new UndoHistory();
  const recordCells = (
    positions: { r: number; c: number }[],
    mutate: () => void,
    extra?: { undoExtra?: () => void; redoExtra?: () => void },
  ) => {
    const sheet = wb.sheets[active]!;
    const changes: UndoCellChange[] = positions.map((pos) => ({ r: pos.r, c: pos.c, before: snapFields(getCell(sheet, pos.r, pos.c)), after: null }));
    mutate();
    for (const ch of changes) ch.after = snapFields(getCell(sheet, ch.r, ch.c));
    history.push({ sheet: active, cells: changes, ...extra });
  };
  const applyStep = (step: { sheet: number; cells: UndoCellChange[]; undoExtra?: () => void; redoExtra?: () => void }, dir: "undo" | "redo") => {
    if (step.sheet !== active) {
      active = step.sheet;
      extraRows = 0;
      extraCols = 0;
      sel = null;
      anchor = null;
      renderTabs();
    }
    const sheet = wb.sheets[step.sheet]!;
    for (const ch of step.cells) applyFields(sheet, ch.r, ch.c, dir === "undo" ? ch.before : ch.after);
    (dir === "undo" ? step.undoExtra : step.redoExtra)?.();
    recalc(wb);
    mark();
    renderGrid();
  };
  const doUndo = () => {
    const step = history.popUndo();
    if (step) applyStep(step, "undo");
  };
  const doRedo = () => {
    const step = history.popRedo();
    if (step) applyStep(step, "redo");
  };

  // --- row / column insertion and deletion -------------------------------------
  // A structural op rewrites references workbook-wide and shifts every recorded
  // undo position, so the history is reduced to the op itself: undo replays the
  // inverse shift, then restores dropped cells, formulas, merges and sizes from
  // a snapshot taken here.
  interface LineSnap {
    formulas: { s: number; r: number; c: number; formula: string; odfFormula?: string }[];
    removed: { r: number; c: number; fields: CellFields | null }[];
    merges: { r1: number; c1: number; r2: number; c2: number }[];
    sizes: [number, number][];
  }
  const captureLineSnap = (op: LineOp): LineSnap => {
    const sheet = wb.sheets[active]!;
    const formulas: LineSnap["formulas"] = [];
    wb.sheets.forEach((s, si) => {
      for (const cell of s.cells.values())
        if (cell.formula != null) formulas.push({ s: si, r: cell.row, c: cell.col, formula: cell.formula, odfFormula: cell.odfFormula });
    });
    const removed: LineSnap["removed"] = [];
    if (op.kind === "delete") {
      for (const cell of sheet.cells.values()) {
        const i = op.axis === "row" ? cell.row : cell.col;
        if (i >= op.at && i < op.at + op.count) removed.push({ r: cell.row, c: cell.col, fields: snapFields(cell) });
      }
    }
    const sizeMap = op.axis === "row" ? sheet.rowHeights : sheet.colWidths;
    return {
      formulas,
      removed,
      merges: (sheet.merges ?? []).map((m) => ({ ...m })),
      sizes: sizeMap ? [...sizeMap.entries()] : [],
    };
  };
  const undoLineOp = (si: number, op: LineOp, snap: LineSnap) => {
    const sheet = wb.sheets[si]!;
    const inverse: LineOp = { ...op, kind: op.kind === "insert" ? "delete" : "insert" };
    applyLineOp(wb, si, inverse, false); // formulas come back from the snapshot below
    sheet.merges = snap.merges.map((m) => ({ ...m }));
    if (wb.kind === "xlsx") syncXlsxMerges(sheet);
    else sheet.odsDirty = true;
    for (const rc of snap.removed) if (rc.fields) applyFields(sheet, rc.r, rc.c, rc.fields);
    for (const f of snap.formulas) {
      const cell = getCell(wb.sheets[f.s]!, f.r, f.c);
      if (cell && cell.formula !== f.formula) {
        cell.formula = f.formula;
        cell.odfFormula = f.odfFormula;
        cell.fDirty = true;
        cell.edited = true;
      }
    }
    for (const [i, px] of snap.sizes) {
      const map = op.axis === "row" ? sheet.rowHeights : sheet.colWidths;
      if (map?.get(i) === px) continue;
      if (op.axis === "row") {
        if (wb.kind === "ods") setOdsRowHeight(wb, sheet, i, px);
        else setXlsxRowHeight(sheet, i, px);
      } else {
        if (wb.kind === "ods") setOdsColWidth(wb, sheet, i, px);
        else setXlsxColWidth(sheet, i, px);
      }
    }
  };
  const lineOp = (op: LineOp) => {
    const si = active;
    const snap = captureLineSnap(op);
    applyLineOp(wb, si, op);
    history.clear();
    history.push({ sheet: si, cells: [], undoExtra: () => undoLineOp(si, op, snap), redoExtra: () => applyLineOp(wb, si, op) });
    sel = null;
    anchor = null;
    activeCell = null;
    recalc(wb);
    mark();
    renderGrid();
  };

  // Context menu on a row/column header: insert before/after and delete, acting on
  // the clicked line or on the selected whole-line run that contains it.
  let lineMenu: HTMLElement | null = null;
  const closeLineMenu = () => {
    lineMenu?.remove();
    lineMenu = null;
  };
  const openLineMenu = (e: MouseEvent, axis: "row" | "col", line: number) => {
    e.preventDefault();
    closeLineMenu();
    let base = line;
    let n = 1;
    if (sel) {
      const wholeRows = axis === "row" && sel.c1 === 1 && sel.c2 >= renderedCols && line >= sel.r1 && line <= sel.r2;
      const wholeCols = axis === "col" && sel.r1 === 1 && sel.r2 >= renderedRows && line >= sel.c1 && line <= sel.c2;
      if (wholeRows) {
        base = sel.r1;
        n = sel.r2 - sel.r1 + 1;
      } else if (wholeCols) {
        base = sel.c1;
        n = sel.c2 - sel.c1 + 1;
      }
    }
    const label = (one: string, many: string) => (n === 1 ? t(one) : t(many, { n }));
    const items: [string, LineOp][] = axis === "row"
      ? [
          [label("rowInsAbove", "rowsInsAbove"), { axis, kind: "insert", at: base, count: n }],
          [label("rowInsBelow", "rowsInsBelow"), { axis, kind: "insert", at: base + n, count: n }],
          [label("rowDelOne", "rowsDel"), { axis, kind: "delete", at: base, count: n }],
        ]
      : [
          [label("colInsBefore", "colsInsBefore"), { axis, kind: "insert", at: base, count: n }],
          [label("colInsAfter", "colsInsAfter"), { axis, kind: "insert", at: base + n, count: n }],
          [label("colDelOne", "colsDel"), { axis, kind: "delete", at: base, count: n }],
        ];
    const pop = document.createElement("div");
    pop.className = "sheetedit-pop";
    for (const [text, op] of items) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-pop-item";
      b.textContent = text;
      b.addEventListener("click", () => {
        closeLineMenu();
        lineOp(op);
      });
      pop.appendChild(b);
    }
    document.body.appendChild(pop);
    lineMenu = pop;
    pop.style.left = `${Math.round(Math.min(e.clientX, window.innerWidth - pop.offsetWidth - 8))}px`;
    pop.style.top = `${Math.round(Math.min(e.clientY, window.innerHeight - pop.offsetHeight - 8))}px`;
    const onOutside = (ev: Event) => {
      if (!pop.contains(ev.target as Node)) {
        closeLineMenu();
        document.removeEventListener("pointerdown", onOutside, true);
      }
    };
    setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
  };

  // Commit a raw value into a cell and refresh (shared by the grid and the formula bar).
  const commitValue = (r: number, c: number, raw: string) => {
    const sheet = wb.sheets[active]!;
    recordCells([{ r, c }], () => setCellInput(sheet, r, c, raw));
    recalc(wb);
    mark();
    // A wrap cell's new text may change its row height; a validated cell may gain/lose its
    // invalid flag; a conditional-format edit can recolour the whole range; any of these needs a
    // re-render. Otherwise just refresh displays.
    if (getCell(sheet, r, c)?.cellStyle?.wrap || dvForCell(sheet, r, c) || sheet.condFormats?.length) renderGrid();
    else refreshDisplays(sheet);
  };

  // On-device formula assistant: the fx-bar sparkle opens a popover that turns a plain-language
  // request into a formula (via localml) and, on Insert, writes it into the active cell.
  const assistTarget = () => activeCell ?? (sel ? { r: sel.r1, c: sel.c1 } : { r: 1, c: 1 });
  formulaAssist = setupFormulaAssist({
    wrap,
    getContext: () => {
      const sheet = wb.sheets[active]!;
      const headers: { col: string; name: string }[] = [];
      for (let c = 1; c <= Math.min(sheet.maxCol, 26); c++) {
        const v = getCell(sheet, 1, c)?.value;
        if (v != null && String(v).trim()) headers.push({ col: colToLetters(c), name: String(v).slice(0, 24) });
      }
      const tgt = assistTarget();
      return { cell: `${colToLetters(tgt.c)}${tgt.r}`, headers };
    },
    onAccept: (formula) => {
      const tgt = assistTarget();
      commitValue(tgt.r, tgt.c, formula);
      renderGrid();
      fxbar.setValue(formula);
    },
  });

  // Sigma / function insertion: with a selected row or column run and no edit in
  // progress, write =FN(range) into the cell after the run; otherwise switch to
  // range-pick mode and insert FN(range) at the caret of the pending edit.
  const applyFn = (fn: string) => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const cellInput = activeCell ? inputs.get(key(activeCell.r, activeCell.c)) : undefined;
    const editingBar = document.activeElement === fxbar.input;
    const editingCell = !!cellInput && document.activeElement === cellInput;
    // A selected single-row/column run wins: a multi-cell selection means the user
    // is not mid-edit (shift-click keeps focus on the anchor input).
    const run = sel && (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) && (sel.r1 === sel.r2 || sel.c1 === sel.c2);
    if (run && sel) {
      const ref = rangeRef(sel);
      const vertical = sel.c1 === sel.c2;
      const tr = vertical ? sel.r2 + 1 : sel.r1;
      const tc = vertical ? sel.c1 : sel.c2 + 1;
      commitValue(tr, tc, `=${fn}(${ref})`);
      renderGrid();
      skipFocusValue = false;
      focusCell(tr, tc);
      return;
    }
    const target = activeCell ?? (sel ? { r: sel.r1, c: sel.c1 } : null);
    if (!target) return;
    const editing = editingBar || editingCell;
    const editor = editingBar ? fxbar.input : (inputs.get(key(target.r, target.c)) ?? fxbar.input);
    fxbar.setHint(t("pickRange"));
    pickCb = (ref) => {
      fxbar.setHint(null);
      const base = editing ? editor.value : ""; // not editing: start a fresh formula
      const empty = base.trim() === "" || base.trim() === "=";
      const insert = `${fn}(${ref})`;
      const caret = editing ? (editor.selectionStart ?? base.length) : 0;
      const next = empty ? `=${insert}` : base.slice(0, caret) + insert + base.slice(caret);
      // Focus first (its handler may reset the value), then write the pending text.
      skipFocusValue = true;
      editor.focus();
      skipFocusValue = false;
      editor.value = next;
      fxbar.setValue(next);
      const ip = inputs.get(key(target.r, target.c));
      if (ip && ip !== editor) ip.value = next;
      const pos = (empty ? 1 : caret) + insert.length;
      editor.setSelectionRange(pos, pos);
      activeCell = target;
      fxbar.setRef(refName(target.r, target.c));
    };
  };

  // Clear every cell in a selection rectangle (multi-cell Delete).
  const clearRange = (range: { r1: number; c1: number; r2: number; c2: number }) => {
    const sheet = wb.sheets[active]!;
    // Only existing cells matter (clearing a blank cell is a no-op), so a
    // whole-column selection over a virtualized grid stays bounded.
    const positions: { r: number; c: number }[] = [];
    for (const cell of sheet.cells.values())
      if (cell.row >= range.r1 && cell.row <= range.r2 && cell.col >= range.c1 && cell.col <= range.c2)
        positions.push({ r: cell.row, c: cell.col });
    recordCells(positions, () => {
      for (const pos of positions) setCellInput(sheet, pos.r, pos.c, "");
    });
    recalc(wb);
    mark();
    refreshDisplays(sheet);
  };
  // Copy a multi-cell selection as TSV (a single cell keeps the native copy).
  const copyRange = (e: ClipboardEvent) => {
    if (!sel) return;
    const sheet = wb.sheets[active]!;
    if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
      // Single cell: copy its raw content, unless the user selected text inside
      // the input (then the native text copy is what they asked for).
      const tgt = e.target as HTMLElement | null;
      if (tgt instanceof HTMLInputElement && tgt.selectionStart !== tgt.selectionEnd) return;
      const cell = getCell(sheet, sel.r1, sel.c1);
      e.clipboardData?.setData("text/plain", cell ? (cell.formula != null ? "=" + cell.formula : cell.value) : "");
      e.preventDefault();
      return;
    }
    // Clamp to the used extent so a whole-column copy stays bounded.
    const r2 = Math.min(sel.r2, Math.max(sel.r1, sheet.maxRow));
    const c2 = Math.min(sel.c2, Math.max(sel.c1, sheet.maxCol));
    const lines: string[] = [];
    for (let r = sel.r1; r <= r2; r++) {
      const vals: string[] = [];
      for (let c = sel.c1; c <= c2; c++) vals.push(getCell(sheet, r, c)?.value ?? "");
      lines.push(vals.join("\t"));
    }
    e.clipboardData?.setData("text/plain", lines.join("\n"));
    e.preventDefault();
  };
  // Paste a TSV block (e.g. an Excel range) across cells, starting at the focused one.
  const pasteTsv = (e: ClipboardEvent, r0: number, c0: number) => {
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (!text.includes("\t") && !text.includes("\n")) return; // single value: native paste
    e.preventDefault();
    const sheet = wb.sheets[active]!;
    const rows = text.replace(/\r\n?/g, "\n").split("\n");
    if (rows.length && rows[rows.length - 1] === "") rows.pop(); // Excel's trailing newline
    const grid2 = rows.map((line) => line.split("\t"));
    const positions: { r: number; c: number }[] = [];
    grid2.forEach((vals, i) => vals.forEach((_v, j) => positions.push({ r: r0 + i, c: c0 + j })));
    recordCells(positions, () => {
      grid2.forEach((vals, i) => vals.forEach((val, j) => setCellInput(sheet, r0 + i, c0 + j, val)));
    });
    recalc(wb);
    mark();
    renderGrid(); // the paste may extend the used range; rebuild
    focusCell(r0, c0);
  };

  // --- virtualized grid rendering ---------------------------------------------
  // Both axes are windowed: only the rows and columns intersecting the viewport
  // (plus an overscan) get DOM. Spacer rows/columns keep the scrollbar geometry,
  // so there is no rendering cap in either direction; the only limits left are
  // the file formats' own grid bounds. Offsets are uniform (ROW_H / COL_W) with
  // small sorted indexes of custom-sized lines, so any position is O(log n).
  let totalRows = 0;
  let totalCols = 0;
  let winR1 = 1;
  let winR2 = 0;
  let winC1 = 1;
  let winC2 = 0;
  let heightRows: number[] = [];
  let heightPrefix: number[] = [];
  let widthCols: number[] = [];
  let widthPrefix: number[] = [];
  let tableElRef: HTMLTableElement | null = null;
  let coveredSet = new Set<string>();
  let spanAtMap = new Map<string, { rs: number; cs: number }>();

  // Wrap: computed extra height (px) so a wrapped cell's text fits, measured against the
  // column width. Keyed by row for the active sheet; recomputed on render / resize / edit.
  const wrapH = new Map<number, number>();
  let measureEl: HTMLElement | null = null;
  const measureWrap = (text: string, widthPx: number, cs: CellStyle | undefined): number => {
    if (!measureEl) {
      measureEl = document.createElement("div");
      measureEl.style.cssText =
        "position:absolute;visibility:hidden;left:-9999px;top:0;white-space:pre-wrap;word-break:break-word;padding:3px 8px;box-sizing:border-box;line-height:1.3;font:inherit;";
      gridScroll.appendChild(measureEl);
    }
    measureEl.style.width = `${widthPx}px`;
    measureEl.style.fontWeight = cs?.bold ? "700" : "";
    measureEl.style.fontStyle = cs?.italic ? "italic" : "";
    measureEl.style.fontSize = cs?.fontSize ? `${cs.fontSize}pt` : "";
    measureEl.style.fontFamily = cs?.fontFamily ?? "";
    measureEl.textContent = text || " ";
    return measureEl.offsetHeight;
  };
  const computeWrapHeights = (sheet: Sheet): void => {
    wrapH.clear();
    for (const cell of sheet.cells.values()) {
      if (!cell.cellStyle?.wrap || cell.value === "") continue;
      if (sheet.hiddenRows?.has(cell.row) || sheet.hiddenCols?.has(cell.col)) continue;
      const h = measureWrap(cellDisplay(cell) ?? cell.value, effColW(sheet, cell.col), cell.cellStyle) + 2; // gridline buffer
      if (h > ROW_H) wrapH.set(cell.row, Math.max(wrapH.get(cell.row) ?? 0, h));
    }
  };

  // Effective row height / column width: a hidden line collapses to zero; otherwise the
  // file's custom size, a wrap-grown height, or the default. Drive the size index and the DOM.
  const effRowH = (sheet: Sheet, r: number): number =>
    sheet.hiddenRows?.has(r) ? 0 : Math.max(sheet.rowHeights?.get(r) ?? ROW_H, wrapH.get(r) ?? 0);
  const effColW = (sheet: Sheet, c: number): number => (sheet.hiddenCols?.has(c) ? 0 : (sheet.colWidths?.get(c) ?? COL_W));

  const rebuildSizeIndexes = (sheet: Sheet) => {
    heightRows = [...new Set([...(sheet.rowHeights?.keys() ?? []), ...(sheet.hiddenRows ?? []), ...wrapH.keys()])].sort((a, b) => a - b);
    heightPrefix = [0];
    for (const r of heightRows) heightPrefix.push(heightPrefix[heightPrefix.length - 1]! + (effRowH(sheet, r) - ROW_H));
    widthCols = [...new Set([...(sheet.colWidths?.keys() ?? []), ...(sheet.hiddenCols ?? [])])].sort((a, b) => a - b);
    widthPrefix = [0];
    for (const c of widthCols) widthPrefix.push(widthPrefix[widthPrefix.length - 1]! + (effColW(sheet, c) - COL_W));
  };
  const sumBefore = (lines: number[], prefix: number[], i: number): number => {
    let lo = 0;
    let hi = lines.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (lines[mid]! < i) lo = mid + 1;
      else hi = mid;
    }
    return prefix[lo]!;
  };
  const yOfRow = (r: number): number => (r - 1) * ROW_H + sumBefore(heightRows, heightPrefix, r);
  const xOfCol = (c: number): number => (c - 1) * COL_W + sumBefore(widthCols, widthPrefix, c);
  const lineAt = (pos: number, total: number, ofLine: (i: number) => number): number => {
    let lo = 1;
    let hi = Math.max(1, total);
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (ofLine(mid) <= pos) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };
  const viewportH = (): number => gridScroll.clientHeight || 600; // jsdom has no layout
  const viewportW = (): number => gridScroll.clientWidth || 1200;
  /** Row-number column width: grows with the digit count of the last row. */
  const rnW = (): number => Math.max(44, 18 + String(totalRows).length * 8);

  // A frozen-pane cell stays put when the grid scrolls: sticky to the top (a frozen row),
  // the left (a frozen column), or both (the frozen corner). z-index stacks it under the
  // row/column headers but over ordinary cells.
  const freezeCell = (td: HTMLElement, opts: { left?: number; top?: number; z: number }): void => {
    td.classList.add("frz");
    td.style.position = "sticky";
    if (opts.left != null) td.style.left = `${opts.left}px`;
    if (opts.top != null) td.style.top = `${opts.top}px`;
    td.style.zIndex = String(opts.z);
  };

  // Apply a cell's visual style (fill/borders on the td, font/colour/align on the input) after
  // resetting the style-derived properties, so this both builds a fresh cell and re-styles an
  // existing one in place (a patch that avoids a full renderGrid, keeping focus and scroll).
  const applyCellVisualStyle = (td: HTMLElement, input: HTMLInputElement, cell: Cell | undefined): void => {
    td.style.background = "";
    td.style.boxShadow = "";
    td.classList.remove("va-top", "va-bottom");
    input.style.fontWeight = "";
    input.style.fontStyle = "";
    input.style.textDecoration = "";
    input.style.textDecorationStyle = "";
    input.style.fontSize = "";
    input.style.fontFamily = "";
    input.style.color = "";
    input.style.textAlign = "";
    const cs = cell?.cellStyle;
    if (!cs) return;
    if (cs.bg) td.style.background = cs.bg;
    if (cs.borders) {
      const bd = cs.borders;
      const g = "#e3e3e6";
      const sh = [`inset -1px 0 0 0 ${bd.right ?? g}`, `inset 0 -1px 0 0 ${bd.bottom ?? g}`];
      if (bd.top) sh.push(`inset 0 1px 0 0 ${bd.top}`);
      if (bd.left) sh.push(`inset 1px 0 0 0 ${bd.left}`);
      td.style.boxShadow = sh.join(", ");
    }
    if (cs.bold) input.style.fontWeight = "700";
    if (cs.italic) input.style.fontStyle = "italic";
    if (cs.underline || cs.strike) {
      input.style.textDecoration = `${cs.underline ? "underline" : ""} ${cs.strike ? "line-through" : ""}`.trim();
      if (cs.underline && cs.underlineStyle) input.style.textDecorationStyle = cs.underlineStyle;
    }
    if (cs.fontSize) input.style.fontSize = `${cs.fontSize}pt`;
    if (cs.fontFamily) input.style.fontFamily = cs.fontFamily;
    if (cs.color) input.style.color = cs.color;
    if (cs.align) input.style.textAlign = cs.align;
    if (cs.valign === "top") td.classList.add("va-top");
    else if (cs.valign === "bottom") td.classList.add("va-bottom");
  };

  // Re-style the given cells in place (only those currently rendered), preserving the DOM,
  // focus and scroll. Used for non-geometry style changes instead of a full renderGrid.
  const patchStyle = (positions: { r: number; c: number }[]): void => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    for (const pos of positions) {
      const td = tds.get(key(pos.r, pos.c));
      const input = inputs.get(key(pos.r, pos.c));
      if (td && input) applyCellVisualStyle(td, input, getCell(sheet, pos.r, pos.c));
    }
  };

  // Build one data cell's <td> (input, styles, listeners). Extracted from buildRow so a
  // frozen column can reuse it outside the horizontal window loop.
  const buildCell = (sheet: Sheet, r: number, c: number): HTMLTableCellElement => {
      const td = document.createElement("td");
      td.dataset.rc = key(r, c);
      tds.set(key(r, c), td);
      const sp = spanAtMap.get(key(r, c));
      if (sp) {
        if (sp.rs > 1) td.rowSpan = sp.rs;
        if (sp.cs > 1) td.colSpan = sp.cs;
      }
      const cell = getCell(sheet, r, c);
      if (cell?.kind === "n") td.classList.add("num");
      const input = document.createElement("input");
      input.type = "text";
      input.value = cellDisplay(cell);
      input.setAttribute("aria-label", `${colToLetters(c)}${r}`);
      if (cell?.calcFailed) {
        td.classList.add("sheetedit-calcerr");
        input.title = t(cell.calcFailed === "circular" ? "calcCircular" : cell.calcFailed === "name" ? "calcName" : "calcEval");
      }
      // Apply the file's visual style (fill/borders on the cell, font/colour/align on the text).
      applyCellVisualStyle(td, input, cell);
      // Hyperlink affordance: style the text as a link and add a small open button.
      if (cell?.link) {
        td.classList.add("has-link");
        const lb = document.createElement("button");
        lb.type = "button";
        lb.className = "sheetedit-linkbtn";
        lb.tabIndex = -1;
        lb.title = cell.link.tip || cell.link.href;
        lb.setAttribute("aria-label", t("linkOpen"));
        lb.innerHTML = `<svg viewBox="0 0 16 16" width="11" height="11" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6.5 9.5 13 3M9.5 3H13v3.5M12 9.5V12a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h2.5"/></svg>`;
        lb.addEventListener("mousedown", (e) => e.preventDefault());
        lb.addEventListener("click", (e) => { e.stopPropagation(); openLink(cell.link!); });
        td.appendChild(lb);
      }
      // Data-validation dropdown: a caret that opens the allowed-value list; the cell is flagged
      // when its value is not one of them.
      const dv = dvForCell(sheet, r, c);
      if (dv) {
        td.classList.add("has-dv");
        const cur = cellDisplay(cell);
        const allowed = resolveDvValues(dv, sheet);
        if (cur !== "" && allowed.length && !allowed.includes(cur)) td.classList.add("sheetedit-dv-invalid");
        const caret = document.createElement("button");
        caret.type = "button";
        caret.className = "sheetedit-dvbtn";
        caret.tabIndex = -1;
        caret.title = t("dvChoose");
        caret.setAttribute("aria-label", t("dvChoose"));
        caret.innerHTML = `<svg viewBox="0 0 16 16" width="11" height="11" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 6l4 4 4-4"/></svg>`;
        caret.addEventListener("mousedown", (e) => e.preventDefault());
        caret.addEventListener("click", (e) => { e.stopPropagation(); openDvMenu(td, r, c, dv); });
        td.appendChild(caret);
      }
      // Conditional formatting: override fill/text/bold from the matched rule's dxf, or render a
      // colour-scale fill / data bar.
      const cfv = condVisuals.get(key(r, c));
      if (cfv) {
        if (cfv.bg) td.style.background = cfv.bg;
        if (cfv.color) input.style.color = cfv.color;
        if (cfv.bold) input.style.fontWeight = "bold";
        if (cfv.italic) input.style.fontStyle = "italic";
        if (cfv.bar) {
          td.classList.add("has-cfbar");
          const bar = document.createElement("div");
          bar.className = "sheetedit-cfbar";
          bar.style.width = `${Math.round(cfv.bar.pct * 100)}%`;
          bar.style.background = cfv.bar.color;
          td.insertBefore(bar, td.firstChild);
        }
      }
      // Comments / notes: a corner marker with a hover popover.
      if (cell?.comments?.length) {
        td.classList.add("has-comment");
        const mark = document.createElement("span");
        mark.className = "sheetedit-commark";
        mark.title = cell.comments.map((cm) => (cm.author ? `${cm.author}: ` : "") + cm.text).join("\n");
        td.appendChild(mark);
        td.addEventListener("mouseenter", () => showComment(td, cell.comments!));
        td.addEventListener("mouseleave", hideComment);
      }
      const ki = key(r, c);
      // Shift-click extends the selection from the anchor (no caret/edit).
      input.addEventListener("mousedown", (e) => {
        if (e.shiftKey) {
          e.preventDefault();
          selectCell(r, c, true);
        }
      });
      input.addEventListener("focus", () => {
        if (justDragged) {
          input.blur(); // a range was just drag-selected; do not enter edit on the trailing tap
          return;
        }
        selectCell(r, c, false); // tapping a cell selects it; toolbar styles target the selection
        activeCell = { r, c };
        fxbar.setRef(refName(r, c));
        // Show the editable underlying value (formula or raw), not the formatted display.
        if (skipFocusValue) skipFocusValue = false;
        else input.value = rawOf(r, c);
        fxbar.setValue(input.value);
      });
      input.addEventListener("input", () => fxbar.setValue(input.value));
      let cancelEdit = false;
      const commit = () => {
        // A re-render may have replaced this input while it was focused; a late blur
        // on the stale element must not commit its outdated value.
        if (inputs.get(ki) !== input) return;
        if (barGrab) return; // focus is moving into the formula bar: the edit continues there
        if (cancelEdit) {
          // Escape: discard whatever is in the input, restore the display.
          cancelEdit = false;
          input.value = displayValue(sheet, r, c);
          return;
        }
        const raw = input.value;
        const before = rawOf(r, c);
        if (raw === before) {
          input.value = displayValue(sheet, r, c);
          return;
        }
        commitValue(r, c, raw);
        input.value = displayValue(sheet, r, c);
        input.parentElement?.classList.toggle("num", getCell(sheet, r, c)?.kind === "n");
      };
      input.addEventListener("blur", commit);
      input.addEventListener("keydown", (e) => {
        const moveTo = (rr: number, cc: number) => {
          if (rr < 1 || cc < 1 || rr > totalRows || cc > totalCols) return;
          e.preventDefault();
          input.blur();
          focusCell(rr, cc);
        };
        if (e.key === "Enter") {
          e.preventDefault();
          input.blur();
          if (r + 1 <= totalRows) focusCell(r + 1, c);
        } else if (e.key === "Escape") {
          cancelEdit = true;
          input.blur();
        } else if (e.key === "ArrowDown") moveTo(r + 1, c);
        else if (e.key === "ArrowUp") moveTo(r - 1, c);
        // Left/right leave the cell only from the caret's edge, so in-cell editing stays natural.
        else if (e.key === "ArrowLeft" && input.selectionStart === 0 && input.selectionEnd === 0) moveTo(r, c - 1);
        else if (e.key === "ArrowRight" && input.selectionStart === input.value.length && input.selectionEnd === input.value.length) moveTo(r, c + 1);
        else if (e.key === "Home" && (e.ctrlKey || e.metaKey)) moveTo(1, 1);
        else if ((e.key === "Delete" || e.key === "Backspace") && sel && (sel.r1 !== sel.r2 || sel.c1 !== sel.c2)) {
          e.preventDefault();
          clearRange(sel);
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
          // Mid-edit, native text undo applies; otherwise grid-level undo/redo.
          if (input.value === rawOf(r, c)) {
            e.preventDefault();
            if (e.shiftKey) doRedo();
            else doUndo();
          }
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "y") {
          e.preventDefault();
          doRedo();
        }
      });
      td.appendChild(input);
      // Furigana: render the phonetic guide as ruby in a display overlay. The input keeps the
      // base text (edited/saved as-is); CSS shows the ruby until the cell is focused for editing.
      if (cell?.phonetic?.length) {
        td.classList.add("has-ruby");
        td.appendChild(buildRuby(cellDisplay(cell), cell.phonetic));
      } else if (cell?.cellStyle?.wrap && cell.value !== "") {
        // Wrap: the single-line input can't wrap, so show a wrapping display overlay (hidden
        // while editing). The row was grown to fit by computeWrapHeights.
        td.classList.add("has-wrap");
        const ov = document.createElement("div");
        ov.className = "sheetedit-cellwrap";
        ov.setAttribute("aria-hidden", "true");
        ov.textContent = cellDisplay(cell);
        const cs = cell.cellStyle;
        if (cs.align) ov.style.textAlign = cs.align;
        if (cs.color) ov.style.color = cs.color;
        if (cs.bold) ov.style.fontWeight = "700";
        if (cs.italic) ov.style.fontStyle = "italic";
        if (cs.fontSize) ov.style.fontSize = `${cs.fontSize}pt`;
        if (cs.fontFamily) ov.style.fontFamily = cs.fontFamily;
        td.appendChild(ov);
      }
      inputs.set(ki, input);
      return td;
  };

  // A ruby display element for base text annotated by phonetic runs (base[sb..eb) -> reading).
  const buildRuby = (base: string, runs: Phonetic[]): HTMLElement => {
    const wrap = document.createElement("div");
    wrap.className = "sheetedit-ruby";
    wrap.setAttribute("aria-hidden", "true");
    let pos = 0;
    for (const p of [...runs].filter((x) => x.reading).sort((a, b) => a.sb - b.sb)) {
      const sb = Math.max(pos, Math.min(base.length, p.sb));
      const eb = Math.max(sb, Math.min(base.length, p.eb || base.length));
      if (sb > pos) wrap.appendChild(document.createTextNode(base.slice(pos, sb)));
      const ruby = document.createElement("ruby");
      ruby.appendChild(document.createTextNode(base.slice(sb, eb)));
      const rt = document.createElement("rt");
      rt.textContent = p.reading;
      ruby.appendChild(rt);
      wrap.appendChild(ruby);
      pos = eb;
    }
    if (pos < base.length) wrap.appendChild(document.createTextNode(base.slice(pos)));
    return wrap;
  };

  const buildRow = (sheet: Sheet, r: number, c1: number, c2: number, fz: { fr: number; fc: number; headerH: number }): HTMLTableRowElement => {
    const tr = document.createElement("tr");
    tr.style.height = `${effRowH(sheet, r)}px`;
    const frozenRow = r <= fz.fr;
    const rowTop = fz.headerH + yOfRow(r); // where a frozen row sticks, just below the header
    const rn = document.createElement("th");
    rn.className = "rownum";
    rn.textContent = String(r);
    rn.title = t("selectRow", { row: r });
    rn.addEventListener("click", () => {
      if (resizing) return;
      anchor = { r, c: 1 };
      setSel(r, 1, r, totalCols);
    });
    rn.addEventListener("contextmenu", (e) => openLineMenu(e, "row", r));
    const rgrip = document.createElement("div");
    rgrip.className = "sheetedit-rowgrip";
    rgrip.addEventListener("pointerdown", (e) => startRowResize(e, r, tr, sheet.rowHeights?.get(r) ?? ROW_H));
    rn.appendChild(rgrip);
    if (frozenRow) {
      rn.classList.add("frz");
      rn.style.top = `${rowTop}px`;
      rn.style.zIndex = "7"; // a frozen row's number sticks above its own frozen cells
    }
    tr.appendChild(rn);
    // Frozen columns: always rendered (independent of the horizontal window), sticky-left.
    for (let c = 1; c <= fz.fc; c++) {
      if (sheet.hiddenCols?.has(c) || coveredSet.has(key(r, c))) continue;
      const td = buildCell(sheet, r, c);
      freezeCell(td, { left: rnW() + xOfCol(c), top: frozenRow ? rowTop : undefined, z: frozenRow ? 6 : 3 });
      tr.appendChild(td);
    }
    tr.appendChild(document.createElement("td")); // left spacer column
    for (let c = c1; c <= c2; c++) {
      if (sheet.hiddenCols?.has(c) || coveredSet.has(key(r, c))) continue; // hidden, or part of a merge
      const td = buildCell(sheet, r, c);
      if (frozenRow) freezeCell(td, { top: rowTop, z: 4 });
      tr.appendChild(td);
    }
    tr.appendChild(document.createElement("td")); // right spacer column
    return tr;
  };

  /** Render (or re-render) the visible window: colgroup, header and data rows.
      Explicit coordinates let a full rebuild render for a scroll position the
      empty container cannot hold yet (the browser clamps scrollTop at 0 until
      the spacers exist). */
  let renderingWindow = false;
  const renderWindow = (force = false, yAt?: number, xAt?: number): void => {
    const sheet = wb.sheets[active];
    if (!sheet || !tableElRef || renderingWindow) return;
    renderingWindow = true;
    try {
      renderWindowInner(force, yAt, xAt);
    } finally {
      renderingWindow = false;
    }
  };
  const renderWindowInner = (force: boolean, yAt?: number, xAt?: number): void => {
    const sheet = wb.sheets[active]!;
    const tableEl = tableElRef!;
    const keepTop = gridScroll.scrollTop;
    const keepLeft = gridScroll.scrollLeft;
    const y = yAt ?? keepTop;
    const x = Math.max(0, (xAt ?? keepLeft) - rnW()); // grid area starts after the row-number column
    let r1 = Math.max(1, lineAt(y, totalRows, yOfRow) - OVERSCAN);
    let r2 = Math.min(totalRows, lineAt(y + viewportH(), totalRows, yOfRow) + OVERSCAN);
    let c1 = Math.max(1, lineAt(x, totalCols, xOfCol) - OVERSCAN_COLS);
    let c2 = Math.min(totalCols, lineAt(x + viewportW(), totalCols, xOfCol) + OVERSCAN_COLS);
    // A merge reaching into the window must render whole (its top-left carries the
    // value and the spans), so extend the window over intersecting merges.
    for (const m of sheet.merges ?? []) {
      const intersects = m.r2 >= r1 && m.r1 <= r2 && m.c2 >= c1 && m.c1 <= c2;
      if (!intersects) continue;
      if (m.r1 < r1) r1 = Math.max(1, m.r1);
      if (m.c1 < c1) c1 = Math.max(1, m.c1);
      if (m.r2 > r2) r2 = Math.min(totalRows, m.r2);
      if (m.c2 > c2) c2 = Math.min(totalCols, m.c2);
    }
    if (!force && r1 >= winR1 && r2 <= winR2 && c1 >= winC1 && c2 <= winC2) return;

    // Keep an in-progress edit alive across the re-render when its cell stays
    // near the window; a far-away edit commits (blur) before its DOM goes away.
    let pin: { r: number; c: number; val: string; ss: number | null; se: number | null } | null = null;
    const ae = document.activeElement;
    if (ae instanceof HTMLInputElement && gridScroll.contains(ae)) {
      const rc = ae.closest("td")?.getAttribute("data-rc");
      if (rc) {
        const [pr, pc] = rc.split(":").map(Number);
        pin = { r: pr!, c: pc!, val: ae.value, ss: ae.selectionStart, se: ae.selectionEnd };
        const nearR = pin.r >= r1 - 200 && pin.r <= r2 + 200;
        const nearC = pin.c >= c1 - 50 && pin.c <= c2 + 50;
        if (nearR && nearC) {
          r1 = Math.min(r1, pin.r);
          r2 = Math.max(r2, pin.r);
          c1 = Math.min(c1, pin.c);
          c2 = Math.max(c2, pin.c);
        } else {
          ae.blur();
          pin = null;
        }
      }
    }

    inputs = new Map();
    tds = new Map();
    tableEl.textContent = "";

    // Column skeleton: row numbers, the frozen columns, a left spacer for the window's
    // horizontal offset, the window's columns, a right spacer. Frozen rows/columns render
    // regardless of the scroll position; the window covers only the rest.
    const fr = sheet.freeze?.rows ?? 0;
    const fc = sheet.freeze?.cols ?? 0;
    const ec1 = Math.max(c1, fc + 1); // horizontal window starts after the frozen columns
    const er1 = Math.max(r1, fr + 1); // vertical window starts after the frozen rows
    const gridW = xOfCol(totalCols + 1);
    const leftW = Math.max(0, xOfCol(ec1) - xOfCol(fc + 1)); // gap between frozen cols and window
    const rightW = Math.max(0, gridW - xOfCol(c2 + 1));
    const colgroup = document.createElement("colgroup");
    const addCol = (w: number) => {
      const col = document.createElement("col");
      col.style.width = `${w}px`;
      colgroup.appendChild(col);
      return col;
    };
    // Hidden columns are skipped everywhere (colgroup, header, cells) so they collapse.
    // The col elements are kept by column number for the resize grips.
    const colElByC = new Map<number, HTMLElement>();
    addCol(rnW());
    for (let c = 1; c <= fc; c++) if (!sheet.hiddenCols?.has(c)) colElByC.set(c, addCol(effColW(sheet, c)));
    addCol(leftW);
    for (let c = ec1; c <= c2; c++) if (!sheet.hiddenCols?.has(c)) colElByC.set(c, addCol(effColW(sheet, c)));
    addCol(rightW);
    tableEl.appendChild(colgroup);
    tableEl.style.width = `${rnW() + gridW}px`;

    // A column header cell (letter, select-column click, resize grip).
    const makeColHead = (c: number, colEl: HTMLElement): HTMLTableCellElement => {
      const th = document.createElement("th");
      th.className = "colhead";
      th.textContent = colToLetters(c);
      th.title = t("selectColumn", { col: colToLetters(c) });
      th.addEventListener("click", () => {
        if (resizing) return;
        anchor = { r: 1, c };
        setSel(1, c, totalRows, c);
      });
      th.addEventListener("contextmenu", (e) => openLineMenu(e, "col", c));
      const grip = document.createElement("div");
      grip.className = "sheetedit-colgrip";
      grip.addEventListener("pointerdown", (e) => startColResize(e, c, colEl, sheet.colWidths?.get(c) ?? COL_W));
      th.appendChild(grip);
      return th;
    };

    const head = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "corner";
    corner.title = t("selectAll");
    corner.addEventListener("click", () => {
      anchor = { r: 1, c: 1 }; // a later shift-click extends from A1, not a stale anchor
      setSel(1, 1, totalRows, totalCols);
    });
    head.appendChild(corner);
    for (let c = 1; c <= fc; c++) {
      if (sheet.hiddenCols?.has(c)) continue;
      const th = makeColHead(c, colElByC.get(c)!);
      freezeCell(th, { left: rnW() + xOfCol(c), z: 9 }); // header is already sticky-top
      head.appendChild(th);
    }
    head.appendChild(document.createElement("th")); // left spacer
    for (let c = ec1; c <= c2; c++) if (!sheet.hiddenCols?.has(c)) head.appendChild(makeColHead(c, colElByC.get(c)!));
    head.appendChild(document.createElement("th")); // right spacer
    tableEl.appendChild(head);

    // Frozen rows stick just below the header; measure the header height only when
    // needed (the read forces a layout on the still-empty table).
    const fz = { fr, fc, headerH: fr > 0 ? head.offsetHeight || 0 : 0 };

    const cellCols = fc + (c2 - ec1 + 1) + 3; // rownum + frozen cols + spacer + window + spacer
    for (let r = 1; r <= fr; r++) if (!sheet.hiddenRows?.has(r)) tableEl.appendChild(buildRow(sheet, r, ec1, c2, fz)); // frozen rows
    const topSpacer = document.createElement("tr");
    topSpacer.appendChild(document.createElement("td")).colSpan = cellCols;
    topSpacer.style.height = `${Math.max(0, yOfRow(er1) - yOfRow(fr + 1))}px`;
    tableEl.appendChild(topSpacer);
    for (let r = er1; r <= r2; r++) if (!sheet.hiddenRows?.has(r)) tableEl.appendChild(buildRow(sheet, r, ec1, c2, fz));
    const bottomSpacer = document.createElement("tr");
    bottomSpacer.appendChild(document.createElement("td")).colSpan = cellCols;
    bottomSpacer.style.height = `${Math.max(0, yOfRow(totalRows + 1) - yOfRow(r2 + 1))}px`;
    tableEl.appendChild(bottomSpacer);

    winR1 = r1;
    winR2 = r2;
    winC1 = c1;
    winC2 = c2;

    // Any layout forced while the table was empty clamps the scroll position to 0;
    // the spacers are back, so put it back too.
    if (gridScroll.scrollTop !== keepTop) gridScroll.scrollTop = keepTop;
    if (gridScroll.scrollLeft !== keepLeft) gridScroll.scrollLeft = keepLeft;

    if (pin) {
      const inp = inputs.get(key(pin.r, pin.c));
      if (inp && document.activeElement !== inp) {
        skipFocusValue = true;
        inp.focus({ preventScroll: true }); // focus-scroll would fight the window logic
        skipFocusValue = false;
        inp.value = pin.val;
        fxbar.setValue(pin.val);
        try {
          inp.setSelectionRange(pin.ss, pin.se);
        } catch {
          /* selection restore is best-effort */
        }
      }
    }
    paintSel();
  };

  /** Follow a cell hyperlink: open an external URL in a new tab, or jump to an internal
      "Sheet!A1" / defined-name target. */
  const openLink = (link: { href: string; internal?: boolean }): void => {
    if (!link.internal) {
      window.open(link.href, "_blank", "noopener,noreferrer");
      return;
    }
    let loc = link.href;
    if (wb.definedNames?.has(loc)) loc = wb.definedNames.get(loc)!;
    const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(loc);
    const sheetName = m ? (m[1] ?? m[2]) : undefined;
    const ref = (m ? m[3] : loc).replace(/\$/g, "").split(":")[0];
    if (sheetName) {
      const idx = wb.sheets.findIndex((s) => s.name === sheetName);
      if (idx >= 0) switchSheet(idx);
    }
    const p = parseA1Ref(ref);
    if (p) focusCell(p.row, p.col);
  };

  // Data-validation dropdowns: find the list validation covering a cell, and resolve its
  // allowed values (inline, or from a referenced range read live).
  const dvForCell = (sheet: Sheet, r: number, c: number): DataValidation | null => {
    for (const v of sheet.validations ?? [])
      if (v.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)) return v;
    return null;
  };
  const resolveDvValues = (dv: DataValidation, home: Sheet): string[] => {
    if (dv.values) return dv.values;
    if (!dv.rangeRef) return [];
    let ref = dv.rangeRef;
    if (wb.definedNames?.has(ref)) ref = wb.definedNames.get(ref)!;
    const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(ref);
    const target = m ? wb.sheets.find((s) => s.name === (m[1] ?? m[2])) : home;
    const body = (m ? m[3] : ref).replace(/\$/g, "");
    if (!target) return [];
    const [a, b] = body.split(":");
    const p1 = parseA1Ref(a ?? "");
    const p2 = b ? parseA1Ref(b) : p1;
    if (!p1 || !p2) return [];
    const out: string[] = [];
    for (let r = p1.row; r <= p2.row; r++)
      for (let c = p1.col; c <= p2.col; c++) {
        const cell = getCell(target, r, c);
        const d = cell ? cellDisplay(cell) : "";
        if (d !== "") out.push(d);
      }
    return out;
  };
  function openDvMenu(td: HTMLElement, r: number, c: number, dv: DataValidation): void {
    const values = resolveDvValues(dv, wb.sheets[active]!);
    const menu = document.createElement("div");
    menu.className = "sheetedit-pop sheetedit-dvmenu";
    const add = (label: string, val: string): void => {
      const it = document.createElement("button");
      it.type = "button";
      it.className = "sheetedit-pop-item";
      it.textContent = label;
      it.addEventListener("click", () => { menu.remove(); commitValue(r, c, val); });
      menu.appendChild(it);
    };
    if (dv.allowBlank) add(t("dvBlank"), "");
    for (const v of values) add(v, v);
    if (!values.length && !dv.allowBlank) { const e = document.createElement("div"); e.className = "sheetedit-pop-item"; e.style.opacity = ".6"; e.textContent = t("dvEmpty"); menu.appendChild(e); }
    wrap.appendChild(menu);
    const rect = td.getBoundingClientRect();
    menu.style.left = `${Math.min(rect.left, window.innerWidth - menu.offsetWidth - 6)}px`;
    menu.style.top = `${rect.bottom + 2}px`;
    const close = (e: MouseEvent): void => { if (!menu.contains(e.target as Node)) { menu.remove(); document.removeEventListener("mousedown", close); } };
    setTimeout(() => document.addEventListener("mousedown", close), 0);
  }

  // Comment popover: shown while hovering a cell that carries notes / threaded comments.
  let commentPop: HTMLElement | null = null;
  const hideComment = (): void => { commentPop?.remove(); commentPop = null; };
  const showComment = (td: HTMLElement, comments: { author?: string; text: string }[]): void => {
    hideComment();
    const pop = document.createElement("div");
    pop.className = "sheetedit-compop";
    for (const cm of comments) {
      const block = document.createElement("div");
      block.className = "sheetedit-comitem";
      if (cm.author) { const a = document.createElement("div"); a.className = "sheetedit-comauthor"; a.textContent = cm.author; block.appendChild(a); }
      const txt = document.createElement("div");
      txt.className = "sheetedit-comtext";
      txt.textContent = cm.text;
      block.appendChild(txt);
      pop.appendChild(block);
    }
    wrap.appendChild(pop);
    const rect = td.getBoundingClientRect();
    pop.style.left = `${Math.min(rect.right + 4, window.innerWidth - pop.offsetWidth - 8)}px`;
    pop.style.top = `${Math.min(rect.top, window.innerHeight - pop.offsetHeight - 8)}px`;
    commentPop = pop;
  };

  /** Focus a cell, scrolling it into the rendered window first if needed. */
  const focusCell = (r: number, c: number): void => {
    if (r < 1 || c < 1 || r > totalRows || c > totalCols) return;
    let inp = inputs.get(key(r, c));
    if (!inp) {
      const y = yOfRow(r);
      const x = rnW() + xOfCol(c);
      if (y < gridScroll.scrollTop) gridScroll.scrollTop = y;
      else if (y > gridScroll.scrollTop + viewportH() - ROW_H * 2)
        gridScroll.scrollTop = Math.max(0, y - Math.max(ROW_H * 2, viewportH() - ROW_H * 2));
      if (x < gridScroll.scrollLeft + rnW()) gridScroll.scrollLeft = Math.max(0, x - rnW());
      else if (x > gridScroll.scrollLeft + viewportW() - COL_W)
        gridScroll.scrollLeft = Math.max(0, x - Math.max(COL_W, viewportW() - COL_W * 2));
      renderWindow(true);
      inp = inputs.get(key(r, c));
    }
    inp?.focus({ preventScroll: true });
  };

  // A plain timeout throttle: requestAnimationFrame stalls in occluded windows,
  // which would leave the window stale after a programmatic scroll.
  // Clipboard events land on the focused element and bubble to the document.
  // Grid inputs, and the body after a drag selection blurred the grid, both
  // resolve here; the formula bar keeps its native text behavior.
  const onDocCopy = (e: ClipboardEvent) => {
    const tgt = e.target as HTMLElement | null;
    if (tgt?.closest?.(".sheetedit-fxbar")) return;
    const inWrap = tgt instanceof Node && wrap.contains(tgt);
    if (inWrap || tgt === document.body) copyRange(e);
  };
  const onDocPaste = (e: ClipboardEvent) => {
    const tgt = e.target as HTMLElement | null;
    if (tgt?.closest?.(".sheetedit-fxbar")) return;
    const td = tgt?.closest?.("td[data-rc]") as HTMLElement | null;
    if (td && wrap.contains(td)) {
      const [r, c] = (td.dataset.rc ?? "").split(":").map(Number);
      if (r && c) pasteTsv(e, r, c);
      return;
    }
    if (tgt === document.body && sel) pasteTsv(e, sel.r1, sel.c1);
  };
  document.addEventListener("copy", onDocCopy);
  document.addEventListener("paste", onDocPaste);

  let scrollScheduled = false;
  gridScroll.addEventListener("scroll", () => {
    if (scrollScheduled) return;
    scrollScheduled = true;
    setTimeout(() => {
      scrollScheduled = false;
      renderWindow();
    }, 16);
  });

  const renderGrid = () => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    inputs = new Map();
    tds = new Map();
    const keepTop = gridScroll.scrollTop;
    const keepLeft = gridScroll.scrollLeft;
    gridScroll.innerHTML = "";
    totalRows = Math.max(ROWS_MIN, sheet.maxRow + 6) + extraRows;
    totalCols = Math.max(COLS_MIN, sheet.maxCol + 2) + extraCols;
    renderedRows = totalRows;
    renderedCols = totalCols;
    condVisuals = sheet.condFormats?.length ? computeCondVisuals(sheet) : new Map();
    computeWrapHeights(sheet); // measure wrap cells so rows grow to fit
    rebuildSizeIndexes(sheet);

    // Merged ranges: the top-left cell spans; covered cells are not rendered.
    coveredSet = new Set<string>();
    spanAtMap = new Map<string, { rs: number; cs: number }>();
    for (const m of sheet.merges ?? []) {
      spanAtMap.set(key(m.r1, m.c1), { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
      for (let r = m.r1; r <= m.r2; r++)
        for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) coveredSet.add(key(r, c));
    }

    const table = document.createElement("table");
    table.className = "sheetedit-table";
    tableElRef = table;
    gridScroll.appendChild(table);

    winR1 = 1;
    winR2 = 0;
    winC1 = 1;
    winC2 = 0;
    renderWindow(true, keepTop, keepLeft); // build the window for the kept position first
    gridScroll.scrollTop = keepTop; // now the spacers exist, so the browser keeps it
    gridScroll.scrollLeft = keepLeft;
  };

  const switchSheet = (i: number): void => {
    if (i === active || !wb.sheets[i]) return;
    active = i;
    extraRows = 0; // each sheet starts at its own extent
    extraCols = 0;
    sel = null;
    anchor = null;
    renderTabs();
    renderGrid();
  };

  // Sheet management (xlsx/ods): add, inline-rename, delete and reorder. Structural, so it
  // marks the workbook dirty; there is no undo step (close without saving to discard).
  const canManageSheets = sheetsEditable(wb);

  const doAddSheet = (): void => {
    try {
      const i = addSheet(wb);
      mark();
      switchSheet(i); // renders tabs + grid
      beginRenameTab(tabs.children[i] as HTMLElement, i); // let the user name it immediately
    } catch { /* unsupported file type */ }
  };
  const doDeleteSheet = (i: number): void => {
    if (wb.sheets.length <= 1) return;
    deleteSheet(wb, i);
    if (active === i) active = Math.min(i, wb.sheets.length - 1);
    else if (active > i) active -= 1;
    mark();
    renderTabs();
    renderGrid();
  };
  const doMoveSheet = (from: number, to: number): void => {
    if (to < 0 || to >= wb.sheets.length || from === to) return;
    const cur = wb.sheets[active];
    moveSheet(wb, from, to);
    active = wb.sheets.indexOf(cur);
    mark();
    renderTabs();
  };
  function beginRenameTab(tab: HTMLElement, i: number): void {
    const input = document.createElement("input");
    input.className = "sheetedit-tab-rename";
    input.value = wb.sheets[i]?.name ?? "";
    tab.textContent = "";
    tab.appendChild(input);
    input.focus();
    input.select();
    let done = false;
    const finish = (commit: boolean): void => {
      if (done) return;
      done = true;
      const v = input.value.trim();
      if (commit && v && v !== wb.sheets[i]?.name) { renameSheet(wb, i, v); mark(); }
      renderTabs();
    };
    input.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") { e.preventDefault(); finish(true); }
      else if (e.key === "Escape") { e.preventDefault(); finish(false); }
    });
    input.addEventListener("blur", () => finish(true));
  }
  function showTabMenu(i: number, x: number): void {
    const menu = document.createElement("div");
    menu.className = "sheetedit-pop sheetedit-tabmenu";
    const add = (label: string, fn: () => void, disabled = false): void => {
      const it = document.createElement("button");
      it.type = "button";
      it.className = "sheetedit-pop-item";
      it.textContent = label;
      it.disabled = disabled;
      it.addEventListener("click", () => { menu.remove(); fn(); });
      menu.appendChild(it);
    };
    add(t("sheetRename"), () => beginRenameTab(tabs.children[i] as HTMLElement, i));
    add(t("sheetMoveLeft"), () => doMoveSheet(i, i - 1), i === 0);
    add(t("sheetMoveRight"), () => doMoveSheet(i, i + 1), i === wb.sheets.length - 1);
    add(t("sheetDelete"), () => doDeleteSheet(i), wb.sheets.length <= 1);
    wrap.appendChild(menu);
    menu.style.left = `${Math.min(x, wrap.getBoundingClientRect().width - menu.offsetWidth - 6)}px`;
    menu.style.bottom = "40px";
    const close = (e: MouseEvent): void => { if (!menu.contains(e.target as Node)) { menu.remove(); document.removeEventListener("mousedown", close); } };
    setTimeout(() => document.addEventListener("mousedown", close), 0);
  }

  const renderTabs = () => {
    tabs.innerHTML = "";
    wb.sheets.forEach((sheet, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-tab";
      b.textContent = sheet.name;
      b.setAttribute("role", "tab");
      b.setAttribute("aria-selected", String(i === active));
      b.tabIndex = i === active ? 0 : -1; // roving tabindex for the tablist
      b.addEventListener("click", () => switchSheet(i));
      if (canManageSheets) {
        b.addEventListener("dblclick", (e) => { e.preventDefault(); beginRenameTab(b, i); });
        b.addEventListener("contextmenu", (e) => { e.preventDefault(); showTabMenu(i, b.offsetLeft); });
      }
      // Left/Right (Home/End) move between sheet tabs, activating and focusing each.
      b.addEventListener("keydown", (e) => {
        const n = wb.sheets.length;
        let to = -1;
        if (e.key === "ArrowRight") to = (i + 1) % n;
        else if (e.key === "ArrowLeft") to = (i - 1 + n) % n;
        else if (e.key === "Home") to = 0;
        else if (e.key === "End") to = n - 1;
        if (to < 0) return;
        e.preventDefault();
        switchSheet(to);
        (tabs.children[to] as HTMLElement | undefined)?.focus();
      });
      tabs.appendChild(b);
    });
    if (canManageSheets) {
      const addBtn = document.createElement("button");
      addBtn.type = "button";
      addBtn.className = "sheetedit-tab-add";
      addBtn.title = t("sheetAdd");
      addBtn.setAttribute("aria-label", t("sheetAdd"));
      addBtn.innerHTML = `<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" aria-hidden="true"><path d="M8 3.5v9M3.5 8h9"/></svg>`;
      addBtn.addEventListener("click", doAddSheet);
      tabs.appendChild(addBtn);
    }
  };

  // Floating style bar near the selection (approach-triggered, like richdoc).
  const selRectNow = (): DOMRect | null => {
    if (!sel || pickCb || dragActive) return null;
    const tl = tds.get(key(sel.r1, sel.c1));
    const br = tds.get(key(sel.r2, sel.c2));
    if (!tl || !br) return null;
    const a = tl.getBoundingClientRect();
    const b = br.getBoundingClientRect();
    return new DOMRect(a.left, a.top, b.right - a.left, b.bottom - a.top);
  };
  const floatBar = wb.kind === "xlsx" || wb.kind === "ods"
    ? setupFloatBar({ wrap, bounds: () => gridScroll.getBoundingClientRect(), selRect: selRectNow, curStyle, applyStyle })
    : { teardown: () => undefined };

  renderTabs();
  renderGrid();

  return {
    isDirty() {
      return dirty;
    },
    markClean() {
      dirty = false;
    },
    getCellValue(ref: string) {
      const p = parseA1Ref(ref);
      const sheet = wb.sheets[active];
      return p && sheet ? (getCell(sheet, p.row, p.col)?.value ?? "") : "";
    },
    setCellValue(ref: string, value: string) {
      const p = parseA1Ref(ref);
      if (p) commitValue(p.row, p.col, String(value));
    },
    getText() {
      // No recalc needed: the model is recalculated after every edit, and CSV
      // stores formula text, never computed results.
      return wb.kind === "csv" ? writeCsv(wb) : null;
    },
    async getBytes() {
      return dirty ? writeWorkbookAsync(wb) : original.slice();
    },
    destroy() {
      document.removeEventListener("copy", onDocCopy);
      document.removeEventListener("paste", onDocPaste);
      window.removeEventListener("pointerup", endDrag);
      window.removeEventListener("pointercancel", endDrag);
      closeLineMenu();
      borderPop?.remove();
      borderPop = null;
      furiPop?.remove();
      furiPop = null;
      toolbarHandle.teardown();
      floatBar.teardown();
      wrap.remove();
    },
  };
}

// Same as createSheetEditor, but a zip-based workbook (.xlsx/.ods) is inflated off the main
// thread first, so opening a large workbook does not freeze the UI on the unzip. CSV/TSV and
// CFB (.xls / encrypted) inputs are not zips: they skip the inflate and parse synchronously.
// The DOM-bound sheet/model parse still runs on the main thread once the map is ready.
export async function createSheetEditorAsync(
  container: HTMLElement,
  bytes: Uint8Array,
  options: SheetEditorOptions = {},
): Promise<SheetEditor> {
  let files: Record<string, Uint8Array> | undefined;
  const isZip = bytes.length > 3 && bytes[0] === 0x50 && bytes[1] === 0x4b; // "PK"
  if (isZip && options.formatHint !== "csv" && options.formatHint !== "tsv") {
    try {
      files = await unzipAsync(bytes);
    } catch {
      /* let createSheetEditor's own read handle the failure (error banner + original bytes) */
    }
  }
  return createSheetEditor(container, bytes, options, files);
}
