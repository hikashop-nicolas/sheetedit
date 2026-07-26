import { t } from "./i18n";
import { capabilitiesFor } from "./capabilities";
import { dateToSerial, hasTimeFmt, isDateFmt, isTimeOnlyFmt, serialToEditText } from "./dates";
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
import type { Cell, CellStyle, DataValidation, Phonetic, Sheet, StyleChange, Workbook, SheetControl } from "./model";
import { cellDisplay, colToLetters, ensureCell, getCell, key, parseA1Ref } from "./model";
import { setOdsAutoFilter, setOdsCellNumFmt, setOdsCellStyle, setOdsColWidth, setOdsMerge, setOdsRowHeight, setOdsSparkline } from "../adapters/ods";
import { makeFormulaEvaluator, recalc } from "./recalc";
import { applyRunStyle, cellRuns, isRunStyleChange, runsUniform, setRunStyle } from "./richtext";
import { csvToXlsx, writeCsv } from "../adapters/csv";
import { applyLineOp, syncXlsxMerges, type LineOp } from "./structure";
import { columnFilterValues, filterHiddenRows, sortedPositions, sortRange } from "./range-ops";
import { addSheet, renameSheet, deleteSheet, moveSheet, setSheetVisibility, sheetsEditable, visibleSheetCount } from "./sheet-ops";
import { SHEETEDIT_CSS } from "./ui/styles.generated";
import { SHEET_LOCK_DEFAULTS, canEditCell, canEditRange, hasPassword, isBlocked, isProtected, isStructureLocked, type SheetLock, type SheetProtection } from "./protection";
import { DEFAULT_MARGINS, DEFAULT_PAPER, PAPER_SIZES, toggleBreak, type PrintSetup } from "./print";
import { BUILTIN_THEMES, setWorkbookTheme } from "./theme";
import { buildPrintJob, type PrintJob } from "./print-render";
import { isSigned, subNames, vbaPartOf } from "./vba";
import { editModuleSource, findSheetHandler, findWorkbookHandler, hasEventHandlers, runWorkbookMacro, runnableSubs } from "./vba-macro";
import { setupControlLayer } from "./ui/control-layer";
import { absoluteRange, absoluteRef, createXlsxControl, defaultLink, deleteXlsxControl, placementFor, updateXlsxControlLinks } from "../adapters/xlsx/control-create";
import { hasActiveX } from "../adapters/xlsx/control-read";
import { setActiveXValue } from "../adapters/xlsx/activex-read";
import { formDialog, type FormField } from "./ui/form-dialog";
import { computeCondVisuals, type CfVisual } from "../adapters/xlsx/condformat";
import { resolveNumbers } from "./chart-data";
import { setupChartLayer } from "./ui/chart-overlay";
import { setupImageLayer } from "./ui/image-layer";
import { setupShapeLayer } from "./ui/shape-layer";
import { setupSlicerLayer } from "./ui/slicer-layer";
import { setupTimelineLayer } from "./ui/timeline-layer";
import { outlineGutterWidth, setupOutlineLayer } from "./ui/outline-layer";
import { setupPaneDividers } from "./ui/pane-divider";
import { clearOutline, groupLines, maxOutlineLevel, outlineLevel, setGroupCollapsed, showOutlineLevel, ungroupLines, type Axis } from "./outline";
import { setupPivotUi } from "./ui/pivot-ui";
import { setupDialogs } from "./ui/dialogs";
import { validateCell } from "./datavalidation";
import { setupPivotLayer } from "./ui/pivot-layer";
import { setupChartUi } from "./ui/chart-insert";
import { readWorkbook, setCellInput, writeWorkbookAsync } from "./workbook";
import { unzipAsync } from "./zip";
import { deleteXlsxShape, flagXlsxPivotRefresh, setXlsxAutoFilter, setXlsxCellNumFmt, setXlsxCellStyle, setXlsxColWidth, setXlsxMerge, setXlsxRowHeight, setXlsxRowHidden, setXlsxSparkline } from "../adapters/xlsx";
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
  s.textContent = SHEETEDIT_CSS;
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

  const caps = capabilitiesFor(wb.kind);

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
  // A transient message for an action protection refused, so a blocked edit is never silent.
  const notice = document.createElement("div");
  notice.className = "sheetedit-notice";
  notice.setAttribute("role", "status");
  notice.hidden = true;
  let noticeTimer: ReturnType<typeof setTimeout> | undefined;
  const showNotice = (text: string): void => {
    notice.textContent = text;
    notice.hidden = false;
    clearTimeout(noticeTimer);
    noticeTimer = setTimeout(() => { notice.hidden = true; }, 3200);
  };
  /**
   * Protection gate: true when the action may proceed. A refusal explains itself, because a
   * control that silently does nothing reads as a bug.
   */
  const allow = (ok: boolean, message: string): boolean => {
    if (!ok) showNotice(t(message));
    return ok;
  };
  const sheetAt = () => wb.sheets[active];
  /** Cell edits: allowed unless the sheet is protected and the target cell is locked. */
  const allowCellEdit = (r: number, c: number): boolean => allow(canEditCell(sheetAt(), r, c), "protectedCell");
  const allowRangeEdit = (g: { r1: number; c1: number; r2: number; c2: number }): boolean =>
    allow(canEditRange(sheetAt(), g), "protectedCell");
  /** Non-cell actions gated by one of the sheet's per-action flags. */
  const allowAction = (flag: SheetLock, message = "protectedSheet"): boolean => allow(!isBlocked(sheetAt(), flag), message);

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
      const ip = inputAt(key(activeCell.r, activeCell.c));
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
      const ip = inputAt(key(activeCell.r, activeCell.c));
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
  wrap.appendChild(notice); // absolute, over the grid: never shifts the layout
  container.appendChild(wrap);

  let active = 0;
  let condVisuals = new Map<string, CfVisual>(); // conditional-format visuals for the active sheet, per render
  let sparkAt = new Map<string, NonNullable<Sheet["sparklines"]>[number]>(); // host cell -> sparkline, per render
  // Print decoration for the active sheet, per render: the lines that start a page and the print
  // area, so a cell can be marked without re-scanning the setup for every one of them.
  let brkRows = new Set<number>();
  let brkCols = new Set<number>();
  let printAreas: { r1: number; c1: number; r2: number; c2: number }[] = [];
  /** Border classes for a cell that sits on a page break or a print-area edge. */
  const markPrintEdges = (td: HTMLElement, sheet: Sheet, r: number, c: number): void => {
    if (brkRows.has(r)) td.classList.add("pgbrk-top");
    if (brkCols.has(c)) td.classList.add("pgbrk-left");
    if (!printAreas.length) return;
    let inside = false;
    for (const a of printAreas) {
      if (r < a.r1 || r > a.r2 || c < a.c1 || c > a.c2) continue;
      inside = true;
      if (r === a.r1) td.classList.add("pa-top");
      if (r === a.r2) td.classList.add("pa-bottom");
      if (c === a.c1) td.classList.add("pa-left");
      if (c === a.c2) td.classList.add("pa-right");
    }
    // Shade what would not be printed, but only within the used range: tinting the empty grid to
    // the right of a small print area would read as a rendering fault.
    if (!inside && r <= sheet.maxRow && c <= sheet.maxCol) td.classList.add("pa-out");
  };
  /** The pane the user last pointed at; a split shows some cells twice and this picks the copy
      they are actually working in. */
  let lastPane: Pane | null = null;
  /** Panes in lookup order: the one last touched first. */
  const lookupPanes = (): Pane[] => (lastPane && panes.includes(lastPane) ? [lastPane, ...panes.filter((p) => p !== lastPane)] : panes);
  const inputAt = (k: string): HTMLInputElement | undefined => {
    for (const p of lookupPanes()) { const el = p.inputs.get(k); if (el) return el; }
    return undefined;
  };
  const tdAt = (k: string): HTMLElement | undefined => {
    for (const p of lookupPanes()) { const el = p.tds.get(k); if (el) return el; }
    return undefined;
  };
  /** Every rendering of a cell: a split can put the same cell in two panes at once. */
  const tdsAt = (k: string): HTMLElement[] => {
    const out: HTMLElement[] = [];
    for (const p of panes) { const el = p.tds.get(k); if (el) out.push(el); }
    return out;
  };
  // Extra rows/columns the user added beyond the sheet's used extent (per active sheet).
  let extraRows = 0;
  let extraCols = 0;
  // Assigned once the toolbar is built; reflects the active cell's style on the toggle buttons.
  let syncToolbar: () => void = () => undefined;
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
    const td = tdAt(key(sel.r2, sel.c2));
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
    // The fill only writes outside the source range, so gate exactly the cells it would touch.
    if (positions.some((p2) => !canEditCell(sheet, p2.r, p2.c))) {
      showNotice(t("protectedCell"));
      return;
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
        // A cell can be on screen in both panes of a split, so mark every rendering of it.
        if (range.horizontal) for (let r = r1; r <= r2; r++) for (const td of tdsAt(key(r, i))) td.classList.add("sheetedit-fillprev");
        else for (let c = c1; c <= c2; c++) for (const td of tdsAt(key(i, c))) td.classList.add("sheetedit-fillprev");
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
    for (const [k, td] of panes.flatMap((p) => [...p.tds])) {
      if (!sel) {
        td.classList.remove("sheetedit-sel");
        continue;
      }
      const [r, c] = k.split(":").map(Number);
      td.classList.toggle("sheetedit-sel", r! >= sel.r1 && r! <= sel.r2 && c! >= sel.c1 && c! <= sel.c2);
      td.classList.remove("sheetedit-fillsrc");
    }
    placeFillHandle();
    syncToolbar();
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

  // In-cell rich text: while a single cell is being edited with a sub-range of its text selected,
  // a run-applicable style change (bold/italic/underline/strike/colour/size/font) formats just that
  // range as a run rather than the whole cell. Returns the target, or null to fall through.
  const richRunTarget = (change: StyleChange): { r: number; c: number; start: number; end: number } | null => {
    if (!isRunStyleChange(change) || !activeCell) return null;
    const inp = inputAt(key(activeCell.r, activeCell.c));
    if (!inp || document.activeElement !== inp) return null; // must be actively editing this cell
    const start = inp.selectionStart ?? 0, end = inp.selectionEnd ?? 0;
    if (start >= end || (start === 0 && end === inp.value.length)) return null; // whole/none -> cell style
    return { r: activeCell.r, c: activeCell.c, start, end };
  };
  const applyRichRun = (rt: { r: number; c: number; start: number; end: number }, change: StyleChange) => {
    const sheet = wb.sheets[active]!;
    const inp = inputAt(key(rt.r, rt.c));
    // Commit any uncommitted typing first so the run offsets match the model's text.
    if (inp && inp.value !== rawOf(rt.r, rt.c)) commitValue(rt.r, rt.c, inp.value);
    const cell = getCell(sheet, rt.r, rt.c);
    if (!cell || cell.kind === "n" || cell.kind === "b" || cell.formula != null || cell.value === "") { syncToolbar(); return; }
    recordCells([{ r: rt.r, c: rt.c }], () => {
      const runs = applyRunStyle(cellRuns(cell), rt.start, rt.end, change);
      cell.richRuns = runsUniform(runs, cell) ? undefined : runs;
      cell.edited = true;
    });
    mark();
    renderGrid();
    const inp2 = inputAt(key(rt.r, rt.c)); // re-render replaced the input; restore the edit selection
    if (inp2) { inp2.focus(); try { inp2.setSelectionRange(rt.start, rt.end); } catch { /* ignore */ } }
    syncToolbar();
  };

  const applyStyle = (change: StyleChange) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    // Changing which cells are locked is protection management, not formatting: like Excel, it
    // requires the sheet to be unprotected first. Everything else is gated on formatCells.
    if (change.locked !== undefined) {
      if (!allow(!isProtected(wb.sheets[active]), "protectedUnprotectFirst")) return;
    } else if (!allowAction("formatCells")) return;
    const rt = richRunTarget(change);
    if (rt) { applyRichRun(rt, change); return; }
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const positions = selPositions(sheet);
    const foldRuns = isRunStyleChange(change); // a run-applicable change must also update rich cells
    let touchedRuns = false;
    recordCells(positions, () => {
      for (const pos of positions) {
        const cell = ensureCell(sheet, pos.r, pos.c);
        setCellStyle(sheet, cell, change);
        if (foldRuns && cell.richRuns?.length) { const runs = setRunStyle(cell.richRuns, 0, cell.value.length, change); cell.richRuns = runsUniform(runs, cell) ? undefined : runs; cell.edited = true; touchedRuns = true; }
      }
    });
    mark();
    // Wrap toggles or wrap cells change row heights, so full re-render; a rich-cell run update needs
    // its overlay rebuilt; everything else patches the rendered cells in place, keeping focus/scroll.
    if (touchedRuns || change.wrap !== undefined || positions.some((p) => getCell(sheet, p.r, p.c)?.cellStyle?.wrap)) renderGrid();
    else patchStyle(positions);
    syncToolbar(); // reflect the new state on the toggle buttons (bold, align, ...)
  };

  // Apply a number format preset to the selection (General when fmt is undefined).
  const applyNumFmt = (fmt: string | number | undefined, currency?: string) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    if (!allowAction("formatCells")) return;
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
  // Authoring controls the editor appends after the style cluster; collected so the toolbar can
  // fold the ones that don't fit into its "⋯" overflow menu.
  const trailingIcons: HTMLElement[] = [];

  // Power Query panel: only when the workbook carries a DataMashup payload. Refresh writes
  // the result through the model (undo-recorded when the target sheet is active) and the
  // table part's @ref is resized by applyQueryResult.
  if (wb.kind === "xlsx") {
    // Files attached for File.Contents, shared between the quick-refresh panel and the editor.
    const pqFiles: Record<string, Uint8Array> = {};
    // The quick-refresh panel + on-open auto-refresh only make sense once queries exist.
    if (workbookHasQueries(wb.files)) {
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
      trailingIcons.push(qBtn);
    }

    // Full Power Query editor (Applied Steps + live preview). Always available for xlsx so the
    // first query can be authored in a query-less workbook; the save path bootstraps the payload.
    const newGuid = (): string => { try { return `{${crypto.randomUUID().toUpperCase()}}`; } catch { return "{00000000-0000-0000-0000-000000000000}"; } };
    const editor = setupQueryEditor({
      wrap,
      wb,
      attachedFiles: pqFiles,
      save: (newM) => {
        void import("mlang/qdeff").then(({ writeWorkbookSectionM }) => {
          wb.files = writeWorkbookSectionM(wb.files, newM, newGuid());
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
        // No queries yet: open an empty section so the editor's Get Data can add the first one.
        editor.open(q ? q.mashup.sectionM : "section Section1;\r\n");
      });
    });
    trailingIcons.push(eBtn);
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
      // Excel raises Worksheet_Change for a replace-all too, with the whole affected block.
      fireSheetChange({
        r1: Math.min(...changes.map((ch) => ch.r)), c1: Math.min(...changes.map((ch) => ch.c)),
        r2: Math.max(...changes.map((ch) => ch.r)), c2: Math.max(...changes.map((ch) => ch.c)),
      });
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
    for (const [k, input] of panes.flatMap((p) => [...p.inputs])) {
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
    if (wb.kind === "xlsx" && wb.pivotCaches) flagXlsxPivotRefresh(wb, sheet.name, positions);
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

  // --- undo for sheet-level settings -------------------------------------------
  // A cell edit records its own fields, but protection, page setup, panes and outline grouping live
  // on the sheet, so a step for one of those carries a before/after snapshot of those fields
  // instead. Maps and sets are copied, or undo would hand back the very object it is meant to
  // restore.
  const SETTING_KEYS = ["protection", "printSetup", "freeze", "paneSplit", "rowOutline", "colOutline",
    "rowCollapsed", "colCollapsed", "hiddenRows", "hiddenCols", "summaryBelow", "summaryRight"] as const;
  type SettingKey = (typeof SETTING_KEYS)[number];
  type SettingsSnap = { [K in SettingKey]?: Sheet[K] };
  const cloneSetting = <T,>(v: T): T => (v instanceof Map ? (new Map(v) as T) : v instanceof Set ? (new Set(v) as T) : v);
  const snapSettings = (sheet: Sheet): SettingsSnap => {
    const out: SettingsSnap = {};
    for (const k of SETTING_KEYS) Object.assign(out, { [k]: cloneSetting(sheet[k]) });
    return out;
  };
  const restoreSettings = (sheet: Sheet, snap: SettingsSnap): void => {
    for (const k of SETTING_KEYS) Object.assign(sheet, { [k]: cloneSetting(snap[k]) });
    // The dirty flags are forced rather than restored: the model has just diverged from whatever was
    // last written, so every affected part has to be re-emitted. Restoring a stale "clean" flag
    // after a save would leave the undone change missing from the file.
    sheet.protectionDirty = true;
    sheet.printDirty = true;
    sheet.freezeDirty = true;
    sheet.outlineDirty = true;
    if (wb.kind === "ods") sheet.odsDirty = true;
  };
  /** Run a change to this sheet's settings as an undoable step. */
  const recordSettings = (mutate: () => void): void => {
    const si = active;
    const sheet = wb.sheets[si];
    if (!sheet) return;
    const before = snapSettings(sheet);
    mutate();
    const after = snapSettings(sheet);
    history.push({
      sheet: si,
      cells: [],
      undoExtra: () => { const s2 = wb.sheets[si]; if (s2) restoreSettings(s2, before); },
      redoExtra: () => { const s2 = wb.sheets[si]; if (s2) restoreSettings(s2, after); },
    });
  };
  /** The same, for a change that lives on the workbook rather than a sheet. */
  const recordWorkbook = (apply: () => void, revert: () => void): void => {
    apply();
    history.push({ sheet: active, cells: [], undoExtra: revert, redoExtra: apply });
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
    // Each of the four insert/delete actions has its own protection flag.
    const flag: SheetLock = op.axis === "row"
      ? (op.kind === "insert" ? "insertRows" : "deleteRows")
      : (op.kind === "insert" ? "insertColumns" : "deleteColumns");
    if (!allowAction(flag)) return;
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
  /** Set (or clear) the pane boundary and re-render; the write path picks it up from freezeDirty. */
  const setFreeze = (rows: number, cols: number, asSplit?: boolean): void => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    recordSettings(() => {
      sheet.freeze = rows > 0 || cols > 0 ? { rows, cols } : undefined;
      sheet.paneSplit = sheet.freeze && asSplit ? true : undefined;
      sheet.freezeDirty = true;
    });
    mark();
    renderGrid();
  };

  /** Set (or clear) this sheet's protection and re-render; the write path picks it up from the flag. */
  const setSheetProtection = (prot: SheetProtection | undefined): void => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    // Unprotecting drops the file's password hash with it, which the dialog warns about.
    recordSettings(() => {
      sheet.protection = prot;
      sheet.protectionDirty = true;
    });
    mark();
    renderGrid();
  };
  const setStructureProtection = (on: boolean): void => {
    const before = wb.protection;
    const after = on ? { ...wb.protection, structure: true } : { ...wb.protection, structure: undefined };
    const set = (v: typeof before): void => { wb.protection = v; wb.protectionDirty = true; renderTabs(); };
    recordWorkbook(() => set(after), () => set(before));
    mark();
  };
  // Protect-sheet dialog: the allowances, phrased as permissions the way Excel and Calc phrase
  // them, then stored as the formats' blocked-action flags.
  const PROT_ALLOWANCES: { key: string; label: string; flag: SheetLock; on: boolean }[] = [
    { key: "selectLocked", label: "protAllowSelectLocked", flag: "selectLockedCells", on: true },
    { key: "format", label: "protAllowFormat", flag: "formatCells", on: false },
    { key: "insertRows", label: "protAllowInsertRows", flag: "insertRows", on: false },
    { key: "insertCols", label: "protAllowInsertCols", flag: "insertColumns", on: false },
    { key: "deleteRows", label: "protAllowDeleteRows", flag: "deleteRows", on: false },
    { key: "deleteCols", label: "protAllowDeleteCols", flag: "deleteColumns", on: false },
    { key: "sort", label: "protAllowSort", flag: "sort", on: false },
    { key: "filter", label: "protAllowFilter", flag: "autoFilter", on: false },
  ];
  const openProtectDialog = (): void => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const cur = sheet.protection;
    const fields: FormField[] = [{ key: "note", label: t("protectNote"), type: "note" }];
    if (hasPassword(cur)) fields.push({ key: "pwnote", label: t("protectPasswordNote"), type: "note" });
    fields.push({ key: "allowHead", label: t("protectAllow"), type: "note" });
    for (const a of PROT_ALLOWANCES) {
      // Pre-tick from the current state when re-opening on an already-protected sheet.
      const blocked = cur?.locks?.[a.flag] ?? SHEET_LOCK_DEFAULTS[a.flag];
      fields.push({ key: a.key, label: t(a.label), type: "checkbox", value: cur ? !blocked : a.on });
    }
    formDialog(wrap, t("protectSheet"), fields, (vals) => {
      const locks: Partial<Record<SheetLock, boolean>> = {};
      for (const a of PROT_ALLOWANCES) locks[a.flag] = !vals[a.key]; // allowed here = not blocked in the file
      // Excel's default protect also blocks objects and scenarios; match it.
      setSheetProtection({ sheet: true, locks: { ...locks, objects: true, scenarios: true }, ...(cur?.password ? { password: cur.password } : {}) });
    });
  };

  /**
   * Run one macro against the workbook. The whole run is a single undo step, and a run that stops
   * part-way is rolled back by runWorkbookMacro before it returns, so a half-run macro can never
   * be saved.
   */
  const runMacro = (source: string, moduleName: string, procName: string): { ok: boolean; text: string } => {
    const res = runWorkbookMacro(wb, source, moduleName, procName, {
      activeSheet: active,
      selection: sel ? { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 } : undefined,
      activeCell: activeCell ?? undefined,
      fileName: options.fileName,
      // Worksheet.PrintOut opens the browser's print dialog, which is the only printing a page has.
      print: (sheetIndex) => doPrint({ scope: "sheet" }, sheetIndex),
    });
    if (!res.ok) {
      const where = res.error?.line != null ? ` (${res.error.module}, line ${res.error.line})` : ` (${res.error?.module})`;
      return { ok: false, text: `${t("vbaRunFailed")} ${res.error?.message ?? ""}${where}` };
    }
    if (res.undo && res.redo) {
      history.push({ sheet: active, cells: [], undoExtra: res.undo, redoExtra: res.redo });
      recalc(wb);
      mark(); // only when the macro actually changed something: a MsgBox does not dirty a workbook
    }
    if (res.activeSheet != null && res.activeSheet !== active && wb.sheets[res.activeSheet]) switchSheet(res.activeSheet);
    renderTabs();
    renderGrid();
    const out = res.messages.length ? `\n${t("vbaRunOutput")}: ${res.messages.join(" | ")}` : "";
    return { ok: true, text: `${t("vbaRunOk")}${out}` };
  };

  /** Show the workbook's macro source, and run a procedure from it. */
  const openMacroViewer = (): void => {
    const project = wb.vba;
    const bin = vbaPartOf(wb.files);
    const signedProject = !!bin && isSigned(bin);
    const modal = document.createElement("div");
    modal.className = "sheetedit-modal";
    const card = document.createElement("div");
    card.className = "sheetedit-card is-wide";
    const h = document.createElement("h3");
    h.textContent = t("vbaTitle");
    card.appendChild(h);
    const note = document.createElement("p");
    note.className = "sheetedit-note";
    note.textContent = project?.modules.length ? t("vbaNote") : t("vbaNone");
    card.appendChild(note);

    // The consent gate. Only shown when this workbook actually has handlers to run, so a workbook
    // with ordinary macros is not asked a question that means nothing for it.
    if (hasEventHandlers(wb)) {
      const row = document.createElement("label");
      row.className = "sheetedit-field is-inline sheetedit-vba-events";
      const box = document.createElement("input");
      box.type = "checkbox";
      box.checked = eventsEnabled;
      box.addEventListener("change", () => {
        eventsEnabled = box.checked;
        // Turning it on is the moment the user consents, so Workbook_Open runs then rather than
        // silently at open time before there was anyone to ask.
        if (eventsEnabled) fireEvent(findWorkbookHandler(wb, "Workbook_Open"));
      });
      const span = document.createElement("span");
      span.textContent = t("vbaEvents");
      row.append(box, span);
      const why = document.createElement("p");
      why.className = "sheetedit-note";
      why.textContent = t("vbaEventsNote");
      card.append(row, why);
    }
    if (signedProject) {
      const note = document.createElement("p");
      note.className = "sheetedit-note";
      note.textContent = t("vbaSigned");
      card.appendChild(note);
    }
    // An ActiveX control draws nothing on the grid, so without this the user sees a gap where a
    // control should be and no reason for it.
    if (hasActiveX(wb.files)) {
      const note = document.createElement("p");
      note.className = "sheetedit-note";
      note.textContent = t("vbaActiveX");
      card.appendChild(note);
    }

    if (project?.modules.length) {
      const list = document.createElement("div");
      list.className = "sheetedit-vba-list";
      // Editable: the source IS the macro now that sheetedit can write it back. Saving parses it
      // first, so syntactic nonsense cannot reach the file.
      const pre = document.createElement("textarea");
      pre.className = "sheetedit-vba-src";
      pre.spellcheck = false;
      const runs = document.createElement("div");
      runs.className = "sheetedit-vba-runs";
      const status = document.createElement("p");
      status.className = "sheetedit-note sheetedit-vba-status";
      let current = 0;
      const show = (i: number): void => {
        current = i;
        const mod = project.modules[i]!;
        pre.value = mod.source;
        for (const [j, b] of [...list.children].entries()) b.classList.toggle("is-current", j === i);
        // Naming the procedures is what makes a long module readable at a glance.
        const subs = subNames(mod.source);
        note.textContent = subs.length ? `${t("vbaNote")}  ${t("vbaSubs")}: ${subs.join(", ")}` : t("vbaNote");
        // Only the procedures that can run on their own get a button: a Sub needing arguments has
        // nothing to supply them, and a Function is not something a user "runs".
        status.textContent = "";
        runs.textContent = "";
        const save = document.createElement("button");
        save.type = "button";
        save.className = "sheetedit-dlg-btn";
        save.textContent = t("vbaSave");
        // A signed project cannot be rewritten without invalidating its signature, so say that
        // before the user has typed an edit rather than after they try to save it.
        if (signedProject) { save.disabled = true; save.title = t("vbaSigned"); }
        save.addEventListener("click", () => {
          const target = project.modules[current]!;
          const res = editModuleSource(wb, target.name, pre.value);
          status.classList.toggle("is-error", !res.ok);
          if (!res.ok) { status.textContent = `${t("vbaSaveFailed")} ${res.error ?? ""}`; return; }
          if (res.undo && res.redo) history.push({ sheet: active, cells: [], undoExtra: res.undo, redoExtra: res.redo });
          mark();
          show(current); // the Run list follows whatever the module now declares
          status.textContent = t("vbaSaved"); // after show(), which clears the status line
        });
        runs.appendChild(save);
        const runnable = runnableSubs(mod.source);
        if (!runnable.length) {
          const none = document.createElement("span");
          none.className = "sheetedit-note";
          none.textContent = t("vbaRunNone");
          runs.appendChild(none);
          return;
        }

        for (const name of runnable) {
          const b = document.createElement("button");
          b.type = "button";
          b.className = "sheetedit-dlg-btn";
          b.textContent = `${t("vbaRun")}: ${name}`;
          b.addEventListener("click", () => {
            const res = runMacro(mod.source, mod.name, name);
            status.textContent = res.text;
            status.classList.toggle("is-error", !res.ok);
          });
          runs.appendChild(b);
        }
      };
      project.modules.forEach((mod, i) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-vba-mod";
        b.textContent = mod.name;
        b.addEventListener("click", () => show(i));
        list.appendChild(b);
      });
      card.append(list, pre, runs, status);
      show(0);
    }
    const actions = document.createElement("div");
    actions.className = "sheetedit-actions";
    const close = document.createElement("button");
    close.className = "sheetedit-dlg-btn";
    close.textContent = t("chartCancel");
    close.addEventListener("click", () => modal.remove());
    actions.appendChild(close);
    card.appendChild(actions);
    modal.appendChild(card);
    wrap.appendChild(modal);
    modal.addEventListener("mousedown", (e) => { if (e.target === modal) modal.remove(); });
  };

  // --- workbook theme ----------------------------------------------------------
  /**
   * Switch the workbook's palette. Cells reference the theme by index, so this recolours every
   * cell that used it and rewrites only theme1.xml; cells given an explicit colour are untouched.
   */
  const openThemeDialog = (): void => {
    const modal = document.createElement("div");
    modal.className = "sheetedit-modal";
    const card = document.createElement("div");
    card.className = "sheetedit-card";
    const h = document.createElement("h3");
    h.textContent = t("themeTitle");
    card.appendChild(h);
    const note = document.createElement("p");
    note.className = "sheetedit-note";
    note.textContent = t("themeNote");
    card.appendChild(note);
    const close = () => modal.remove();

    const list = document.createElement("div");
    list.className = "sheetedit-theme-list";
    const current = wb.theme;
    // The file's own theme leads the list when it is not one of the built-ins, so switching away
    // and back is possible without losing it.
    const known = BUILTIN_THEMES.some((th) => th.name === current?.name);
    const options = current && !known ? [current, ...BUILTIN_THEMES] : BUILTIN_THEMES;
    for (const theme of options) {
      const opt = document.createElement("button");
      opt.type = "button";
      opt.className = "sheetedit-theme-opt";
      opt.dataset.theme = theme.name;
      if (theme.name === current?.name) opt.classList.add("is-current");
      const name = document.createElement("span");
      name.className = "sheetedit-theme-name";
      name.textContent = theme.name;
      opt.appendChild(name);
      if (theme.name === current?.name) {
        const tag = document.createElement("span");
        tag.className = "sheetedit-theme-tag";
        tag.textContent = t("themeCurrent");
        opt.appendChild(tag);
      }
      // The six accents are what a palette is recognised by.
      const chips = document.createElement("span");
      chips.className = "sheetedit-theme-swatches";
      for (const slot of ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const) {
        const chip = document.createElement("span");
        chip.className = "sheetedit-theme-chip";
        chip.style.background = theme.colors[slot]; // the palette itself: data, not styling
        chips.appendChild(chip);
      }
      opt.appendChild(chips);
      opt.addEventListener("click", () => {
        close();
        // Switching back to the previous palette is an exact inverse: the cells kept their theme
        // references, so re-resolving against the old theme restores every colour it had changed.
        const previous = wb.theme;
        const next = { ...theme, colors: { ...theme.colors } };
        recordWorkbook(
          () => setWorkbookTheme(wb, next),
          () => { if (previous) setWorkbookTheme(wb, previous); },
        );
        mark();
        renderGrid();
      });
      list.appendChild(opt);
    }
    card.appendChild(list);
    const actions = document.createElement("div");
    actions.className = "sheetedit-actions";
    const cancel = document.createElement("button");
    cancel.className = "sheetedit-dlg-btn";
    cancel.textContent = t("chartCancel");
    cancel.addEventListener("click", close);
    actions.appendChild(cancel);
    card.appendChild(actions);
    modal.appendChild(card);
    wrap.appendChild(modal);
    modal.addEventListener("mousedown", (e) => { if (e.target === modal) close(); });
  };

  // --- print setup -------------------------------------------------------------
  /** Change the active sheet's page setup and flag it for the writer. */
  const updatePrint = (change: (p: PrintSetup) => void): void => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    recordSettings(() => {
      const p: PrintSetup = { ...(sheet.printSetup ?? {}) };
      change(p);
      sheet.printSetup = p;
      sheet.printDirty = true;
    });
    mark();
    renderGrid();
  };
  const MARGIN_PRESETS: Record<string, PrintSetup["margins"]> = {
    narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    normal: { ...DEFAULT_MARGINS },
    wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  };
  /** Which preset the current margins match, for pre-selecting the dropdown. */
  const marginPresetName = (m: PrintSetup["margins"]): string => {
    for (const [name, preset] of Object.entries(MARGIN_PRESETS))
      if (m && preset && (["left", "right", "top", "bottom", "header", "footer"] as const).every((k) => Math.abs(m[k] - preset[k]) < 0.001)) return name;
    return "normal";
  };
  /**
   * Lay the print area out as pages and hand them to the browser. The pages live in the document
   * (off-screen) rather than in an iframe, so they inherit the fonts already loaded, and a print
   * stylesheet hides everything else.
   */
  let printRoot: HTMLElement | null = null;
  const doPrint = (job: PrintJob, sheetIndex = active): void => {
    printRoot?.remove();
    const result = buildPrintJob(wb, sheetIndex, options.fileName ?? "", job);
    if (!result) {
      showNotice(t("printNothing"));
      return;
    }
    // One @page rule covers the whole job, so a workbook whose sheets disagree on paper cannot be
    // printed faithfully in one go. Say so rather than silently printing them all at one size.
    if (result.mixedPaper) showNotice(t("printMixedPaper"));
    printRoot = result.root;
    document.body.appendChild(printRoot);
    const cleanup = (): void => { printRoot?.remove(); printRoot = null; };
    // afterprint is the reliable signal in every current browser; the timeout is the belt and
    // braces for one that fires nothing, so the pages can never be left behind in the document.
    window.addEventListener("afterprint", cleanup, { once: true });
    setTimeout(() => { if (printRoot) cleanup(); }, 60000);
    // Force layout before printing rather than waiting for a frame: requestAnimationFrame is
    // throttled in a background tab, which left the pages built and the dialog never opened.
    void printRoot.offsetHeight;
    window.print();
  };

  /** Ask what the job covers, then print it. A single-sheet workbook with no selection just prints. */
  const openPrintScopeDialog = (): void => {
    const cur = sel;
    const multiSheet = wb.sheets.length > 1;
    const hasSelection = !!cur && (cur.r1 !== cur.r2 || cur.c1 !== cur.c2);
    if (!multiSheet && !hasSelection) {
      doPrint({ scope: "sheet" });
      return;
    }
    const options2 = [{ value: "sheet", label: t("printScopeSheet") }];
    if (multiSheet) options2.push({ value: "all", label: t("printScopeAll") });
    if (hasSelection) options2.push({ value: "selection", label: t("printScopeSelection") });
    formDialog(wrap, t("printScope"), [{ key: "scope", label: t("printScope"), type: "select", value: "sheet", options: options2 }], (vals) => {
      const scope = String(vals.scope);
      doPrint(scope === "selection" && cur ? { scope: "sheet", selection: { ...cur } } : { scope: scope === "all" ? "all" : "sheet" });
    });
  };

  const openPrintDialog = (): void => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const p = sheet.printSetup ?? {};
    const scaling = p.fitToPage && (p.fitToWidth || p.fitToHeight) ? "fit" : p.scale != null && p.scale !== 100 ? "percent" : "none";
    const fields: FormField[] = [
      { key: "note", label: t("printNoPrinting"), type: "note" },
      { key: "orientation", label: t("printOrientation"), type: "select", value: p.orientation ?? "portrait", options: [
        { value: "portrait", label: t("printPortrait") },
        { value: "landscape", label: t("printLandscape") },
      ] },
      { key: "paper", label: t("printPaper"), type: "select", value: String(p.paperSize ?? DEFAULT_PAPER), options: Object.entries(PAPER_SIZES).map(([id, s]) => ({ value: id, label: s.name })) },
      { key: "scaling", label: t("printScaling"), type: "select", value: scaling, options: [
        { value: "none", label: t("printScaleNormal") },
        { value: "percent", label: t("printScalePercent") },
        { value: "fit", label: t("printScaleFit") },
      ] },
      { key: "scale", label: t("printScaleValue"), type: "text", value: String(p.scale ?? 100), showFor: { key: "scaling", values: ["percent"] } },
      { key: "margins", label: t("printMargins"), type: "select", value: marginPresetName(p.margins), options: [
        { value: "narrow", label: t("printMarginNarrow") },
        { value: "normal", label: t("printMarginNormal") },
        { value: "wide", label: t("printMarginWide") },
      ] },
      { key: "gridLines", label: t("printGridLines"), type: "checkbox", value: !!p.gridLines },
      { key: "headings", label: t("printHeadings"), type: "checkbox", value: !!p.headings },
      { key: "centerH", label: t("printCenterH"), type: "checkbox", value: !!p.horizontalCentered },
      { key: "centerV", label: t("printCenterV"), type: "checkbox", value: !!p.verticalCentered },
      { key: "hint", label: t("printFieldHint"), type: "note" },
      { key: "hl", label: t("printHeaderLeft"), type: "text", value: p.header?.left ?? "" },
      { key: "hc", label: t("printHeaderCenter"), type: "text", value: p.header?.center ?? "" },
      { key: "hr", label: t("printHeaderRight"), type: "text", value: p.header?.right ?? "" },
      { key: "fc", label: t("printFooterCenter"), type: "text", value: p.footer?.center ?? "" },
    ];
    formDialog(wrap, t("printSetup"), fields, (vals) => {
      updatePrint((next) => {
        next.orientation = String(vals.orientation) as PrintSetup["orientation"];
        next.paperSize = Number(vals.paper) || DEFAULT_PAPER;
        const mode = String(vals.scaling);
        // The three scaling modes are mutually exclusive, so each one clears the others.
        if (mode === "fit") { next.fitToPage = true; next.fitToWidth = 1; next.fitToHeight = 0; next.scale = 100; }
        else if (mode === "percent") { next.fitToPage = false; next.fitToWidth = undefined; next.fitToHeight = undefined; next.scale = Math.min(400, Math.max(10, Number(vals.scale) || 100)); }
        else { next.fitToPage = false; next.fitToWidth = undefined; next.fitToHeight = undefined; next.scale = 100; }
        next.margins = { ...MARGIN_PRESETS[String(vals.margins)]! };
        next.gridLines = !!vals.gridLines;
        next.headings = !!vals.headings;
        next.horizontalCentered = !!vals.centerH;
        next.verticalCentered = !!vals.centerV;
        const hl = String(vals.hl), hc = String(vals.hc), hr = String(vals.hr), fc = String(vals.fc);
        next.header = hl || hc || hr ? { ...(hl ? { left: hl } : {}), ...(hc ? { center: hc } : {}), ...(hr ? { right: hr } : {}) } : undefined;
        next.footer = fc ? { center: fc } : undefined;
      });
    });
  };

  const openLineMenu = (e: MouseEvent, axisOf: "row" | "col", line: number) => {
    const axis = axisOf;
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
    // Outline grouping: group/ungroup the clicked span, and clear the axis when it has groups.
    if (caps.outline) {
      const sheet = wb.sheets[active];
      const axis: Axis = axisOf;
      const limit = axis === "row" ? totalRows : totalCols;
      const from = base, to = base + n - 1;
      const act = (text: string, run: () => void) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-pop-item";
        b.textContent = text;
        b.addEventListener("click", () => {
          closeLineMenu();
          if (!sheet) return;
          recordSettings(() => { run(); if (wb.kind === "ods") sheet.odsDirty = true; });
          mark();
          renderGrid();
        });
        pop.appendChild(b);
      };
      const sep = document.createElement("div");
      sep.className = "sheetedit-pop-sep";
      pop.appendChild(sep);
      act(t("outlineGroup"), () => groupLines(sheet!, axis, from, to));
      if (sheet && outlineLevel(sheet, axis, base) > 0) {
        act(t("outlineUngroup"), () => ungroupLines(sheet, axis, from, to));
        const level = outlineLevel(sheet, axis, base);
        act(t("outlineCollapse"), () => setGroupCollapsed(sheet, axis, base, level, true, limit));
        act(t("outlineExpand"), () => setGroupCollapsed(sheet, axis, base, level, false, limit));
      }
      if (sheet && maxOutlineLevel(sheet, axis) > 0) act(t("outlineClear"), () => clearOutline(sheet, axis));
    }
    // Freeze panes, expressed on the axis whose header was clicked.
    if (caps.freezePanes) {
      const sheet = wb.sheets[active];
      const fr = sheet?.freeze?.rows ?? 0, fc = sheet?.freeze?.cols ?? 0;
      const sep2 = document.createElement("div");
      sep2.className = "sheetedit-pop-sep";
      pop.appendChild(sep2);
      const item = (text: string, run: () => void) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-pop-item";
        b.textContent = text;
        b.addEventListener("click", () => { closeLineMenu(); run(); });
        pop.appendChild(b);
      };
      // "Freeze rows above" pins everything before the clicked line, which is what you want when
      // you right-click the first row of the body.
      if (axis === "row") item(t("freezeRowsAbove"), () => setFreeze(Math.max(0, base - 1), fc));
      else item(t("freezeColsLeft"), () => setFreeze(fr, Math.max(0, base - 1)));
      if (fr > 0 || fc > 0) item(t("unfreeze"), () => setFreeze(0, 0));
    }
    // Print: a page break at the clicked line, and repeating this run on every page.
    if (caps.printSetup) {
      const sheet = wb.sheets[active];
      const p = sheet?.printSetup;
      const sep3 = document.createElement("div");
      sep3.className = "sheetedit-pop-sep";
      pop.appendChild(sep3);
      const item = (text: string, run: () => void) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-pop-item";
        b.textContent = text;
        b.addEventListener("click", () => { closeLineMenu(); run(); });
        pop.appendChild(b);
      };
      const breaks = (axisOf === "row" ? p?.rowBreaks : p?.colBreaks) ?? [];
      // A break at line 1 would start a page before the sheet begins, so it is never offered.
      if (base > 1) {
        item(breaks.includes(base) ? t("printBreakRemove") : t("printBreakAdd"), () =>
          updatePrint((next) => {
            if (axisOf === "row") next.rowBreaks = toggleBreak(next.rowBreaks, base);
            else next.colBreaks = toggleBreak(next.colBreaks, base);
          }));
      }
      const titles = axisOf === "row" ? p?.titleRows : p?.titleCols;
      const span = { from: base, to: base + n - 1 };
      const sameSpan = titles && titles.from === span.from && titles.to === span.to;
      item(sameSpan ? t("printTitlesClear") : t(axisOf === "row" ? "printTitleRows" : "printTitleCols"), () =>
        updatePrint((next) => {
          const value = sameSpan ? undefined : span;
          if (axisOf === "row") next.titleRows = value;
          else next.titleCols = value;
        }));
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
    if (!allowCellEdit(r, c)) return;
    recordCells([{ r, c }], () => setCellInput(sheet, r, c, raw));
    recalc(wb);
    mark();
    // A wrap cell's new text may change its row height; a validated cell may gain/lose its
    // invalid flag; a conditional-format edit can recolour the whole range; any of these needs a
    // re-render. Otherwise just refresh displays.
    if (getCell(sheet, r, c)?.cellStyle?.wrap || dvForCell(sheet, r, c) || sheet.condFormats?.length) renderGrid();
    else refreshDisplays(sheet);
    if (sheet.charts?.length) chartLayer.update(); // live-update any chart reading this cell
    fireSheetChange({ r1: r, c1: c, r2: r, c2: c });
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
    const cellInput = activeCell ? inputAt(key(activeCell.r, activeCell.c)) : undefined;
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
    const editor = editingBar ? fxbar.input : (inputAt(key(target.r, target.c)) ?? fxbar.input);
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
      const ip = inputAt(key(target.r, target.c));
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
    if (!allowRangeEdit(range)) return;
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
    const height = grid2.length, width = Math.max(...grid2.map((g) => g.length));
    if (!allowRangeEdit({ r1: r0, c1: c0, r2: r0 + height - 1, c2: c0 + width - 1 })) return;
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
  let heightRows: number[] = [];
  let heightPrefix: number[] = [];
  let widthCols: number[] = [];
  let widthPrefix: number[] = [];
  /**
   * A scrolling viewport over the sheet. There is always the main pane; a row SPLIT adds a second
   * one below it that scrolls rows on its own while sharing the horizontal scroll. Each pane keeps
   * its own rendered window and its own cell elements, because a split can show the same row twice.
   */
  interface Pane {
    scrollEl: HTMLElement;
    tableEl: HTMLTableElement | null;
    winR1: number; winR2: number; winC1: number; winC2: number;
    inputs: Map<string, HTMLInputElement>;
    tds: Map<string, HTMLElement>;
    /** Only the top band draws the column header; the band below continues under it. */
    header: boolean;
    /** Only the left band draws the row-number column; the band right of it continues past it. */
    rowHeader: boolean;
    /** Which quadrant this is: row band 0/1 (above/below a row split), column band 0/1. */
    band: { r: 0 | 1; c: 0 | 1 };
  }
  const newPane = (scrollEl: HTMLElement, r: 0 | 1, c: 0 | 1): Pane =>
    ({ scrollEl, tableEl: null, winR1: 1, winR2: 0, winC1: 1, winC2: 0, inputs: new Map(), tds: new Map(), header: r === 0, rowHeader: c === 0, band: { r, c } });
  const mainPane: Pane = newPane(gridScroll, 0, 0);
  let panes: Pane[] = [mainPane];
  /** The pane currently being built, so the cell builders write into the right maps. */
  let cur: Pane = mainPane;

  // The extra viewports. A row split adds a band below, a column split adds a band to the right,
  // and both together give Excel's four panes. Each pane scrolls the axis its boundary cuts and
  // shares the other: panes in the same row band scroll vertically together, panes in the same
  // column band scroll horizontally together.
  const mkPane = (cls: string, r: 0 | 1, c: 0 | 1): Pane => {
    const el = document.createElement("div");
    el.className = `sheetedit-grid ${cls}`;
    el.style.display = "none";
    return newPane(el, r, c);
  };
  const splitPane = mkPane("sheetedit-grid-split", 1, 0);
  const rightPane = mkPane("sheetedit-grid-right", 0, 1);
  const rightSplitPane = mkPane("sheetedit-grid-right sheetedit-grid-split", 1, 1);
  const splitScroll = splitPane.scrollEl;
  // Two column bands side by side, each stacking its row bands.
  const gridArea = document.createElement("div");
  gridArea.className = "sheetedit-gridarea";
  const colBand0 = document.createElement("div");
  colBand0.className = "sheetedit-colband";
  const colBand1 = document.createElement("div");
  colBand1.className = "sheetedit-colband sheetedit-colband-right";
  // Take gridScroll's place in the wrap FIRST, then adopt it: assembling the other way round would
  // put the area inside a band that is inside the area.
  gridScroll.replaceWith(gridArea);
  colBand0.append(gridScroll, splitScroll);
  colBand1.append(rightPane.scrollEl, rightSplitPane.scrollEl);
  gridArea.append(colBand0, colBand1);

  let syncing = false;
  /** Mirror the axis a pane does not own onto the panes that share it. */
  const shareScroll = (from: Pane): void => {
    if (syncing) return;
    syncing = true;
    for (const p of panes) {
      if (p === from) continue;
      if (p.band.r === from.band.r) p.scrollEl.scrollTop = from.scrollEl.scrollTop;
      if (p.band.c === from.band.c) p.scrollEl.scrollLeft = from.scrollEl.scrollLeft;
    }
    syncing = false;
  };

  /** A freshly opened split has to be snapped to the rendered boundary once its rows exist. */
  let snapSplit = false;
  /** Lay the panes out for the sheet's boundary: one viewport, or two split at the boundary. */
  const layoutPanes = (sheet: Sheet): void => {
    const rows = sheet.paneSplit ? sheet.freeze?.rows ?? 0 : 0;
    const cols = sheet.paneSplit ? sheet.freeze?.cols ?? 0 : 0;
    const rowOn = rows > 0, colOn = cols > 0;
    if ((rowOn || colOn) && panes.length < 2) snapSplit = true;
    const wanted = [mainPane];
    if (colOn) wanted.push(rightPane);
    if (rowOn) wanted.push(splitPane);
    if (rowOn && colOn) wanted.push(rightSplitPane);
    panes = wanted;
    for (const p of [splitPane, rightPane, rightSplitPane]) {
      const on = wanted.includes(p);
      p.scrollEl.style.display = on ? "block" : "none";
      if (!on) { p.tableEl = null; p.scrollEl.innerHTML = ""; }
    }
    colBand1.style.display = colOn ? "flex" : "none";
    // The leading bands are exactly as big as their boundary. This runs before the lines exist, so
    // start from the layout model and correct it against the rendered ones in fitSplitSizes.
    const hh = headerH();
    const bandH = rowOn ? `0 0 ${Math.max(ROW_H * 2, hh + yOfRow(rows + 1))}px` : "";
    gridScroll.style.flex = bandH;
    rightPane.scrollEl.style.flex = bandH;
    colBand0.style.flex = colOn ? `0 0 ${Math.max(60, rnW() + xOfCol(cols + 1))}px` : "1";
  };

  /**
   * Trim the top viewport to the rendered bottom of its last row, so a split does not leave a
   * sliver of the next row showing. The declared row heights and the rendered ones differ, which is
   * exactly the gap this closes.
   */
  const fitSplitSizes = (): void => {
    const sheet = wb.sheets[active];
    const rows = sheet?.paneSplit ? sheet.freeze?.rows ?? 0 : 0;
    const cols = sheet?.paneSplit ? sheet.freeze?.cols ?? 0 : 0;
    if (panes.length < 2) return;
    const gr = gridScroll.getBoundingClientRect();
    const last = rows > 0 ? (gridScroll.querySelector(`th.rownum[data-r="${rows}"]`) as HTMLElement | null) : null;
    if (last) {
      const h = last.getBoundingClientRect().bottom - gr.top;
      if (h > ROW_H) { const v = `0 0 ${Math.round(h)}px`; gridScroll.style.flex = v; rightPane.scrollEl.style.flex = v; }
    }
    const lastC = cols > 0 ? (gridScroll.querySelector(`th.colhead[data-c="${cols}"]`) as HTMLElement | null) : null;
    if (lastC) {
      const w = lastC.getBoundingClientRect().right - gr.left;
      if (w > 40) colBand0.style.flex = `0 0 ${Math.round(w)}px`;
    }
    // Line the lower viewport up with the first row past the boundary, so it does not open showing
    // the tail of the row above. Only on creation: after that the pane scrolls where the user puts it.
    if (!snapSplit) return;
    // Only give up the snap once the reference lines were actually there to measure: an early
    // render can happen before the panes have any size, and a half-applied snap looks like a bug.
    let done = true;
    const firstBelow = rows > 0 ? (splitScroll.querySelector(`th.rownum[data-r="${rows + 1}"]`) as HTMLElement | null) : null;
    if (rows > 0 && !firstBelow) done = false;
    if (firstBelow) {
      const d = firstBelow.getBoundingClientRect().top - splitScroll.getBoundingClientRect().top;
      if (Math.abs(d) > 0.5) splitScroll.scrollTop += d;
    }
    const rel = rightPane.scrollEl;
    const firstRight = cols > 0 ? (rel.querySelector(`th.colhead[data-c="${cols + 1}"]`) as HTMLElement | null) : null;
    if (cols > 0 && !firstRight) done = false;
    if (firstRight) {
      const d = firstRight.getBoundingClientRect().left - rel.getBoundingClientRect().left;
      if (Math.abs(d) > 0.5) { rel.scrollLeft += d; rightSplitPane.scrollEl.scrollLeft = rel.scrollLeft; }
      else if (rel.clientWidth === 0) done = false;
    }
    snapSplit = !done;
  };

  let coveredSet = new Set<string>();
  let spanAtMap = new Map<string, { rs: number; cs: number }>();

  // Wrap: computed extra height (px) so a wrapped cell's text fits, measured against the
  // column width. Keyed by row for the active sheet; recomputed on render / resize / edit.
  const wrapH = new Map<number, number>();
  let measureEl: HTMLElement | null = null;
  const measureWrap = (text: string, widthPx: number, cs: CellStyle | undefined): number => {
    if (!measureEl) {
      measureEl = document.createElement("div");
      measureEl.className = "sheetedit-measure";
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
    sheet.hiddenRows?.has(r) || sheet.filterHidden?.has(r) ? 0 : Math.max(sheet.rowHeights?.get(r) ?? ROW_H, wrapH.get(r) ?? 0);
  const effColW = (sheet: Sheet, c: number): number => (sheet.hiddenCols?.has(c) ? 0 : (sheet.colWidths?.get(c) ?? COL_W));
  const rowShown = (sheet: Sheet, r: number): boolean => !sheet.hiddenRows?.has(r) && !sheet.filterHidden?.has(r);

  const rebuildSizeIndexes = (sheet: Sheet) => {
    heightRows = [...new Set([...(sheet.rowHeights?.keys() ?? []), ...(sheet.hiddenRows ?? []), ...(sheet.filterHidden ?? []), ...wrapH.keys()])].sort((a, b) => a - b);
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
  /** Rendered height of the column-header row. The table has no <thead> (the header row is appended
      straight to it), and the declared row height is not the rendered one, so measure the corner cell. */
  const headerH = (): number => {
    const corner = gridScroll.querySelector("th.corner") as HTMLElement | null;
    return corner ? corner.getBoundingClientRect().height : ROW_H;
  };
  /** Row-number column width: grows with the digit count of the last row, plus the outline gutter. */
  const rnW = (): number => Math.max(44, 18 + String(totalRows).length * 8) + outlineGutterWidth(wb.sheets[active]);

  // Chart overlay + create/edit UI: a floating Chart.js layer glued to the cells (xlsx/ods only),
  // Chart.js lazy-loaded on the first chart.
  const chartsOn = caps.charts;
  const getSelRect = (): { r1: number; c1: number; r2: number; c2: number } =>
    sel ?? (activeCell ? { r1: activeCell.r, c1: activeCell.c, r2: activeCell.r, c2: activeCell.c } : { r1: 1, c1: 1, r2: 1, c2: 1 });
  const chartLayer = chartsOn
    ? setupChartLayer({
        wrap,
        panes: () => panes.map((p) => ({ el: p.scrollEl, header: p.header, rowHeader: p.rowHeader })),
        getSheet: () => wb.sheets[active],
        getWorkbook: () => wb,
        geom: () => ({ xOfCol, yOfRow, colAt: (px) => lineAt(px, totalCols, xOfCol), rowAt: (px) => lineAt(px, totalRows, yOfRow), rnW: rnW(), headerH: headerH() }),
        onSelect: (c) => { if (c) chartUi.showEdit(c); else chartUi.hideEdit(); },
        onEdit: () => { mark(); },
      })
    : { refresh: () => undefined, update: () => undefined, select: () => undefined, boxRect: () => null, teardown: () => undefined };
  const imageLayer = setupImageLayer({
    wrap,
    panes: () => panes.map((p) => ({ el: p.scrollEl, header: p.header, rowHeader: p.rowHeader })),
    getSheet: () => wb.sheets[active],
    geom: () => ({ xOfCol, yOfRow, colAt: (px) => lineAt(px, totalCols, xOfCol), rowAt: (px) => lineAt(px, totalRows, yOfRow), rnW: rnW(), headerH: headerH() }),
    editable: () => wb.kind === "xlsx" || wb.kind === "ods",
    onEdit: () => { mark(); imageLayer.refresh(); },
    onReplace: (im) => replaceImage(im),
  });
  // Replace a picture's bytes: pick a new image file, swap the data URI + stage the bytes for the
  // writer (keeping the anchor). The media part is rewritten on save.
  const replaceImage = (im: import("./model").SheetImage): void => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.style.display = "none";
    input.addEventListener("change", async () => {
      const file = input.files?.[0];
      input.remove();
      if (!file) return;
      const bytes = new Uint8Array(await file.arrayBuffer());
      const ext = (file.name.split(".").pop() || (file.type.split("/")[1] ?? "png")).toLowerCase();
      const mime = file.type || `image/${ext}`;
      let bin = "";
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      im.replaceBytes = bytes;
      im.replaceExt = ext;
      im.dataUri = `data:${mime};base64,${btoa(bin)}`;
      im.dirty = true;
      mark();
      imageLayer.refresh();
    });
    document.body.appendChild(input);
    input.click();
  };
  const shapeLayer = setupShapeLayer({
    wrap,
    panes: () => panes.map((p) => ({ el: p.scrollEl, header: p.header, rowHeader: p.rowHeader })),
    getSheet: () => wb.sheets[active],
    geom: () => ({ xOfCol, yOfRow, colAt: (px) => lineAt(px, totalCols, xOfCol), rowAt: (px) => lineAt(px, totalRows, yOfRow), rnW: rnW(), headerH: headerH() }),
    editable: () => wb.kind === "xlsx" || wb.kind === "ods",
    onEdit: () => { mark(); shapeLayer.refresh(); },
    onActivate: (sh) => openShapeDialog(sh),
    onDelete: (sh) => deleteShape(sh),
  });
  // Remove a shape: drop it from the model, and (for a saved shape) stage the drawing edit.
  const deleteShape = (sh: import("./model").SheetShape): void => {
    const sheet = wb.sheets[active]!;
    const arr = sheet.shapes;
    const i = arr?.indexOf(sh) ?? -1;
    if (!arr || i < 0) return;
    arr.splice(i, 1);
    if (!arr.length) sheet.shapes = undefined;
    if (wb.kind === "xlsx" && sh.drawingPath != null && sh.anchorIndex != null) deleteXlsxShape(wb, sheet, sh);
    else if (wb.kind === "ods" && sh.odsShapeEl) sh.odsShapeEl.parentNode?.removeChild(sh.odsShapeEl);
    mark(); shapeLayer.refresh();
  };
  const slicerLayer = setupSlicerLayer({
    wrap,
    panes: () => panes.map((p) => ({ el: p.scrollEl, header: p.header, rowHeader: p.rowHeader })),
    getSheet: () => wb.sheets[active],
    getWorkbook: () => wb,
    geom: () => ({ xOfCol, yOfRow, colAt: (px) => lineAt(px, totalCols, xOfCol), rowAt: (px) => lineAt(px, totalRows, yOfRow), rnW: rnW(), headerH: headerH() }),
    onChange: (sl) => { applySlicer(sl); mark(); slicerLayer.refresh(); },
  });
  /** "$D$1:$D$3" (optionally sheet-qualified) -> a 1-based inclusive range. */
  const parseRangeRefLocal = (ref: string): { r1: number; c1: number; r2: number; c2: number } | null => {
    const body = (ref.includes("!") ? ref.slice(ref.lastIndexOf("!") + 1) : ref).replace(/\$/g, "");
    const [a, b] = body.split(":");
    const p1 = parseA1Ref(a ?? "");
    const p2 = b ? parseA1Ref(b) : p1;
    if (!p1 || !p2) return null;
    return { r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) };
  };

  // Form controls. A control exists to drive its linked cell, so a change writes there and
  // recalculates, exactly as it would in Excel.
  // Only xlsx models controls this way; ODF keeps them in office:forms, which is preserved rather
  // than rendered, so the layer is not built at all for other formats.
  const controlLayer = setupControlLayer({
    wrap,
    panes: () => panes.map((p) => ({ el: p.scrollEl, header: p.header, rowHeader: p.rowHeader })),
    getSheet: () => wb.sheets[active],
    geom: () => ({ xOfCol, yOfRow, colAt: (px) => lineAt(px, totalCols, xOfCol), rowAt: (px) => lineAt(px, totalRows, yOfRow), rnW: rnW(), headerH: headerH() }),
    inertTitle: t("ctrlMacroInert"),
    macroTitle: t("ctrlMacroRun"),
    activeXTitle: t("ctrlActiveX"),
    runMacro: (name) => runControlMacro(name),
    // Only xlsx writes controls back, so only there is dragging one offered.
    editable: () => wb.kind === "xlsx",
    onPlace: (ctl) => {
      const sheet = wb.sheets[active];
      if (!sheet) return;
      updateXlsxControlLinks(wb, ctl, sheet);
      mark();
      renderGrid();
    },
    itemsFor: (ctl) => {
      // An ActiveX listFillRange is usually a DEFINED NAME rather than a reference, which is what
      // the real files use, so the name is resolved before the reference is parsed.
      const raw = ctl.sourceRange
        ? wb.definedNames?.get(ctl.sourceRange) ?? wb.definedNames?.get(ctl.sourceRange.toUpperCase()) ?? ctl.sourceRange
        : undefined;
      // A named range may name another sheet; its own sheet is where the items live.
      const named = raw?.includes("!") ? raw.split("!") : undefined;
      const sheet = named ? wb.sheets.find((s) => s.name === named[0]!.replace(/^'|'$/g, "")) ?? wb.sheets[active] : wb.sheets[active];
      const rng = raw ? parseRangeRefLocal(named ? named[1]! : raw) : null;
      if (!sheet || !rng) return [];
      const out: string[] = [];
      for (let r = rng.r1; r <= rng.r2; r++)
        for (let c = rng.c1; c <= rng.c2; c++) out.push(displayValue(sheet, r, c));
      return out;
    },
    onChange: (changed) => {
      const sheet = wb.sheets[active];
      // An ActiveX control keeps its own state in a persisted binary, so the change goes there as
      // well as to any linked cell. A part whose layout was not understood is left alone.
      for (const ctl of changed) {
        if (!ctl.activeX || !ctl.activeXBinPath || ctl.activeXValue === undefined) continue;
        const bin = wb.files[ctl.activeXBinPath];
        const next = bin ? setActiveXValue(bin, ctl.activeXValue) : undefined;
        if (next) wb.files[ctl.activeXBinPath] = next;
      }
      // A control with no linked cell still remembers its own state; there is just nowhere to put
      // it. A radio that cleared its group brings the rest along, in one undo step.
      const writes = changed.flatMap((ctl) => {
        const at = ctl.linkedCell ? parseA1Ref(ctl.linkedCell.replace(/\$/g, "").split("!").pop() ?? "") : null;
        if (!at) return [];
        // An ActiveX list writes the chosen TEXT to its linked cell where a form control writes the
        // 1-based index, which is the difference between the two families that actually shows.
        const value = ctl.kind === "checkbox" || ctl.kind === "radio"
          ? (ctl.checked ? "TRUE" : "FALSE")
          : ctl.kind === "dropdown" || ctl.kind === "list"
            ? (ctl.activeX ? ctl.activeXValue ?? "" : String(ctl.selected ?? 0))
            : String(ctl.value ?? 0);
        return [{ r: at.row, c: at.col, value }];
      });
      if (sheet && writes.length) {
        recordCells(writes.map((w) => ({ r: w.r, c: w.c })), () => {
          for (const w of writes) setCellInput(sheet, w.r, w.c, w.value);
        });
        recalc(wb);
      }
      mark();
      renderGrid();
    },
  });

  // --- the workbook's own event macros -----------------------------------------
  // Excel runs Workbook_Open and Worksheet_Change by itself. sheetedit does not, unless the user
  // turns it on: a workbook that runs code the moment it opens is exactly what Excel's own
  // "enable content" bar exists for. The consent lasts for this session and is never persisted.
  let eventsEnabled = false;
  let inEvent = false;

  /** Run one event handler, if there is one and the user allowed it. */
  const fireEvent = (handler: { module: string; source: string; proc: string } | undefined, target?: { sheetIndex: number; rect: { r1: number; c1: number; r2: number; c2: number } }): void => {
    if (!eventsEnabled || !handler || inEvent) return;
    // A Worksheet_Change handler that writes a cell would otherwise fire itself, for ever.
    inEvent = true;
    try {
      const res = runWorkbookMacro(wb, handler.source, handler.module, handler.proc, {
        activeSheet: active,
        selection: sel ? { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 } : undefined,
        activeCell: activeCell ?? undefined,
        fileName: options.fileName,
        ...(target ? { eventTarget: target } : {}),
      });
      if (!res.ok) { showNotice(`${t("vbaEventFailed")} ${res.error?.message ?? ""}`); return; }
      if (res.undo && res.redo) {
        history.push({ sheet: active, cells: [], undoExtra: res.undo, redoExtra: res.redo });
        recalc(wb);
        mark();
        renderGrid();
      }
      if (res.messages.length) showNotice(res.messages.join(" | "));
    } finally {
      inEvent = false;
    }
  };

  /** Worksheet_Change, with the cells that changed as its Target. */
  const fireSheetChange = (rect: { r1: number; c1: number; r2: number; c2: number }): void => {
    if (!eventsEnabled || inEvent) return;
    fireEvent(findSheetHandler(wb, active, "Worksheet_Change"), { sheetIndex: active, rect });
  };

  /**
   * Run the macro a control names. Excel qualifies the name with its module when it needs to; the
   * VML usually carries the bare Sub name, so the module that declares it is looked up here.
   */
  const runControlMacro = (name: string): void => {
    const bare = name.replace(/^.*\./, "");
    const mod = wb.vba?.modules.find((m) => runnableSubs(m.source).some((s) => s.toLowerCase() === bare.toLowerCase()));
    if (!mod) { showNotice(`${t("ctrlMacroMissing")} ${name}`); return; }
    const proc = runnableSubs(mod.source).find((s) => s.toLowerCase() === bare.toLowerCase())!;
    showNotice(runMacro(mod.source, mod.name, proc).text);
  };

  /** Add a control at the selection, linked to the cell below it by default. */
  const insertControl = (kind: SheetControl["kind"]): void => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    // Where the user is looking, in order of how clearly they asked for it. With no selection at
    // all, drop it clear of the used range: defaulting to A1 lands it on top of the header row,
    // which is the one place a control is guaranteed to be in the way.
    const at = activeCell ?? (sel ? { r: sel.r1, c: sel.c1 } : { r: 1, c: Math.max(1, sheet.maxCol + 2) });
    const created = createXlsxControl(wb, sheet, {
      kind,
      label: kind === "checkbox" || kind === "button" ? t("ctrlNewLabel") : undefined,
      // Default to the cell just right of the control, which is where a reader looks for its value.
      linkedCell: defaultLink(at.r, at.c + 3),
      sourceRange: kind === "dropdown" ? undefined : undefined,
      at: placementFor(at.r, at.c),
    });
    if (!created) return;
    mark();
    renderGrid();
    openControlDialog(created);
  };

  /** Edit a control's label, linked cell and item source, or delete it. */
  const openControlDialog = (ctl: SheetControl): void => {
    const fields: FormField[] = [
      { key: "note", label: t("ctrlNote"), type: "note" },
      { key: "label", label: t("ctrlLabel"), type: "text", value: ctl.label ?? "" },
      { key: "link", label: t("ctrlLinkedCell"), type: "text", value: (ctl.linkedCell ?? "").replace(/\$/g, "") },
    ];
    if (ctl.kind === "dropdown" || ctl.kind === "list")
      fields.push({ key: "range", label: t("ctrlSourceRange"), type: "text", value: (ctl.sourceRange ?? "").replace(/\$/g, "") });
    // Only offered when the workbook has macros to offer, and only the Subs that can run alone.
    const macros = (wb.vba?.modules ?? []).flatMap((m) => runnableSubs(m.source));
    if (macros.length || ctl.macro) {
      const options = [{ value: "", label: t("ctrlMacroNone") }, ...macros.map((m) => ({ value: m, label: m }))];
      // A macro the workbook no longer has still shows, so assigning it away is possible.
      if (ctl.macro && !macros.includes(ctl.macro)) options.push({ value: ctl.macro, label: ctl.macro });
      fields.push({ key: "macro", label: t("ctrlMacro"), type: "select", options, value: ctl.macro ?? "" });
    }
    fields.push({ key: "remove", label: t("ctrlDelete"), type: "checkbox", value: false });
    formDialog(wrap, t("ctrlEdit"), fields, (vals) => {
      const sheet = wb.sheets[active];
      if (!sheet) return;
      if (vals.remove) {
        deleteXlsxControl(wb, sheet, ctl);
        mark();
        renderGrid();
        return;
      }
      const link = String(vals.link).trim();
      const range = vals.range == null ? "" : String(vals.range).trim();
      // A reference that does not parse is rejected rather than written as nonsense.
      if (link && !absoluteRef(link)) { showNotice(t("ctrlBadRef")); return; }
      if (range && !absoluteRange(range)) { showNotice(t("ctrlBadRef")); return; }
      ctl.label = String(vals.label).trim() || undefined;
      ctl.linkedCell = link ? absoluteRef(link) : undefined;
      if (vals.range != null) ctl.sourceRange = range ? absoluteRange(range) : undefined;
      if (vals.macro != null) ctl.macro = String(vals.macro) || undefined;
      updateXlsxControlLinks(wb, ctl, sheet);
      mark();
      renderGrid();
    });
  };

  // The row-outline gutter, left of the row numbers.
  const outlineLayer = setupOutlineLayer({
    wrap,
    gridScroll,
    getSheet: () => wb.sheets[active],
    geom: () => {
      const hh = headerH();
      return {
        headerH: hh,
        totalRows,
        rowRect: (r: number) => {
          const th = gridScroll.querySelector(`th.rownum[data-r="${r}"]`) as HTMLElement | null;
          if (!th) return null;
          const gr = gridScroll.getBoundingClientRect();
          const rc = th.getBoundingClientRect();
          return { top: rc.top - gr.top + gridScroll.scrollTop - hh, height: rc.height };
        },
      };
    },
    onToggle: (level, line, collapse) => {
      const sheet = wb.sheets[active];
      if (!sheet) return;
      recordSettings(() => {
        setGroupCollapsed(sheet, "row", line, level, collapse, totalRows);
        if (wb.kind === "ods") sheet.odsDirty = true;
      });
      mark(); renderGrid();
    },
    onLevel: (level) => {
      const sheet = wb.sheets[active];
      if (!sheet) return;
      recordSettings(() => {
        showOutlineLevel(sheet, "row", level, totalRows);
        if (wb.kind === "ods") sheet.odsDirty = true;
      });
      mark(); renderGrid();
    },
  });
  // The frozen / split boundary bars: drag to move, double-click to remove. Every measurement comes
  // off the rendered headers, since the declared line sizes and the rendered ones differ.
  const paneDividers = setupPaneDividers({
    wrap,
    panes: () => panes.map((p) => p.scrollEl),
    getSheet: () => wb.sheets[active],
    geom: () => {
      const headEl = (sel: string): HTMLElement | null => wrap.querySelector(sel) as HTMLElement | null;
      const nearest = (sel: string, attr: string, coord: (r: DOMRect) => number, size: (r: DOMRect) => number, client: number): number => {
        let best = 0, bestD = Infinity;
        for (const el of Array.from(wrap.querySelectorAll(sel)) as HTMLElement[]) {
          const line = Number(el.dataset[attr] ?? "0");
          if (!line) continue;
          const r = el.getBoundingClientRect();
          for (const [edge, n] of [[coord(r), line - 1], [coord(r) + size(r), line]] as const) {
            const d = Math.abs(edge - client);
            if (d < bestD) { bestD = d; best = n; }
          }
        }
        return best;
      };
      return {
        // A row split's boundary is the seam between the two containers; a freeze's is the top edge
        // of the first row below it, or the bottom edge of the last row above when that is offscreen.
        rowBoundaryY: (rows: number) => {
          if (panes.includes(splitPane)) return splitScroll.getBoundingClientRect().top;
          const below = headEl(`th.rownum[data-r="${rows + 1}"]`);
          if (below) return below.getBoundingClientRect().top;
          const above = headEl(`th.rownum[data-r="${rows}"]`);
          return above ? above.getBoundingClientRect().bottom : null;
        },
        colBoundaryX: (cols: number) => {
          if (panes.includes(rightPane)) return rightPane.scrollEl.getBoundingClientRect().left;
          const after = headEl(`th.colhead[data-c="${cols + 1}"]`);
          if (after) return after.getBoundingClientRect().left;
          const before = headEl(`th.colhead[data-c="${cols}"]`);
          return before ? before.getBoundingClientRect().right : null;
        },
        bodyTop: () => {
          const r1 = headEl('th.rownum[data-r="1"]');
          return r1 ? r1.getBoundingClientRect().top : gridScroll.getBoundingClientRect().top;
        },
        nearestRow: (clientY: number) => nearest("th.rownum", "r", (r) => r.top, (r) => r.height, clientY),
        nearestCol: (clientX: number) => nearest("th.colhead", "c", (r) => r.left, (r) => r.width, clientX),
      };
    },
    onMove: (rows, cols) => setFreeze(rows, cols, wb.sheets[active]?.paneSplit),
  });
  const timelineLayer = setupTimelineLayer({
    wrap,
    panes: () => panes.map((p) => ({ el: p.scrollEl, header: p.header, rowHeader: p.rowHeader })),
    getSheet: () => wb.sheets[active],
    geom: () => ({ xOfCol, yOfRow, colAt: (px) => lineAt(px, totalCols, xOfCol), rowAt: (px) => lineAt(px, totalRows, yOfRow), rnW: rnW(), headerH: headerH() }),
    onChange: (tl) => { applyTimeline(tl); mark(); timelineLayer.refresh(); },
  });
  const pivotLayer = setupPivotLayer({
    wrap,
    panes: () => panes.map((p) => ({ el: p.scrollEl, header: p.header, rowHeader: p.rowHeader })),
    getSheet: () => wb.sheets[active],
    geom: () => ({ xOfCol, yOfRow, colAt: (px) => lineAt(px, totalCols, xOfCol), rowAt: (px) => lineAt(px, totalRows, yOfRow), rnW: rnW(), headerH: headerH() }),
    label: (name) => t("pivotTag", { name }),
    onTag: (pivot, x, y) => openPivotMenu(pivot, x, y),
  });
  const chartUi = chartsOn
    ? setupChartUi({
        wrap,
        gridScroll,
        getWorkbook: () => wb,
        activeSheetName: () => wb.sheets[active]?.name ?? "Sheet1",
        onCreate: (m) => { (wb.sheets[active].charts ??= []).push(m); mark(); chartLayer.refresh(); },
        onDelete: (m) => { const cs = wb.sheets[active]?.charts; const i = cs?.indexOf(m) ?? -1; if (cs && i >= 0) cs.splice(i, 1); mark(); chartLayer.refresh(); },
        onChange: () => { mark(); chartLayer.refresh(); },
        boxRect: (id) => chartLayer.boxRect(id),
      })
    : { openInsert: () => undefined, showEdit: () => undefined, hideEdit: () => undefined, teardown: () => undefined };
  if (chartsOn) {
    const CHART_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 2v12h12"/><rect x="4" y="8" width="2.2" height="4"/><rect x="8" y="5" width="2.2" height="7"/><rect x="12" y="9" width="2.2" height="3"/></svg>`;
    trailingIcons.push(tbIcon(CHART_ICON, t("chartInsert"), () => chartUi.openInsert(getSelRect())));
  }
  // Authoring controls, each advertised only when the open format supports it (see capabilities.ts).
  if (caps.autofilter) {
    const FILTER_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 3h12l-4.5 5.5V13l-3 1.5V8.5z"/></svg>`;
    trailingIcons.push(tbIcon(FILTER_ICON, t("filterToggle"), () => toggleAutoFilter()));
  }
  if (caps.pivots) {
    const PIVOT_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="2" width="12" height="12" rx="1"/><path d="M2 6h12M6 6v8M6 6V2"/><path d="M9 9.5l2 2M11 9.5l-2 2"/></svg>`;
    trailingIcons.push(tbIcon(PIVOT_ICON, t("pivotInsert"), () => openPivotDialog()));
  }
  // A table slicer needs xlsx parts, so it rides on the same capability as pivots.
  if (caps.pivots && wb.kind === "xlsx") {
    const SLICER_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="2" width="10" height="12" rx="1.5"/><rect x="5" y="5" width="6" height="2.2" rx="0.6"/><rect x="5" y="9" width="6" height="2.2" rx="0.6"/></svg>`;
    trailingIcons.push(tbIcon(SLICER_ICON, t("slicerInsert"), () => addTableSlicer()));
  }
  if (caps.freezePanes) {
    const FREEZE_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="2" width="12" height="12" rx="1"/><path d="M2 6h12M6 2v12"/></svg>`;
    // A small menu, like Excel's: freeze at the cursor, or just the top row / first column.
    const freezeBtn: HTMLButtonElement = tbIcon(FREEZE_ICON, t("freezePanes"), () => {
      closeLineMenu();
      const sheet = wb.sheets[active];
      const cur = sel ? { r: sel.r1, c: sel.c1 } : { r: 1, c: 1 };
      const pop = document.createElement("div");
      pop.className = "sheetedit-pop";
      const item = (text: string, rows: number, cols: number, asSplit?: boolean) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-pop-item";
        b.textContent = text;
        b.addEventListener("click", () => { closeLineMenu(); setFreeze(rows, cols, asSplit); });
        pop.appendChild(b);
      };
      // "Freeze at this cell" pins everything above and left of the selection, as Excel does.
      item(t("freezeHere"), Math.max(0, cur.r - 1), Math.max(0, cur.c - 1));
      item(t("freezeTopRow"), 1, 0);
      item(t("freezeFirstCol"), 0, 1);
      // A split is the same boundary, but draggable and recorded as such where the format allows.
      item(t("splitHere"), Math.max(0, cur.r - 1), Math.max(0, cur.c - 1), true);
      if ((sheet?.freeze?.rows ?? 0) > 0 || (sheet?.freeze?.cols ?? 0) > 0) item(t("unfreeze"), 0, 0);
      document.body.appendChild(pop);
      lineMenu = pop;
      const r = freezeBtn.getBoundingClientRect();
      pop.style.left = `${Math.round(Math.min(r.left, window.innerWidth - pop.offsetWidth - 8))}px`;
      pop.style.top = `${Math.round(r.bottom + 4)}px`;
      const onOutside = (ev: Event) => {
        if (!pop.contains(ev.target as Node)) { closeLineMenu(); document.removeEventListener("pointerdown", onOutside, true); }
      };
      setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
    });
    trailingIcons.push(freezeBtn);
  }
  // Macros: shown when the workbook actually has some, so the button is never a dead end.
  if (wb.vba?.modules.length) {
    const VBA_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m5.5 5.5-3 2.5 3 2.5M10.5 5.5l3 2.5-3 2.5"/></svg>`;
    trailingIcons.push(tbIcon(VBA_ICON, t("vbaTitle"), () => openMacroViewer()));
  }
  if (caps.formControls) {
    const CTRL_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="3" width="5" height="5" rx="1"/><path d="m3.2 5.5 1.1 1.1 1.6-2"/><rect x="9" y="3" width="5" height="5" rx="1"/><path d="m10.2 5 1.3 1.3L12.8 5"/><path d="M2 11.5h12"/></svg>`;
    const ctrlBtn: HTMLButtonElement = tbIcon(CTRL_ICON, t("ctrlInsert"), () => {
      closeLineMenu();
      const pop = document.createElement("div");
      pop.className = "sheetedit-pop";
      const item = (text: string, run: () => void) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-pop-item";
        b.textContent = text;
        b.addEventListener("click", () => { closeLineMenu(); run(); });
        pop.appendChild(b);
      };
      item(t("ctrlAddCheckbox"), () => insertControl("checkbox"));
      item(t("ctrlAddDropdown"), () => insertControl("dropdown"));
      item(t("ctrlAddSpinner"), () => insertControl("spin"));
      const existing = wb.sheets[active]?.controls ?? [];
      if (existing.length) {
        const sep = document.createElement("div");
        sep.className = "sheetedit-pop-sep";
        pop.appendChild(sep);
        for (const ctl of existing) item(`${t("ctrlEdit")}: ${ctl.label ?? ctl.name}`, () => openControlDialog(ctl));
      }
      document.body.appendChild(pop);
      lineMenu = pop;
      const r = ctrlBtn.getBoundingClientRect();
      pop.style.left = `${Math.round(Math.min(r.left, window.innerWidth - pop.offsetWidth - 8))}px`;
      pop.style.top = `${Math.round(r.bottom + 4)}px`;
      const onOutside = (ev: Event) => {
        if (!pop.contains(ev.target as Node)) { closeLineMenu(); document.removeEventListener("pointerdown", onOutside, true); }
      };
      setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
    });
    trailingIcons.push(ctrlBtn);
  }
  if (caps.themes) {
    const THEME_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="8" cy="8" r="6"/><path d="M8 2a6 6 0 0 0 0 12 3 3 0 0 0 0-6 3 3 0 0 1 0-6z"/></svg>`;
    trailingIcons.push(tbIcon(THEME_ICON, t("themeTitle"), () => openThemeDialog()));
  }
  if (caps.printSetup) {
    const PRINT_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4.5 6.5V2.5h7v4"/><rect x="2" y="6.5" width="12" height="5" rx="1"/><path d="M4.5 10h7v3.5h-7z"/></svg>`;
    const printBtn: HTMLButtonElement = tbIcon(PRINT_ICON, t("printSetup"), () => {
      closeLineMenu();
      const p = wb.sheets[active]?.printSetup;
      const pop = document.createElement("div");
      pop.className = "sheetedit-pop";
      const item = (text: string, run: () => void) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-pop-item";
        b.textContent = text;
        b.addEventListener("click", () => { closeLineMenu(); run(); });
        pop.appendChild(b);
      };
      item(t("printNow"), () => openPrintScopeDialog());
      item(t("printSetup"), () => openPrintDialog());
      const sep = document.createElement("div");
      sep.className = "sheetedit-pop-sep";
      pop.appendChild(sep);
      // The print area comes from the selection, so it is only offered when there is one.
      if (sel) item(t("printAreaSet"), () => updatePrint((next) => { next.printArea = [{ ...sel! }]; }));
      if (p?.printArea?.length) item(t("printAreaClear"), () => updatePrint((next) => { next.printArea = undefined; }));
      if (p?.rowBreaks?.length || p?.colBreaks?.length)
        item(t("printBreaksClear"), () => updatePrint((next) => { next.rowBreaks = undefined; next.colBreaks = undefined; }));
      document.body.appendChild(pop);
      lineMenu = pop;
      const r = printBtn.getBoundingClientRect();
      pop.style.left = `${Math.round(Math.min(r.left, window.innerWidth - pop.offsetWidth - 8))}px`;
      pop.style.top = `${Math.round(r.bottom + 4)}px`;
      const onOutside = (ev: Event) => {
        if (!pop.contains(ev.target as Node)) { closeLineMenu(); document.removeEventListener("pointerdown", onOutside, true); }
      };
      setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
    });
    trailingIcons.push(printBtn);
  }
  if (caps.protection) {
    const LOCK_ICON =`<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="7" width="10" height="7" rx="1.5"/><path d="M5.5 7V4.8a2.5 2.5 0 0 1 5 0V7"/></svg>`;
    const protBtn: HTMLButtonElement = tbIcon(LOCK_ICON, t("protection"), () => {
      closeLineMenu();
      const sheet = wb.sheets[active];
      const pop = document.createElement("div");
      pop.className = "sheetedit-pop";
      const item = (text: string, run: () => void) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-pop-item";
        b.textContent = text;
        b.addEventListener("click", () => { closeLineMenu(); run(); });
        pop.appendChild(b);
      };
      if (isProtected(sheet)) item(t("unprotectSheet"), () => setSheetProtection(undefined));
      else item(t("protectSheet"), () => openProtectDialog());
      // Cell locking only makes sense while the sheet is unprotected (Excel behaves the same).
      if (!isProtected(sheet)) {
        const sep = document.createElement("div");
        sep.className = "sheetedit-pop-sep";
        pop.appendChild(sep);
        item(t("lockCells"), () => applyStyle({ locked: true }));
        item(t("unlockCells"), () => applyStyle({ locked: false }));
      }
      const sep2 = document.createElement("div");
      sep2.className = "sheetedit-pop-sep";
      pop.appendChild(sep2);
      if (isStructureLocked(wb)) item(t("unprotectStructure"), () => setStructureProtection(false));
      else item(t("protectStructure"), () => setStructureProtection(true));
      document.body.appendChild(pop);
      lineMenu = pop;
      const r = protBtn.getBoundingClientRect();
      pop.style.left = `${Math.round(Math.min(r.left, window.innerWidth - pop.offsetWidth - 8))}px`;
      pop.style.top = `${Math.round(r.bottom + 4)}px`;
      const onOutside = (ev: Event) => {
        if (!pop.contains(ev.target as Node)) { closeLineMenu(); document.removeEventListener("pointerdown", onOutside, true); }
      };
      setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
    });
    trailingIcons.push(protBtn);
  }
  if (caps.hyperlinks) {
    const LINK_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6.5 9.5 13 3M9.5 3H13v3.5M12 9.5V12a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h2.5"/></svg>`;
    trailingIcons.push(tbIcon(LINK_ICON, t("linkEdit"), () => openLinkDialog()));
  }
  if (caps.dataValidation) {
    const DV_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="3" width="12" height="10" rx="1.5"/><path d="M5 6.5h6M5 9.5h4M11 9l1 1 1.5-1.5"/></svg>`;
    trailingIcons.push(tbIcon(DV_ICON, t("dvEdit"), () => openDvDialog()));
  }
  if (caps.comments) {
    const NOTE_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 3h10a1 1 0 0 1 1 1v6a1 1 0 0 1-1 1H7l-3 2.5V11H3a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/></svg>`;
    trailingIcons.push(tbIcon(NOTE_ICON, t("noteEdit"), () => openNoteDialog()));
  }
  if (caps.conditionalFormat) {
    const CF_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="2" width="5" height="5" rx="1"/><rect x="9" y="2" width="5" height="5" rx="1"/><rect x="2" y="9" width="5" height="5" rx="1"/><rect x="9" y="9" width="5" height="5" rx="1"/></svg>`;
    trailingIcons.push(tbIcon(CF_ICON, t("cfEdit"), () => openCfDialog()));
  }
  if (caps.sparklines) {
    const SPARK_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 11l3-4 3 2 3-5 3 3"/></svg>`;
    trailingIcons.push(tbIcon(SPARK_ICON, t("sparkEdit"), () => openSparkDialog()));
  }
  if (caps.shapes) {
    const SHAPE_ICON = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="4" width="7" height="7" rx="1"/><circle cx="11.5" cy="10.5" r="3.5"/></svg>`;
    trailingIcons.push(tbIcon(SHAPE_ICON, t("shapeInsert"), () => openShapeDialog()));
  }
  toolbarHandle.setTrailing(trailingIcons);
  syncToolbar = () => toolbarHandle.syncActive();
  syncToolbar();

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
      const td = tdAt(key(pos.r, pos.c));
      const input = inputAt(key(pos.r, pos.c));
      if (td && input) applyCellVisualStyle(td, input, getCell(sheet, pos.r, pos.c));
    }
  };

  // Draw a mini line / column / win-loss sparkline into a cell-sized canvas.
  const drawSparkline = (cv: HTMLCanvasElement, type: "line" | "column" | "stacked", color: string, valuesRaw: number[], negColor = "#d1493f"): void => {
    const w = cv.clientWidth || 60, h = cv.clientHeight || 18;
    const dpr = window.devicePixelRatio || 1;
    cv.width = Math.round(w * dpr); cv.height = Math.round(h * dpr);
    const ctx = cv.getContext("2d"); if (!ctx) return;
    ctx.scale(dpr, dpr);
    const vals = valuesRaw.filter((v) => !isNaN(v));
    if (!vals.length) return;
    const pad = 1;
    if (type === "stacked") { // win/loss: equal up/down bars from the middle
      const n = valuesRaw.length, bw = (w - pad * 2) / n, mid = h / 2;
      valuesRaw.forEach((v, i) => { if (isNaN(v) || v === 0) return; const x = pad + i * bw; ctx.fillStyle = v > 0 ? color : negColor; const bh = h * 0.32; ctx.fillRect(x + bw * 0.15, v > 0 ? mid - bh : mid, bw * 0.7, bh); });
      return;
    }
    const lo = Math.min(...vals), hi = Math.max(...vals), span = hi - lo || 1;
    const yOf = (v: number): number => h - pad - ((v - lo) / span) * (h - pad * 2);
    const n = valuesRaw.length, step = n > 1 ? (w - pad * 2) / (n - 1) : 0;
    if (type === "column") {
      const bw = (w - pad * 2) / n;
      valuesRaw.forEach((v, i) => { if (isNaN(v)) return; const x = pad + i * bw; const y = yOf(v); ctx.fillStyle = v < 0 ? negColor : color; ctx.fillRect(x + bw * 0.15, y, bw * 0.7, h - pad - y); });
    } else {
      ctx.strokeStyle = color; ctx.lineWidth = 1; ctx.lineJoin = "round"; ctx.beginPath();
      let started = false;
      valuesRaw.forEach((v, i) => { if (isNaN(v)) return; const x = pad + i * step, y = yOf(v); if (started) ctx.lineTo(x, y); else { ctx.moveTo(x, y); started = true; } });
      ctx.stroke();
    }
  };

  // Build one data cell's <td> (input, styles, listeners). Extracted from buildRow so a
  // frozen column can reuse it outside the horizontal window loop.
  const buildCell = (sheet: Sheet, r: number, c: number): HTMLTableCellElement => {
      const td = document.createElement("td");
      td.dataset.rc = key(r, c);
      cur.tds.set(key(r, c), td);
      const sp = spanAtMap.get(key(r, c));
      if (sp) {
        if (sp.rs > 1) td.rowSpan = sp.rs;
        if (sp.cs > 1) td.colSpan = sp.cs;
      }
      const cell = getCell(sheet, r, c);
      if (cell?.kind === "n") td.classList.add("num");
      markPrintEdges(td, sheet, r, c);
      const sheetProtected = isProtected(sheet);
      const input = document.createElement("input");
      input.type = "text";
      input.value = cellDisplay(cell);
      input.setAttribute("aria-label", `${colToLetters(c)}${r}`);
      // A locked cell on a protected sheet stays selectable and copyable (Excel's default) but
      // refuses typing. readOnly does that natively, keyboard navigation included.
      if (sheetProtected && !cell?.cellStyle?.unlocked) {
        input.readOnly = true;
        td.classList.add("locked");
        input.title = t("protectedCell");
      }
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
        const isList = (dv.type ?? "list") === "list";
        const cur = cellDisplay(cell);
        const allowed = isList ? resolveDvValues(dv, sheet) : [];
        // Flag a cell whose value breaks the rule (list membership, or the typed constraint).
        if (cur !== "" && !validateCell(dv, cell?.value ?? "", cur, allowed)) td.classList.add("sheetedit-dv-invalid");
        // A dropdown caret only for list rules; the constraint rules just outline invalid input.
        if (isList) {
          td.classList.add("has-dv");
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
      }
      // Autofilter: a caret on each header-row cell that opens the sort/filter menu.
      const af = sheet.autoFilter;
      if (af && r === af.r1 && c >= af.c1 && c <= af.c2) {
        td.classList.add("has-filter");
        if (sheet.filters?.has(c)) td.classList.add("sheetedit-filter-on");
        const fb = document.createElement("button");
        fb.type = "button";
        fb.className = "sheetedit-filterbtn";
        fb.tabIndex = -1;
        fb.title = t("filterToggle");
        fb.setAttribute("aria-label", t("filterToggle"));
        fb.innerHTML = `<svg viewBox="0 0 16 16" width="11" height="11" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 3h12l-4.5 5.5V13l-3 1.5V8.5z"/></svg>`;
        fb.addEventListener("mousedown", (e) => e.preventDefault());
        fb.addEventListener("click", (e) => { e.stopPropagation(); openFilterMenu(td, c); });
        td.appendChild(fb);
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
        if (cfv.icon) {
          td.classList.add("has-cficon");
          const ic = document.createElement("span");
          ic.className = "sheetedit-cficon";
          ic.innerHTML = cfv.icon;
          td.insertBefore(ic, td.firstChild);
        }
      }
      // Sparkline: a mini line/column/win-loss chart drawn into the host cell.
      const spk = sparkAt.get(key(r, c));
      if (spk) {
        td.classList.add("has-spark");
        const cv = document.createElement("canvas");
        cv.className = "sheetedit-spark";
        const vals = resolveNumbers(wb, { ref: spk.dataRef }).map((v) => (v == null ? NaN : v));
        requestAnimationFrame(() => drawSparkline(cv, spk.type, spk.color, vals, spk.negColor ?? undefined));
        td.insertBefore(cv, td.firstChild);
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
        if (cur.inputs.get(ki) !== input) return;
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
      // Rich text: a display overlay of per-run styled spans (the input keeps the plain text for
      // editing; CSS hides the overlay while the cell is focused).
      if (cell?.richRuns?.length) {
        td.classList.add("has-rich");
        const ov = document.createElement("div");
        ov.className = "sheetedit-cellrich";
        ov.setAttribute("aria-hidden", "true");
        for (const run of cell.richRuns) {
          const sp = document.createElement("span");
          sp.textContent = run.text;
          if (run.bold) sp.style.fontWeight = "700";
          if (run.italic) sp.style.fontStyle = "italic";
          const deco = `${run.underline ? "underline " : ""}${run.strike ? "line-through" : ""}`.trim();
          if (deco) sp.style.textDecoration = deco;
          if (run.size) sp.style.fontSize = `${run.size}pt`;
          if (run.color) sp.style.color = run.color;
          if (run.font) sp.style.fontFamily = run.font;
          ov.appendChild(sp);
        }
        if (cell.cellStyle?.align) ov.style.textAlign = cell.cellStyle.align;
        td.appendChild(ov);
      } else if (cell?.phonetic?.length) {
        // Furigana: render the phonetic guide as ruby in a display overlay. The input keeps the
        // base text (edited/saved as-is); CSS shows the ruby until the cell is focused for editing.
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
      cur.inputs.set(ki, input);
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
    const withRowNums = cur.rowHeader; // the band right of a column split continues past them
    tr.style.height = `${effRowH(sheet, r)}px`;
    const frozenRow = r <= fz.fr;
    const rowTop = fz.headerH + yOfRow(r); // where a frozen row sticks, just below the header
    const rn = document.createElement("th");
    rn.className = "rownum";
    rn.dataset.r = String(r); // the outline gutter measures rows off these
    if (brkRows.has(r)) rn.classList.add("pgbrk-top"); // carry the break line into the gutter
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
    if (withRowNums) tr.appendChild(rn);
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
    if (!sheet || renderingWindow) return;
    renderingWindow = true;
    const was = cur;
    try {
      for (const p of panes) {
        if (!p.tableEl) continue;
        cur = p;
        // Explicit coordinates belong to the pane that asked; the others use their own scroll.
        renderWindowInner(force, p === was ? yAt : undefined, p === was ? xAt : undefined);
      }
    } finally {
      cur = was;
      renderingWindow = false;
    }
  };
  /** Re-render just one pane (its own scroll moved). */
  const renderPane = (pane: Pane): void => {
    if (!pane.tableEl || renderingWindow) return;
    renderingWindow = true;
    const was = cur;
    try { cur = pane; renderWindowInner(false); } finally { cur = was; renderingWindow = false; }
  };
  const renderWindowInner = (force: boolean, yAt?: number, xAt?: number): void => {
    const sheet = wb.sheets[active]!;
    const tableEl = cur.tableEl!;
    const scroller = cur.scrollEl;
    const keepTop = scroller.scrollTop;
    const keepLeft = scroller.scrollLeft;
    const y = yAt ?? keepTop;
    const x = Math.max(0, (xAt ?? keepLeft) - rnW()); // grid area starts after the row-number column
    let r1 = Math.max(1, lineAt(y, totalRows, yOfRow) - OVERSCAN);
    let r2 = Math.min(totalRows, lineAt(y + (scroller.clientHeight || viewportH()), totalRows, yOfRow) + OVERSCAN);
    let c1 = Math.max(1, lineAt(x, totalCols, xOfCol) - OVERSCAN_COLS);
    let c2 = Math.min(totalCols, lineAt(x + (scroller.clientWidth || viewportW()), totalCols, xOfCol) + OVERSCAN_COLS);
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
    if (!force && r1 >= cur.winR1 && r2 <= cur.winR2 && c1 >= cur.winC1 && c2 <= cur.winC2) return;

    // Keep an in-progress edit alive across the re-render when its cell stays
    // near the window; a far-away edit commits (blur) before its DOM goes away.
    let pin: { r: number; c: number; val: string; ss: number | null; se: number | null } | null = null;
    const ae = document.activeElement;
    if (ae instanceof HTMLInputElement && scroller.contains(ae)) {
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

    cur.inputs = new Map();
    cur.tds = new Map();
    tableEl.textContent = "";

    // Column skeleton: row numbers, the frozen columns, a left spacer for the window's
    // horizontal offset, the window's columns, a right spacer. Frozen rows/columns render
    // regardless of the scroll position; the window covers only the rest.
    // With a SPLIT the boundary is between real viewports, so nothing is sticky inside any of them;
    // a freeze keeps the sticky rows / columns as before.
    const fr = sheet.paneSplit ? 0 : sheet.freeze?.rows ?? 0;
    const fc = sheet.paneSplit ? 0 : sheet.freeze?.cols ?? 0;
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
    const paneRnW = cur.rowHeader ? rnW() : 0;
    if (cur.rowHeader) addCol(rnW());
    for (let c = 1; c <= fc; c++) if (!sheet.hiddenCols?.has(c)) colElByC.set(c, addCol(effColW(sheet, c)));
    addCol(leftW);
    for (let c = ec1; c <= c2; c++) if (!sheet.hiddenCols?.has(c)) colElByC.set(c, addCol(effColW(sheet, c)));
    addCol(rightW);
    tableEl.appendChild(colgroup);
    tableEl.style.width = `${paneRnW + gridW}px`;

    // A column header cell (letter, select-column click, resize grip).
    const makeColHead = (c: number, colEl: HTMLElement): HTMLTableCellElement => {
      const th = document.createElement("th");
      th.className = "colhead";
      th.dataset.c = String(c); // the pane divider snaps to these edges
      if (brkCols.has(c)) th.classList.add("pgbrk-left");
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
    if (cur.rowHeader) head.appendChild(corner);
    for (let c = 1; c <= fc; c++) {
      if (sheet.hiddenCols?.has(c)) continue;
      const th = makeColHead(c, colElByC.get(c)!);
      freezeCell(th, { left: rnW() + xOfCol(c), z: 9 }); // header is already sticky-top
      head.appendChild(th);
    }
    head.appendChild(document.createElement("th")); // left spacer
    for (let c = ec1; c <= c2; c++) if (!sheet.hiddenCols?.has(c)) head.appendChild(makeColHead(c, colElByC.get(c)!));
    head.appendChild(document.createElement("th")); // right spacer
    // The pane below a split continues under the header the top pane already draws.
    if (cur.header) tableEl.appendChild(head);

    // Frozen rows stick just below the header; measure the header height only when
    // needed (the read forces a layout on the still-empty table).
    const fz = { fr, fc, headerH: fr > 0 ? head.offsetHeight || 0 : 0 };

    const cellCols = fc + (c2 - ec1 + 1) + 3; // rownum + frozen cols + spacer + window + spacer
    for (let r = 1; r <= fr; r++) if (rowShown(sheet, r)) tableEl.appendChild(buildRow(sheet, r, ec1, c2, fz)); // frozen rows
    const topSpacer = document.createElement("tr");
    topSpacer.appendChild(document.createElement("td")).colSpan = cellCols;
    topSpacer.style.height = `${Math.max(0, yOfRow(er1) - yOfRow(fr + 1))}px`;
    tableEl.appendChild(topSpacer);
    for (let r = er1; r <= r2; r++) if (rowShown(sheet, r)) tableEl.appendChild(buildRow(sheet, r, ec1, c2, fz));
    const bottomSpacer = document.createElement("tr");
    bottomSpacer.appendChild(document.createElement("td")).colSpan = cellCols;
    bottomSpacer.style.height = `${Math.max(0, yOfRow(totalRows + 1) - yOfRow(r2 + 1))}px`;
    tableEl.appendChild(bottomSpacer);

    cur.winR1 = r1;
    cur.winR2 = r2;
    cur.winC1 = c1;
    cur.winC2 = c2;

    // Any layout forced while the table was empty clamps the scroll position to 0;
    // the spacers are back, so put it back too.
    if (scroller.scrollTop !== keepTop) scroller.scrollTop = keepTop;
    if (scroller.scrollLeft !== keepLeft) scroller.scrollLeft = keepLeft;

    if (pin) {
      const inp = inputAt(key(pin.r, pin.c));
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

  // ---- Sort & filter (autofilter) ----
  const columnValues = (sheet: Sheet, col: number): string[] => columnFilterValues(sheet, col);
  // Recompute which data rows are hidden from the per-column value filters, write the row-hidden
  // state, and re-render.
  const recomputeFilter = (sheet: Sheet): void => {
    const af = sheet.autoFilter; if (!af) return;
    const hidden = filterHiddenRows(sheet);
    for (let r = af.r1 + 1; r <= af.r2; r++) {
      const wasHidden = sheet.filterHidden?.has(r) ?? false;
      if (hidden.has(r) !== wasHidden && wb.kind === "xlsx") setXlsxRowHidden(sheet, r, hidden.has(r));
    }
    sheet.filterHidden = hidden.size ? hidden : undefined;
    // ODS persists row visibility from the model on save; flag the sheet so it is re-emitted.
    if (wb.kind === "ods") sheet.odsDirty = true;
    mark();
    renderGrid();
  };
  // Sort the autofilter data rows by a column. The ordering itself lives in range-ops, shared with
  // VBA's Range.Sort so the two cannot disagree about what "sorted" means.
  const sortByColumn = (sheet: Sheet, col: number, asc: boolean): void => {
    const af = sheet.autoFilter; if (!af) return;
    const r0 = af.r1 + 1;
    if (af.r2 <= r0) return;
    // Sorting rewrites the data rows, so it needs both the sort permission and writable cells.
    if (!allowAction("sort")) return;
    if (!allowRangeEdit({ r1: r0, c1: af.c1, r2: af.r2, c2: af.c2 })) return;
    const rect = { r1: af.r1, c1: af.c1, r2: af.r2, c2: af.c2 };
    recordCells(sortedPositions(rect, true), () => sortRange(sheet, rect, [{ col, ascending: asc }], true));
    recalc(wb);
    sheet.filters = undefined; sheet.filterHidden = undefined; // sort invalidates the value filter
    mark();
    renderGrid();
  };

  function openFilterMenu(td: HTMLElement, col: number): void {
    const sheet = wb.sheets[active]!;
    const menu = document.createElement("div");
    menu.className = "sheetedit-pop sheetedit-filtermenu";
    const btn = (label: string, fn: () => void): void => { const b = document.createElement("button"); b.type = "button"; b.className = "sheetedit-pop-item"; b.textContent = label; b.addEventListener("click", () => { menu.remove(); fn(); }); menu.appendChild(b); };
    btn(t("filterSortAsc"), () => sortByColumn(sheet, col, true));
    btn(t("filterSortDesc"), () => sortByColumn(sheet, col, false));
    const sep = document.createElement("div"); sep.className = "sheetedit-pop-sep"; menu.appendChild(sep);
    const values = columnValues(sheet, col);
    const active0 = sheet.filters?.get(col);
    const list = document.createElement("div"); list.className = "sheetedit-filter-list";
    const boxes: HTMLInputElement[] = [];
    const allWrap = document.createElement("label"); allWrap.className = "sheetedit-filter-opt";
    const allCb = document.createElement("input"); allCb.type = "checkbox"; allCb.checked = !active0;
    allCb.addEventListener("change", () => boxes.forEach((b) => { b.checked = allCb.checked; }));
    allWrap.append(allCb, Object.assign(document.createElement("span"), { textContent: t("filterSelectAll") }));
    list.appendChild(allWrap);
    for (const v of values) {
      const l = document.createElement("label"); l.className = "sheetedit-filter-opt";
      const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !active0 || active0.has(v);
      boxes.push(cb);
      l.append(cb, Object.assign(document.createElement("span"), { textContent: v === "" ? t("filterBlanks") : v }));
      list.appendChild(l);
    }
    menu.appendChild(list);
    const foot = document.createElement("div"); foot.className = "sheetedit-filter-foot";
    const ok = document.createElement("button"); ok.type = "button"; ok.className = "sheetedit-chart-btn primary"; ok.textContent = t("chartApply");
    ok.addEventListener("click", () => {
      menu.remove();
      const allowed = new Set<string>(); values.forEach((v, i) => { if (boxes[i]!.checked) allowed.add(v); });
      sheet.filters ??= new Map();
      if (allowed.size === values.length) sheet.filters.delete(col); else sheet.filters.set(col, allowed);
      if (!sheet.filters.size) sheet.filters = undefined;
      recomputeFilter(sheet);
    });
    foot.appendChild(ok);
    menu.appendChild(foot);
    wrap.appendChild(menu);
    const rect = td.getBoundingClientRect();
    menu.style.left = `${Math.min(rect.left, window.innerWidth - menu.offsetWidth - 6)}px`;
    menu.style.top = `${rect.bottom + 2}px`;
    const close = (e: MouseEvent): void => { if (!menu.contains(e.target as Node)) { menu.remove(); document.removeEventListener("mousedown", close); } };
    setTimeout(() => document.addEventListener("mousedown", close), 0);
  }

  // Toggle the autofilter over the current selection (or the used range), or clear it.
  // Excel's "current region": the contiguous block of non-empty cells containing the active cell,
  // bounded by empty rows/columns and the used range. Used as the filter range when the user has
  // not explicitly selected one, so a filter is scoped to the data table, not the whole sheet.
  const currentRegion = (sheet: Sheet, r0: number, c0: number): { r1: number; c1: number; r2: number; c2: number } => {
    const has = (r: number, c: number): boolean => { const v = getCell(sheet, r, c); return !!v && v.value !== "" && v.kind !== "blank"; };
    const rowHas = (r: number, a: number, b: number): boolean => { for (let c = a; c <= b; c++) if (has(r, c)) return true; return false; };
    const colHas = (c: number, a: number, b: number): boolean => { for (let r = a; r <= b; r++) if (has(r, c)) return true; return false; };
    let r1 = r0, c1 = c0, r2 = r0, c2 = c0;
    const maxR = Math.max(1, sheet.maxRow), maxC = Math.max(1, sheet.maxCol);
    let grew = true;
    while (grew) {
      grew = false;
      if (r1 > 1 && rowHas(r1 - 1, c1, c2)) { r1--; grew = true; }
      if (r2 < maxR && rowHas(r2 + 1, c1, c2)) { r2++; grew = true; }
      if (c1 > 1 && colHas(c1 - 1, r1, r2)) { c1--; grew = true; }
      if (c2 < maxC && colHas(c2 + 1, r1, r2)) { c2++; grew = true; }
    }
    return { r1, c1, r2, c2 };
  };

  const toggleAutoFilter = (): void => {
    const sheet = wb.sheets[active]!;
    if (wb.kind !== "xlsx" && wb.kind !== "ods") return;
    if (sheet.autoFilter) {
      if (wb.kind === "xlsx") { setXlsxAutoFilter(sheet, null); if (sheet.filterHidden) for (const r of sheet.filterHidden) setXlsxRowHidden(sheet, r, false); }
      else setOdsAutoFilter(wb, sheet, null);
      sheet.autoFilter = undefined; sheet.filters = undefined; sheet.filterHidden = undefined;
    } else {
      const s = getSelRect();
      // Explicit multi-cell selection wins; otherwise scope to the data region around the cell.
      const range = (s.r2 > s.r1 || s.c2 > s.c1) ? s : currentRegion(sheet, s.r1, s.c1);
      sheet.autoFilter = range;
      if (wb.kind === "xlsx") setXlsxAutoFilter(sheet, `${colToLetters(range.c1)}${range.r1}:${colToLetters(range.c2)}${range.r2}`);
      else setOdsAutoFilter(wb, sheet, range);
    }
    mark();
    renderGrid();
  };


  const { openPivotMenu, openPivotDialog, closeMenu: closePivotMenu, applySlicer, applyTimeline, addTableSlicer } = setupPivotUi({
    wb, wrap,
    active: () => active,
    getSelRect,
    mark: () => mark(),
    renderGrid: () => renderGrid(),
    switchSheet: (i) => switchSheet(i),
    refreshPivotLayer: () => pivotLayer.refresh(),
    refreshSlicers: () => slicerLayer.refresh(),
    setRowHidden: (sheet, row, hidden) => { if (wb.kind === "xlsx") setXlsxRowHidden(sheet, row, hidden); else { (sheet.hiddenRows ??= new Set())[hidden ? "add" : "delete"](row); sheet.odsDirty = true; } },
    currentRegion: (sheet, r, c) => currentRegion(sheet, r, c),
    chartsOn,
    chartInsert: (rect) => chartUi.openInsert(rect),
  });
  // The sparkline hosted by the single focused cell, if any (drives the float-bar actions).
  const focusedSparkline = (): { r: number; c: number } | null => {
    const s = getSelRect();
    if (s.r1 !== s.r2 || s.c1 !== s.c2) return null;
    const sheet = wb.sheets[active]!;
    return sheet.sparklines?.some((sp) => sp.host.r === s.r1 && sp.host.c === s.c1) ? { r: s.r1, c: s.c1 } : null;
  };
  const sparkSingle = (sheet: Sheet, host: { r: number; c: number }, spec: Parameters<typeof setXlsxSparkline>[2]): void => {
    if (wb.kind === "ods") setOdsSparkline(sheet, host, spec);
    else setXlsxSparkline(sheet, host, spec);
  };
  const deleteFocusedSparkline = (): void => {
    const h = focusedSparkline();
    if (!h) return;
    sparkSingle(wb.sheets[active]!, h, null);
    mark(); renderGrid();
  };

  const { openLinkDialog, openDvDialog, openCfDialog, openSparkDialog, openShapeDialog, openNoteDialog } = setupDialogs({
    wb, wrap,
    active: () => active,
    getSelRect,
    mark: () => mark(),
    renderGrid: () => renderGrid(),
    refreshShapes: () => shapeLayer.refresh(),
    applySparkline: (sheet, host, spec) => sparkSingle(sheet, host, spec),
  });

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
    let inp = inputAt(key(r, c));
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
      inp = inputAt(key(r, c));
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
  const onPaneScroll = (pane: Pane) => () => {
    shareScroll(pane); // the panes sharing this axis follow along
    if (scrollScheduled) return;
    scrollScheduled = true;
    setTimeout(() => {
      scrollScheduled = false;
      for (const p of panes) renderPane(p); // a shared axis moved the neighbours too
    }, 16);
  };
  for (const p of [mainPane, splitPane, rightPane, rightSplitPane]) {
    p.scrollEl.addEventListener("scroll", onPaneScroll(p));
    p.scrollEl.addEventListener("pointerdown", () => { lastPane = p; }, true);
  }

  const renderGrid = () => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    for (const p of panes) { p.inputs = new Map(); p.tds = new Map(); }
    const keepTop = gridScroll.scrollTop;
    const keepLeft = gridScroll.scrollLeft;
    gridScroll.innerHTML = "";
    totalRows = Math.max(ROWS_MIN, sheet.maxRow + 6) + extraRows;
    totalCols = Math.max(COLS_MIN, sheet.maxCol + 2) + extraCols;
    renderedRows = totalRows;
    renderedCols = totalCols;
    condVisuals = sheet.condFormats?.length ? computeCondVisuals(sheet, { evaluator: makeFormulaEvaluator(wb), sheetName: sheet.name }, dateToSerial(new Date().getFullYear(), new Date().getMonth() + 1, new Date().getDate())) : new Map();
    sparkAt = new Map();
    for (const sp of sheet.sparklines ?? []) sparkAt.set(key(sp.host.r, sp.host.c), sp);
    brkRows = new Set(sheet.printSetup?.rowBreaks ?? []);
    brkCols = new Set(sheet.printSetup?.colBreaks ?? []);
    // Only an explicit print area is drawn; the implicit "whole used range" is not a boundary the
    // user set, so outlining it would be noise.
    printAreas = sheet.printSetup?.printArea ?? [];
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
    mainPane.tableEl = table;
    gridScroll.appendChild(table);
    layoutPanes(sheet);
    // Every active pane past the main one gets its own table.
    for (const p of panes) {
      if (p === mainPane) continue;
      p.scrollEl.innerHTML = "";
      const t = document.createElement("table");
      t.className = "sheetedit-table";
      p.tableEl = t;
      p.scrollEl.appendChild(t);
    }

    for (const p of panes) { p.winR1 = 1; p.winR2 = 0; p.winC1 = 1; p.winC2 = 0; }
    // A trailing band opens just past its boundary unless the user already scrolled it.
    const splitTop = splitScroll.scrollTop || yOfRow((sheet.freeze?.rows ?? 0) + 1);
    const splitLeft = rightPane.scrollEl.scrollLeft || xOfCol((sheet.freeze?.cols ?? 0) + 1);
    cur = mainPane;
    renderWindow(true, keepTop, keepLeft); // build the window for the kept position first
    gridScroll.scrollTop = keepTop; // now the spacers exist, so the browser keeps it
    gridScroll.scrollLeft = keepLeft;
    for (const p of panes) {
      if (p === mainPane || !p.tableEl) continue;
      const y = p.band.r === 1 ? splitTop : keepTop;
      const x = p.band.c === 1 ? splitLeft : keepLeft;
      cur = p;
      renderWindowInner(true, y, x);
      p.scrollEl.scrollTop = y;
      p.scrollEl.scrollLeft = x;
    }
    cur = mainPane;
    if (panes.length > 1) fitSplitSizes(); // the lines exist now, so the boundaries can be measured
    chartLayer.refresh();
    imageLayer.refresh();
    shapeLayer.refresh();
    slicerLayer.refresh();
    if (caps.formControls) controlLayer.refresh();
    timelineLayer.refresh();
    outlineLayer.refresh();
    paneDividers.refresh();
    pivotLayer.refresh();
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

  /** Workbook structure protection blocks the whole add / delete / rename / reorder set. */
  const allowSheetOps = (): boolean => allow(!isStructureLocked(wb), "protectedStructure");

  const doAddSheet = (): void => {
    if (!allowSheetOps()) return;
    try {
      const i = addSheet(wb);
      mark();
      switchSheet(i); // renders tabs + grid
      beginRenameTab(tabs.children[i] as HTMLElement, i); // let the user name it immediately
    } catch { /* unsupported file type */ }
  };
  const doDeleteSheet = (i: number): void => {
    if (wb.sheets.length <= 1) return;
    if (!allowSheetOps()) return;
    deleteSheet(wb, i);
    if (active === i) active = Math.min(i, wb.sheets.length - 1);
    else if (active > i) active -= 1;
    mark();
    renderTabs();
    renderGrid();
  };
  const doMoveSheet = (from: number, to: number): void => {
    if (to < 0 || to >= wb.sheets.length || from === to) return;
    if (!allowSheetOps()) return;
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
      if (commit && v && v !== wb.sheets[i]?.name && allowSheetOps()) { renameSheet(wb, i, v); mark(); }
      renderTabs();
    };
    input.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") { e.preventDefault(); finish(true); }
      else if (e.key === "Escape") { e.preventDefault(); finish(false); }
    });
    input.addEventListener("blur", () => finish(true));
  }
  function showTabMenu(i: number, x: number, tabEl: HTMLElement): void {
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
    add(t("sheetRename"), () => beginRenameTab(tabEl, i));
    add(t("sheetMoveLeft"), () => doMoveSheet(i, i - 1), i === 0);
    add(t("sheetMoveRight"), () => doMoveSheet(i, i + 1), i === wb.sheets.length - 1);
    add(t("sheetDelete"), () => doDeleteSheet(i), wb.sheets.length <= 1);
    // A workbook has to keep one tab the user can reach, so the last visible sheet cannot hide.
    add(t("sheetHide"), () => doSetVisibility(i, "hidden"), visibleSheetCount(wb) <= 1);
    // "Very hidden" is Excel's macro-only state, so a sheet in it is not offered here either.
    const hidden = wb.sheets.map((s, k) => ({ s, k })).filter(({ s }) => s.visibility === "hidden");
    if (hidden.length) {
      const sep = document.createElement("div");
      sep.className = "sheetedit-pop-sep";
      menu.appendChild(sep);
      for (const { s: h, k } of hidden) add(`${t("sheetUnhide")}: ${h.name}`, () => doSetVisibility(k, undefined));
    }
    wrap.appendChild(menu);
    menu.style.left = `${Math.min(x, wrap.getBoundingClientRect().width - menu.offsetWidth - 6)}px`;
    menu.style.bottom = "40px";
    const close = (e: MouseEvent): void => { if (!menu.contains(e.target as Node)) { menu.remove(); document.removeEventListener("mousedown", close); } };
    setTimeout(() => document.addEventListener("mousedown", close), 0);
  }

  /** Hide or show a sheet, as one undo step. Hiding the active sheet moves off it first. */
  const doSetVisibility = (index: number, visibility: Sheet["visibility"]): void => {
    if (!allow(!isStructureLocked(wb), "protectedStructure")) return;
    const before = wb.sheets[index]?.visibility;
    const apply = (v: Sheet["visibility"]): void => {
      setSheetVisibility(wb, index, v);
      if (v && active === index) {
        const to = wb.sheets.findIndex((s) => !s.visibility);
        if (to >= 0) switchSheet(to);
      }
      renderTabs();
    };
    try {
      recordWorkbook(() => apply(visibility), () => apply(before));
    } catch (e) {
      showNotice((e as Error).message);
      return;
    }
    mark();
  };

  const renderTabs = () => {
    tabs.innerHTML = "";
    // A hidden sheet draws no tab, so the DOM order stops matching the sheet order and the
    // keyboard walk has to move through the visible ones rather than through the indices.
    const shown = wb.sheets.map((s, i) => ({ s, i })).filter(({ s }) => !s.visibility);
    shown.forEach(({ s: sheet, i }, pos) => {
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
        b.addEventListener("contextmenu", (e) => { e.preventDefault(); showTabMenu(i, b.offsetLeft, b); });
      }
      // Left/Right (Home/End) move between sheet tabs, activating and focusing each.
      b.addEventListener("keydown", (e) => {
        const n = shown.length;
        let to = -1;
        if (e.key === "ArrowRight") to = (pos + 1) % n;
        else if (e.key === "ArrowLeft") to = (pos - 1 + n) % n;
        else if (e.key === "Home") to = 0;
        else if (e.key === "End") to = n - 1;
        if (to < 0) return;
        e.preventDefault();
        switchSheet(shown[to]!.i);
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
    const tl = tdAt(key(sel.r1, sel.c1));
    const br = tdAt(key(sel.r2, sel.c2));
    if (!tl || !br) return null;
    const a = tl.getBoundingClientRect();
    const b = br.getBoundingClientRect();
    return new DOMRect(a.left, a.top, b.right - a.left, b.bottom - a.top);
  };
  const floatBar = wb.kind === "xlsx" || wb.kind === "ods"
    ? setupFloatBar({
        wrap,
        bounds: () => gridScroll.getBoundingClientRect(),
        selRect: selRectNow,
        curStyle,
        applyStyle,
        spark: caps.sparklines ? { has: () => !!focusedSparkline(), edit: () => openSparkDialog(), remove: () => deleteFocusedSparkline() } : undefined,
      })
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
      chartLayer.teardown();
      imageLayer.teardown();
      shapeLayer.teardown();
      slicerLayer.teardown();
      controlLayer.teardown();
      timelineLayer.teardown();
      outlineLayer.teardown();
      paneDividers.teardown();
      pivotLayer.teardown();
      closePivotMenu();
      chartUi.teardown();
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
