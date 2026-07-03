import { t } from "./i18n";
import { createFormulaBar } from "./formulabar";
import { buildToolbar } from "./toolbar";
import { setupFloatBar } from "./floatbar";
import { UndoHistory, applyFields, snapFields, type UndoCellChange } from "./history";
import type { Cell, Sheet, StyleChange, Workbook } from "./model";
import { cellDisplay, colToLetters, ensureCell, getCell, key } from "./model";
import { setOdsCellStyle, setOdsColWidth, setOdsMerge, setOdsRowHeight } from "./ods";
import { recalc } from "./recalc";
import { readWorkbook, setCellInput, writeWorkbook } from "./workbook";
import { setXlsxCellStyle, setXlsxColWidth, setXlsxMerge, setXlsxRowHeight } from "./xlsx";
// ---------------------------------------------------------------------------
// Editor
// ---------------------------------------------------------------------------

export interface SheetEditorOptions {
  onChange?: () => void;
}
export interface SheetEditor {
  getBytes(): Promise<Uint8Array>;
  isDirty(): boolean;
  destroy(): void;
}

export const ROWS_MIN = 24;
export const COLS_MIN = 12;
export const ROWS_CAP = 5000;
export const COLS_CAP = 256;
export const ROW_CHUNK = 20; // rows added per "+ Row" click
export const COL_CHUNK = 6; // columns added per "+ Col" click
export const STYLE_ID = "sheetedit-style";

export function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const s = document.createElement("style");
  s.id = STYLE_ID;
  s.textContent = `
    .sheetedit-wrap { display:flex; flex-direction:column; height:100%; background:#1f2227; color:#e6e6e6; font:13px system-ui, sans-serif; }
    .sheetedit-toolbar { display:flex; flex-wrap:nowrap; overflow:hidden; align-items:center; gap:5px; padding:5px 8px; background:#2b2f36; border-bottom:1px solid #1c1f24; }
    .sheetedit-btn {
      font:inherit; font-size:13px; background:#3a3f47; color:#e6e6e6; border:1px solid #4a4f57;
      border-radius:6px; padding:4px 9px; cursor:pointer; min-width:32px; line-height:1.1;
    }
    .sheetedit-btn:hover { background:#454b54; }
    .sheetedit-btn:focus-visible { outline:2px solid #6e7bff; outline-offset:1px; }
    .sheetedit-tb-sep { width:1px; align-self:stretch; background:#4a4f57; margin:1px 3px; }
    .sheetedit-color { width:30px; height:28px; padding:0; border:1px solid #4a4f57; border-radius:6px; background:#3a3f47; cursor:pointer; }
    .sheetedit-btn svg { display:block; width:16px; height:16px; }
    .sheetedit-table th.colhead, .sheetedit-table th.rownum, .sheetedit-table th.corner { cursor:pointer; }
    .sheetedit-table th.colhead:hover, .sheetedit-table th.rownum:hover, .sheetedit-table th.corner:hover { background:#e3e3e8; }
    .sheetedit-table td.sheetedit-sel input { background:rgba(110,123,255,0.18); }
    .sheetedit-pop { position:fixed; z-index:30; background:#2b2f36; border:1px solid #4a4f57; border-radius:8px; padding:4px; box-shadow:0 6px 18px rgba(0,0,0,0.45); display:flex; flex-direction:column; min-width:130px; }
    .sheetedit-pop-item { font:inherit; font-size:13px; text-align:left; background:transparent; color:#e6e6e6; border:0; border-radius:5px; padding:7px 11px; cursor:pointer; }
    .sheetedit-pop-item:hover { background:#3a3f47; }
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
      position:sticky; top:0; z-index:2; background:#f1f1f4; color:#555; font-weight:600;
      padding:3px 8px; text-align:center; user-select:none;
    }
    .sheetedit-table th.corner { left:0; z-index:3; }
    .sheetedit-table th.rownum { position:sticky; left:0; z-index:1; top:auto; text-align:right; background:#f1f1f4; }
    /* Resize grips: a thin strip on the header border, wide enough to grab on touch. */
    .sheetedit-colgrip { position:absolute; top:0; right:-4px; width:9px; height:100%; cursor:col-resize; z-index:4; touch-action:none; }
    .sheetedit-rowgrip { position:absolute; left:0; bottom:-4px; width:100%; height:9px; cursor:row-resize; z-index:4; touch-action:none; }
    .sheetedit-colgrip:hover { box-shadow:inset -2px 0 0 0 #6e7bff; }
    .sheetedit-rowgrip:hover { box-shadow:inset 0 -2px 0 0 #6e7bff; }
    .sheetedit-table input {
      border:0; background:transparent; color:#1a1a1a; font:inherit; padding:3px 8px;
      width:100%; box-sizing:border-box; outline:none;
    }
    .sheetedit-table td.num input { text-align:right; font-variant-numeric:tabular-nums; }
    .sheetedit-table input:focus { box-shadow:inset 0 0 0 2px #6e7bff; background:#eef0ff; }
    .sheetedit-tb-slot { display:inline-flex; align-items:center; gap:5px; }
    .sheetedit-tb-groupmenu { position:absolute; z-index:30; display:flex; align-items:center; gap:5px; padding:6px 8px; background:#2b2f36; border:1px solid #1c1f24; border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
    .sheetedit-fxbar { display:flex; align-items:center; gap:6px; padding:4px 8px; background:#23262c; border-bottom:1px solid #1c1f24; position:relative; }
    .sheetedit-fxref { min-width:52px; text-align:center; font:12px/1.6 ui-monospace,monospace; color:#aab2bf; background:#2b2f36; border-radius:5px; padding:2px 6px; }
    .sheetedit-fxbtns { position:relative; display:inline-flex; gap:2px; }
    .sheetedit-fxsum { font-weight:700; }
    .sheetedit-fxmenu { position:absolute; top:100%; left:0; z-index:30; display:flex; flex-direction:column; background:#2b2f36; border:1px solid #1c1f24; border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); padding:4px; }
    .sheetedit-fxmenu-item { background:none; border:0; color:#e7eaf0; text-align:left; padding:5px 12px; font:12px system-ui,sans-serif; cursor:pointer; border-radius:5px; }
    .sheetedit-fxmenu-item:hover { background:#3a4047; }
    .sheetedit-fxinput { flex:1; min-width:60px; background:#1c1f24; border:1px solid #3a4047; border-radius:5px; color:#e7eaf0; font:13px ui-monospace,monospace; padding:4px 8px; }
    .sheetedit-fxbar.is-picking .sheetedit-fxinput { border-color:#4f8ef7; }
    .sheetedit-fxmenu[hidden], .sheetedit-tb-groupmenu[hidden] { display:none; }
    .sheetedit-floatbar { position:fixed; z-index:40; display:flex; align-items:center; gap:2px; padding:4px 6px; background:#2b2f36; border:1px solid #1c1f24; border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
    .sheetedit-floatbar[hidden] { display:none; }
    .sheetedit-error { background:#7a2b2b; color:#ffd7d7; padding:10px 14px; font:13px/1.5 system-ui,sans-serif; }
    .sheetedit-tabs { display:flex; align-items:center; gap:2px; padding:5px 8px; background:#2b2f36; border-top:1px solid #1c1f24; overflow-x:auto; }
    .sheetedit-tab {
      font:inherit; background:#3a3f47; color:#cfd3da; border:1px solid #4a4f57; border-bottom:none;
      border-radius:5px 5px 0 0; padding:4px 12px; cursor:pointer; white-space:nowrap;
    }
    .sheetedit-tab[aria-selected="true"] { background:#6e7bff; color:#fff; border-color:#6e7bff; }
    .sheetedit-tab:focus-visible { outline:2px solid #fff; outline-offset:1px; }
  `;
  document.head.appendChild(s);
}

export function createSheetEditor(
  container: HTMLElement,
  bytes: Uint8Array,
  options: SheetEditorOptions = {},
): SheetEditor {
  const original = bytes.slice();
  let dirty = false;
  injectStyles();

  let wb: Workbook;
  try {
    wb = readWorkbook(bytes);
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
      isDirty: () => false,
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
    return live ? (live.formula != null ? "=" + live.formula : live.value) : "";
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
  });
  fxbar.input.addEventListener("pointerdown", () => {
    barGrab = true;
  });
  fxbar.input.addEventListener("focus", () => {
    setTimeout(() => (barGrab = false), 0);
  });
  wrap.append(toolbar, fxbar.el, gridScroll, tabs);
  container.appendChild(wrap);

  let active = 0;
  let inputs = new Map<string, HTMLInputElement>();
  let tds = new Map<string, HTMLElement>();
  // Extra rows/columns the user added beyond the sheet's used extent (per active sheet).
  let extraRows = 0;
  let extraCols = 0;
  // Selection rectangle (1-based, inclusive) and the anchor for shift-extend.
  let sel: { r1: number; c1: number; r2: number; c2: number } | null = null;
  let anchor: { r: number; c: number } | null = null;

  const paintSel = () => {
    for (const td of tds.values()) td.classList.remove("sheetedit-sel");
    if (!sel) return;
    for (let r = sel.r1; r <= sel.r2; r++)
      for (let c = sel.c1; c <= sel.c2; c++) tds.get(key(r, c))?.classList.add("sheetedit-sel");
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

  const applyStyle = (change: StyleChange) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const positions: { r: number; c: number }[] = [];
    let n = 0;
    for (let r = sel.r1; r <= sel.r2 && n < 4000; r++)
      for (let c = sel.c1; c <= sel.c2 && n < 4000; c++, n++) positions.push({ r, c });
    recordCells(positions, () => {
      for (const pos of positions) setCellStyle(sheet, ensureCell(sheet, pos.r, pos.c), change);
    });
    mark();
    renderGrid();
  };

  // Apply a border mode across the selection, computing per-cell sides from each cell's
  // position in the rectangle (e.g. "outer" only borders the perimeter).
  type BorderMode = "all" | "outer" | "top" | "bottom" | "left" | "right" | "none";
  const applyBorder = (mode: BorderMode) => {
    if ((wb.kind !== "xlsx" && wb.kind !== "ods") || !sel) return;
    const sheet = wb.sheets[active];
    if (!sheet) return;
    const { r1, c1, r2, c2 } = sel;
    const positions: { r: number; c: number }[] = [];
    let n = 0;
    for (let r = r1; r <= r2 && n < 4000; r++) for (let c = c1; c <= c2 && n < 4000; c++, n++) positions.push({ r, c });
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
    renderGrid();
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
    const close = () => {
      pop.remove();
      borderPop = null;
      document.removeEventListener("pointerdown", onOutside, true);
    };
    const onOutside = (e: Event) => {
      const tgt = e.target as Node;
      if (!pop.contains(tgt) && !btn.contains(tgt)) close();
    };
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
      b.textContent = label;
      b.addEventListener("click", () => {
        applyBorder(mode);
        close();
      });
      pop.appendChild(b);
    }
    document.body.appendChild(pop);
    borderPop = pop;
    const r = btn.getBoundingClientRect();
    pop.style.left = `${Math.round(r.left)}px`;
    pop.style.top = `${Math.round(r.bottom + 4)}px`;
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

  const toolbarHandle = buildToolbar({
    toolbar,
    wrap,
    styled: wb.kind === "xlsx" || wb.kind === "ods",
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
    applyStyle,
    curStyle,
    openBorderPopover,
    toggleMerge,
  });

  const mark = () => {
    if (!dirty) {
      dirty = true;
    }
    options.onChange?.();
  };

  const displayValue = (sheet: Sheet, r: number, c: number): string => cellDisplay(getCell(sheet, r, c));

  const refreshDisplays = (sheet: Sheet, except?: HTMLInputElement) => {
    for (const [k, input] of inputs) {
      if (input === except) continue;
      const [r, c] = k.split(":").map(Number);
      input.value = displayValue(sheet, r!, c!);
      const cell = getCell(sheet, r!, c!);
      input.parentElement?.classList.toggle("num", cell?.kind === "n");
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

  // Commit a raw value into a cell and refresh (shared by the grid and the formula bar).
  const commitValue = (r: number, c: number, raw: string) => {
    const sheet = wb.sheets[active]!;
    recordCells([{ r, c }], () => setCellInput(sheet, r, c, raw));
    recalc(wb);
    mark();
    refreshDisplays(sheet);
  };

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
      inputs.get(key(tr, tc))?.focus();
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
    const positions: { r: number; c: number }[] = [];
    for (let r = range.r1; r <= range.r2; r++) for (let c = range.c1; c <= range.c2; c++) positions.push({ r, c });
    recordCells(positions, () => {
      for (const pos of positions) setCellInput(sheet, pos.r, pos.c, "");
    });
    recalc(wb);
    mark();
    refreshDisplays(sheet);
  };
  // Copy a multi-cell selection as TSV (a single cell keeps the native copy).
  const copyRange = (e: ClipboardEvent) => {
    if (!sel || (sel.r1 === sel.r2 && sel.c1 === sel.c2)) return;
    const sheet = wb.sheets[active]!;
    const lines: string[] = [];
    for (let r = sel.r1; r <= sel.r2; r++) {
      const vals: string[] = [];
      for (let c = sel.c1; c <= sel.c2; c++) vals.push(getCell(sheet, r, c)?.value ?? "");
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
    inputs.get(key(r0, c0))?.focus();
  };

  const renderGrid = () => {
    const sheet = wb.sheets[active];
    if (!sheet) return;
    inputs = new Map();
    tds = new Map();
    gridScroll.innerHTML = "";
    const rows = Math.min(ROWS_CAP, Math.max(ROWS_MIN, sheet.maxRow + 6) + extraRows);
    const cols = Math.min(COLS_CAP, Math.max(COLS_MIN, sheet.maxCol + 2) + extraCols);

    const table = document.createElement("table");
    table.className = "sheetedit-table";

    // Column widths (table-layout is fixed, so these are authoritative). The table is
    // sized to the sum so columns keep their width and the grid scrolls horizontally,
    // rather than the table shrinking to the viewport and squashing every column.
    const colgroup = document.createElement("colgroup");
    const rnCol = document.createElement("col");
    rnCol.style.width = "44px";
    colgroup.appendChild(rnCol);
    let totalW = 44;
    const colEls: HTMLElement[] = [];
    for (let c = 1; c <= cols; c++) {
      const w = sheet.colWidths?.get(c) ?? 96;
      const col = document.createElement("col");
      col.style.width = `${w}px`;
      colgroup.appendChild(col);
      colEls.push(col);
      totalW += w;
    }
    table.appendChild(colgroup);
    table.style.width = `${totalW}px`;

    const head = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "corner";
    corner.title = t("selectAll");
    corner.addEventListener("click", () => setSel(1, 1, rows, cols));
    head.appendChild(corner);
    for (let c = 1; c <= cols; c++) {
      const th = document.createElement("th");
      th.className = "colhead";
      th.textContent = colToLetters(c);
      th.title = t("selectColumn", { col: colToLetters(c) });
      th.addEventListener("click", () => {
        if (resizing) return;
        anchor = { r: 1, c };
        setSel(1, c, rows, c);
      });
      const grip = document.createElement("div");
      grip.className = "sheetedit-colgrip";
      const colEl = colEls[c - 1]!;
      grip.addEventListener("pointerdown", (e) =>
        startColResize(e, c, colEl, sheet.colWidths?.get(c) ?? 96),
      );
      th.appendChild(grip);
      head.appendChild(th);
    }
    table.appendChild(head);

    // Merged ranges: the top-left cell spans; covered cells are not rendered.
    const covered = new Set<string>();
    const spanAt = new Map<string, { rs: number; cs: number }>();
    for (const m of sheet.merges ?? []) {
      spanAt.set(key(m.r1, m.c1), { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
      for (let r = m.r1; r <= m.r2; r++)
        for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(key(r, c));
    }

    for (let r = 1; r <= rows; r++) {
      const tr = document.createElement("tr");
      const rh = sheet.rowHeights?.get(r);
      if (rh) tr.style.height = `${rh}px`;
      const rn = document.createElement("th");
      rn.className = "rownum";
      rn.textContent = String(r);
      rn.title = t("selectRow", { row: r });
      rn.addEventListener("click", () => {
        if (resizing) return;
        anchor = { r, c: 1 };
        setSel(r, 1, r, cols);
      });
      const rgrip = document.createElement("div");
      rgrip.className = "sheetedit-rowgrip";
      rgrip.addEventListener("pointerdown", (e) => startRowResize(e, r, tr, rh ?? 22));
      rn.appendChild(rgrip);
      tr.appendChild(rn);
      for (let c = 1; c <= cols; c++) {
        if (covered.has(key(r, c))) continue; // part of a merge; the top-left cell spans it
        const td = document.createElement("td");
        td.dataset.rc = key(r, c);
        tds.set(key(r, c), td);
        const sp = spanAt.get(key(r, c));
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
        // Apply the file's visual style (fill/borders on the cell, font/colour/align on the text).
        const cs = cell?.cellStyle;
        if (cs) {
          if (cs.bg) td.style.background = cs.bg;
          if (cs.borders) {
            // Override the default gridline box-shadow: keep light right/bottom unless the
            // file specifies a border there, and add the file's top/left where present.
            const bd = cs.borders;
            const g = "#e3e3e6";
            const sh = [`inset -1px 0 0 0 ${bd.right ?? g}`, `inset 0 -1px 0 0 ${bd.bottom ?? g}`];
            if (bd.top) sh.push(`inset 0 1px 0 0 ${bd.top}`);
            if (bd.left) sh.push(`inset 1px 0 0 0 ${bd.left}`);
            td.style.boxShadow = sh.join(", ");
          }
          if (cs.bold) input.style.fontWeight = "700";
          if (cs.italic) input.style.fontStyle = "italic";
          if (cs.color) input.style.color = cs.color;
          if (cs.align) input.style.textAlign = cs.align;
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
          const live = getCell(sheet, r, c);
          const before = live ? (live.formula != null ? "=" + live.formula : live.value) : "";
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
            const target = inputs.get(key(rr, cc));
            if (!target) return;
            e.preventDefault();
            input.blur();
            target.focus();
          };
          if (e.key === "Enter") {
            e.preventDefault();
            input.blur();
            inputs.get(key(r + 1, c))?.focus();
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
        input.addEventListener("copy", copyRange);
        input.addEventListener("paste", (e) => pasteTsv(e, r, c));
        td.appendChild(input);
        tr.appendChild(td);
        inputs.set(ki, input);
      }
      table.appendChild(tr);
    }
    gridScroll.appendChild(table);
    paintSel(); // restore the selection highlight after a re-render
  };

  const renderTabs = () => {
    tabs.innerHTML = "";
    wb.sheets.forEach((sheet, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sheetedit-tab";
      b.textContent = sheet.name;
      b.setAttribute("role", "tab");
      b.setAttribute("aria-selected", String(i === active));
      b.addEventListener("click", () => {
        if (i === active) return;
        active = i;
        extraRows = 0; // each sheet starts at its own extent
        extraCols = 0;
        sel = null;
        anchor = null;
        renderTabs();
        renderGrid();
      });
      tabs.appendChild(b);
    });
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
    async getBytes() {
      return dirty ? writeWorkbook(wb) : original.slice();
    },
    destroy() {
      toolbarHandle.teardown();
      floatBar.teardown();
      wrap.remove();
    },
  };
}
