import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet, SheetControl } from "../model";
import type { ChartGeom } from "./chart-overlay";

// Form controls drawn over the grid and made to work. The point of a control is its linked cell:
// a checkbox writes TRUE/FALSE there, a dropdown the 1-based index of the chosen item, a spinner
// its number. Formulas read that cell, so these are inputs to the sheet, not decoration.
//
// A control can also have a macro assigned, which is what a button is for. That runs through the
// same path the macro viewer uses, after the linked cell is written, as it does in Excel. A button
// with no macro assigned stays inert with a tooltip saying so, as do labels and group boxes.

export interface ControlLayerDeps {
  wrap: HTMLElement;
  panes: () => { el: HTMLElement; header: boolean; rowHeader: boolean }[];
  getSheet: () => Sheet | undefined;
  geom: () => ChartGeom;
  /** The items a list control offers, resolved from its source range. */
  itemsFor: (control: SheetControl) => string[];
  /**
   * The state changed: the models are already updated and dirty set. The first is the control the
   * user touched; any others are the radios its group had to clear, which belong in the same undo
   * step rather than one of their own.
   */
  onChange: (controls: SheetControl[]) => void;
  /** After a drag: the control's anchor was updated and dirty set; the host persists it. */
  onPlace?: (control: SheetControl) => void;
  /** True when this sheet's controls can be written back, which gates the drag handles. */
  editable?: () => boolean;
  /** Text for the tooltip on a button with no macro to run. */
  inertTitle: string;
  /** Run a control's assigned macro. Absent when the workbook carries no macros at all. */
  runMacro?: (name: string) => void;
  /** Tooltip prefix for a control that does have a macro. */
  macroTitle: string;
}

/** The rectangle a control occupies, in cells. Absent for a control the file never placed. */
const rectOf = (c: SheetControl): { r1: number; c1: number; r2: number; c2: number } | undefined => {
  const a = c.anchor;
  return a ? { r1: a.fromRow, c1: a.fromCol, r2: a.toRow, c2: a.toCol } : undefined;
};

/** Whether `inner`'s top-left sits inside `outer`, which is how Excel decides group membership. */
function inside(outer: SheetControl, inner: SheetControl): boolean {
  const o = rectOf(outer), i = rectOf(inner);
  if (!o || !i) return false;
  return i.r1 >= o.r1 && i.r1 <= o.r2 && i.c1 >= o.c1 && i.c1 <= o.c2;
}

/**
 * The radios a newly checked one has to turn off: the other checked radios of its group. A group is
 * the group box drawn around them, and every radio outside all of them shares one group, which is
 * how Excel decides it too.
 */
export function radioPeersToClear(controls: SheetControl[], ctl: SheetControl): SheetControl[] {
  const groupOf = (c: SheetControl): SheetControl | undefined =>
    controls.find((g) => g.kind === "groupBox" && inside(g, c));
  const box = groupOf(ctl);
  return controls.filter((other) => other !== ctl && other.kind === "radio" && other.checked && groupOf(other) === box);
}

export function setupControlLayer(deps: ControlLayerDeps): { refresh(): void; teardown(): void } {
  /** Turn off every other radio in the control's group, and report which ones changed. */
  const clearGroup = (ctl: SheetControl): SheetControl[] => {
    const cleared = radioPeersToClear(deps.getSheet()?.controls ?? [], ctl);
    for (const other of cleared) other.checked = false;
    return cleared;
  };

  const hosts = setupOverlayHosts({
    wrap: deps.wrap,
    panes: deps.panes,
    geom: () => { const g = deps.geom(); return { rnW: g.rnW, headerH: g.headerH, yOfRow: g.yOfRow }; },
    className: "sheetedit-ctrllayer",
    innerClassName: "sheetedit-ctrllayer-inner",
  });

  /** The control's body, by kind. Returns null for a kind with nothing to draw. */
  const build = (ctl: SheetControl): HTMLElement => {
    // Excel runs a control's macro after its linked cell is written, so a macro that reads that
    // cell sees the new state. Same order here.
    const fire = (): void => { if (ctl.macro && deps.runMacro) deps.runMacro(ctl.macro); };
    const commit = (also: SheetControl[] = []): void => {
      ctl.dirty = true;
      for (const o of also) o.dirty = true;
      deps.onChange([ctl, ...also]);
      fire();
    };
    switch (ctl.kind) {
      case "checkbox":
      case "radio": {
        const label = document.createElement("label");
        label.className = "sheetedit-ctrl-check";
        const input = document.createElement("input");
        input.type = ctl.kind === "radio" ? "radio" : "checkbox";
        input.checked = !!ctl.checked;
        input.addEventListener("change", () => {
          ctl.checked = input.checked;
          // One radio on means the rest of its group off, which is the only thing that makes a
          // radio a radio. Its group is the group box drawn around it, or the sheet when none is.
          const cleared = ctl.kind === "radio" && input.checked ? clearGroup(ctl) : [];
          commit(cleared);
        });
        const span = document.createElement("span");
        span.textContent = ctl.label ?? ctl.name;
        label.append(input, span);
        return label;
      }
      case "dropdown":
      case "list": {
        const select = document.createElement("select");
        select.className = "sheetedit-ctrl-select";
        // A list control's linked cell holds the 1-based index, so index 0 means "nothing chosen".
        if (ctl.kind === "dropdown") {
          const blank = document.createElement("option");
          blank.value = "0";
          blank.textContent = "";
          select.appendChild(blank);
        }
        deps.itemsFor(ctl).forEach((text, i) => {
          const opt = document.createElement("option");
          opt.value = String(i + 1);
          opt.textContent = text;
          select.appendChild(opt);
        });
        if (ctl.kind === "list") select.size = Math.max(2, Math.min(8, select.options.length));
        select.value = String(ctl.selected ?? 0);
        select.addEventListener("change", () => { ctl.selected = Number(select.value) || 0; commit(); });
        return select;
      }
      case "spin":
      case "scroll": {
        const input = document.createElement("input");
        input.type = ctl.kind === "scroll" ? "range" : "number";
        input.className = "sheetedit-ctrl-num";
        input.min = String(ctl.min ?? 0);
        input.max = String(ctl.max ?? 100);
        input.step = String(ctl.inc || 1);
        input.value = String(ctl.value ?? ctl.min ?? 0);
        input.addEventListener("change", () => { ctl.value = Number(input.value) || 0; commit(); });
        return input;
      }
      case "button": {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-ctrl-button";
        b.textContent = ctl.label ?? ctl.name;
        // A button's only job is running its macro. With none assigned there is nothing to do, so
        // say so rather than presenting a control that looks live and silently does nothing.
        if (ctl.macro && deps.runMacro) {
          b.title = `${deps.macroTitle} ${ctl.macro}`;
          b.addEventListener("click", fire);
        } else {
          b.disabled = true;
          b.title = deps.inertTitle;
        }
        return b;
      }
      default: {
        const span = document.createElement("span");
        span.className = ctl.kind === "groupBox" ? "sheetedit-ctrl-group" : "sheetedit-ctrl-label";
        span.textContent = ctl.label ?? ctl.name;
        return span;
      }
    }
  };

  /**
   * Move / resize a control by dragging. The body keeps its own clicks, so the grip strip down the
   * left edge is what moves it and the corner handle is what resizes it; a control whose whole
   * face were draggable could not be ticked.
   */
  const attachDrag = (box: HTMLElement, grip: HTMLElement, handle: HTMLElement, ctl: SheetControl): void => {
    const start = (e: PointerEvent, mode: "move" | "resize"): void => {
      e.preventDefault();
      e.stopPropagation();
      const sx = e.clientX, sy = e.clientY;
      const x0 = parseFloat(box.style.left) || 0, y0 = parseFloat(box.style.top) || 0;
      const w0 = box.offsetWidth, h0 = box.offsetHeight;
      const onMove = (ev: PointerEvent): void => {
        const dx = ev.clientX - sx, dy = ev.clientY - sy;
        if (mode === "move") { box.style.left = `${Math.max(0, x0 + dx)}px`; box.style.top = `${Math.max(0, y0 + dy)}px`; }
        else { box.style.width = `${Math.max(24, w0 + dx)}px`; box.style.height = `${Math.max(16, h0 + dy)}px`; }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        const x = parseFloat(box.style.left) || 0, y = parseFloat(box.style.top) || 0;
        if (x === x0 && y === y0 && box.offsetWidth === w0 && box.offsetHeight === h0) return; // a click, not a drag
        const g = deps.geom();
        const at = (px: number, cell: (p: number) => number, of: (i: number) => number): [number, number] => {
          const i = Math.max(1, cell(px));
          return [i, Math.max(0, Math.round(px - of(i)))];
        };
        const [fc, fco] = at(x, g.colAt, g.xOfCol);
        const [fr, fro] = at(y, g.rowAt, g.yOfRow);
        const [tc, tco] = at(x + box.offsetWidth, g.colAt, g.xOfCol);
        const [tr, tro] = at(y + box.offsetHeight, g.rowAt, g.yOfRow);
        ctl.anchor = { fromCol: fc, fromRow: fr, fromColOff: fco, fromRowOff: fro, toCol: tc, toRow: tr, toColOff: tco, toRowOff: tro };
        ctl.dirty = true;
        deps.onPlace?.(ctl);
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    grip.addEventListener("pointerdown", (e) => start(e, "move"));
    handle.addEventListener("pointerdown", (e) => start(e, "resize"));
  };

  const refresh = (): void => {
    hosts.clear();
    const sheet = deps.getSheet();
    const controls = sheet?.controls ?? [];
    hosts.setVisible(controls.length > 0);
    if (!controls.length) return;
    const g = deps.geom();
    for (const ctl of controls) {
      const a = ctl.anchor;
      const x = a ? g.xOfCol(a.fromCol) + a.fromColOff : 20;
      const y = a ? g.yOfRow(a.fromRow) + a.fromRowOff : 20;
      const w = a ? Math.max(24, g.xOfCol(a.toCol) + a.toColOff - x) : 120;
      const h = a ? Math.max(18, g.yOfRow(a.toRow) + a.toRowOff - y) : 22;
      const box = document.createElement("div");
      box.className = `sheetedit-ctrlbox kind-${ctl.kind}`;
      box.dataset.control = ctl.name;
      Object.assign(box.style, { left: `${x}px`, top: `${y}px`, width: `${w}px`, height: `${h}px` });
      if (ctl.linkedCell) box.title = `${ctl.label ?? ctl.name} -> ${ctl.linkedCell.replace(/\$/g, "")}`;
      box.appendChild(build(ctl));
      if (deps.editable?.()) {
        box.classList.add("editable");
        const grip = document.createElement("div");
        grip.className = "sheetedit-ctrl-grip";
        const handle = document.createElement("div");
        handle.className = "sheetedit-ctrl-resize";
        box.append(grip, handle);
        attachDrag(box, grip, handle, ctl);
      }
      hosts.hostFor(a?.fromRow ?? 1, a?.fromCol).appendChild(box);
    }
    hosts.layout();
  };

  return { refresh, teardown: hosts.teardown };
}
