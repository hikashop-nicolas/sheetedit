import { setupOverlayHosts } from "./overlay-hosts";
import type { Sheet, SheetControl } from "../model";
import type { ChartGeom } from "./chart-overlay";

// Form controls drawn over the grid and made to work. The point of a control is its linked cell:
// a checkbox writes TRUE/FALSE there, a dropdown the 1-based index of the chosen item, a spinner
// its number. Formulas read that cell, so these are inputs to the sheet, not decoration.
//
// A button runs a macro, which a browser cannot and should not execute, so it is drawn and left
// inert with a tooltip saying why. Same for labels and group boxes, which never did anything.

export interface ControlLayerDeps {
  wrap: HTMLElement;
  panes: () => { el: HTMLElement; header: boolean; rowHeader: boolean }[];
  getSheet: () => Sheet | undefined;
  geom: () => ChartGeom;
  /** The items a list control offers, resolved from its source range. */
  itemsFor: (control: SheetControl) => string[];
  /** The state changed: the model is already updated and dirty set. */
  onChange: (control: SheetControl) => void;
  /** Text for the inert-control tooltip (macros cannot run in a browser). */
  inertTitle: string;
}

export function setupControlLayer(deps: ControlLayerDeps): { refresh(): void; teardown(): void } {
  const hosts = setupOverlayHosts({
    wrap: deps.wrap,
    panes: deps.panes,
    geom: () => { const g = deps.geom(); return { rnW: g.rnW, headerH: g.headerH, yOfRow: g.yOfRow }; },
    className: "sheetedit-ctrllayer",
    innerClassName: "sheetedit-ctrllayer-inner",
  });

  /** The control's body, by kind. Returns null for a kind with nothing to draw. */
  const build = (ctl: SheetControl): HTMLElement => {
    const commit = (): void => { ctl.dirty = true; deps.onChange(ctl); };
    switch (ctl.kind) {
      case "checkbox":
      case "radio": {
        const label = document.createElement("label");
        label.className = "sheetedit-ctrl-check";
        const input = document.createElement("input");
        input.type = ctl.kind === "radio" ? "radio" : "checkbox";
        input.checked = !!ctl.checked;
        input.addEventListener("change", () => { ctl.checked = input.checked; commit(); });
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
        // A button's only job is running a macro, which this cannot do. Say so rather than
        // presenting a control that looks live and silently does nothing.
        b.disabled = true;
        b.title = deps.inertTitle;
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
      hosts.hostFor(a?.fromRow ?? 1, a?.fromCol).appendChild(box);
    }
    hosts.layout();
  };

  return { refresh, teardown: hosts.teardown };
}
