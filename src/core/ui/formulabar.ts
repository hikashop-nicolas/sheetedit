import { t } from "../i18n";

// The formula bar above the grid: the active cell's reference, the function buttons
// (Σ plus a small menu), and a wide input mirroring the cell being edited, so a
// formula stays readable even when its cell is narrow. The editor owns all state;
// this module only builds the DOM and forwards events.

export const FX_FUNCTIONS = ["SUM", "AVERAGE", "MIN", "MAX", "COUNT"] as const;

export interface FormulaBar {
  el: HTMLElement;
  input: HTMLInputElement;
  setRef(name: string): void;
  setValue(v: string): void;
  /** Show a hint (e.g. "select a range") in place of the value; null restores. */
  setHint(msg: string | null): void;
}

const SPARKLE = `<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 2.5l1.1 2.9L12 6.5 9.1 7.6 8 10.5 6.9 7.6 4 6.5l2.9-1.1z"/><path d="M12.5 11l.5 1.3 1.3.5-1.3.5-.5 1.3-.5-1.3-1.3-.5 1.3-.5z"/></svg>`;

export function createFormulaBar(opts: {
  onInput(value: string): void;
  onEnter(value: string): void;
  onEscape(): void;
  onFn(fn: string): void;
  /** When provided, an on-device "help me write a formula" button is shown; it calls this
      with its own element so the caller can anchor a popover under it. */
  onAssist?(anchor: HTMLElement): void;
}): FormulaBar {
  const el = document.createElement("div");
  el.className = "sheetedit-fxbar";

  const ref = document.createElement("span");
  ref.className = "sheetedit-fxref";
  ref.title = t("activeCell");

  // Σ applies SUM directly; the caret opens the function menu. mousedown is
  // prevented on both so clicking them never blurs (and commits) a pending edit.
  const sum = document.createElement("button");
  sum.type = "button";
  sum.className = "sheetedit-btn sheetedit-fxsum";
  sum.textContent = "Σ";
  sum.title = t("insertSum");
  sum.setAttribute("aria-label", t("insertSum"));
  sum.addEventListener("mousedown", (e) => e.preventDefault());
  sum.addEventListener("click", () => opts.onFn("SUM"));

  const more = document.createElement("button");
  more.type = "button";
  more.className = "sheetedit-btn sheetedit-fxmore";
  more.textContent = "▾";
  more.title = t("functions");
  more.setAttribute("aria-label", t("functions"));
  more.setAttribute("aria-haspopup", "menu");
  const menu = document.createElement("div");
  menu.className = "sheetedit-fxmenu";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  for (const fn of FX_FUNCTIONS) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sheetedit-fxmenu-item";
    b.setAttribute("role", "menuitem");
    b.textContent = fn;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", () => {
      menu.hidden = true;
      opts.onFn(fn);
    });
    menu.appendChild(b);
  }
  more.addEventListener("mousedown", (e) => e.preventDefault());
  more.addEventListener("click", (e) => {
    e.stopPropagation();
    menu.hidden = !menu.hidden;
  });
  document.addEventListener("click", (e) => {
    if (!menu.hidden && !menu.contains(e.target as Node) && e.target !== more) menu.hidden = true;
  });
  menu.addEventListener("keydown", (e) => {
    if (e.key === "Escape") menu.hidden = true;
  });

  const input = document.createElement("input");
  input.type = "text";
  input.className = "sheetedit-fxinput";
  input.setAttribute("aria-label", t("formulaBar"));
  input.autocomplete = "off";
  input.spellcheck = false;
  input.addEventListener("input", () => opts.onInput(input.value));
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      opts.onEnter(input.value);
    } else if (e.key === "Escape") {
      e.preventDefault();
      opts.onEscape();
    }
  });

  const fxwrap = document.createElement("span");
  fxwrap.className = "sheetedit-fxbtns";
  fxwrap.append(sum, more, menu);
  if (opts.onAssist) {
    const assist = document.createElement("button");
    assist.type = "button";
    assist.className = "sheetedit-btn sheetedit-fxassist";
    assist.innerHTML = SPARKLE;
    assist.title = t("fxAssist");
    assist.setAttribute("aria-label", t("fxAssist"));
    assist.addEventListener("mousedown", (e) => e.preventDefault());
    assist.addEventListener("click", () => opts.onAssist!(assist));
    fxwrap.append(assist);
  }
  el.append(ref, fxwrap, input);

  return {
    el,
    input,
    setRef(name) {
      ref.textContent = name;
    },
    setValue(v) {
      if (input.value !== v) input.value = v;
    },
    setHint(msg) {
      input.placeholder = msg ?? "";
      el.classList.toggle("is-picking", msg != null);
    },
  };
}
