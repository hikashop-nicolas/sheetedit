import { t } from "../i18n";

// On-device "help me write a formula": turns a plain-language request into a spreadsheet
// formula via localml, entirely in the browser. localml is imported LAZILY (dynamic import)
// the first time it runs, so the editor's base bundle never pulls in transformers.js until
// used. The result streams into an editable field; nothing changes the sheet until Insert.

export interface FormulaContext {
  cell: string; // active cell ref, e.g. "C5"
  headers: { col: string; name: string }[]; // row-1 labels, e.g. [{ col: "A", name: "Name" }]
}

export interface FormulaAssistDeps {
  wrap: HTMLElement; // the editor chrome the popover overlays (positioned, not toolbar-clipped)
  getContext(): FormulaContext;
  onAccept(formula: string): void;
}

// The system prompt keeps the small model on rails: one formula, no prose, plus the active
// cell and the column headers so it can map names to columns.
function system(ctx: FormulaContext): string {
  const cols = ctx.headers.length
    ? ` Column headers (row 1): ${ctx.headers.map((h) => `${h.col}=${h.name}`).join(", ")}.`
    : "";
  return (
    "You are a spreadsheet formula assistant. Given a plain-language request, reply with ONLY " +
    "one formula that begins with =, no explanation and no code fences. Use standard functions " +
    "like SUM, AVERAGE, COUNT, COUNTIF, SUMIF, IF, VLOOKUP, INDEX, MATCH, ROUND." +
    ` The active cell is ${ctx.cell}.${cols}`
  );
}

// Pull the formula out of the reply: the first "=…" run, else the trimmed reply with an "="
// prepended, so a model that drops the leading "=" still yields a usable formula.
export function extractFormula(text: string): string {
  const m = text.match(/=\s*[^\n`]+/);
  let f = (m ? m[0] : text).trim().replace(/^`+|`+$/g, "").trim();
  if (f && !f.startsWith("=")) f = "=" + f;
  return f;
}

export function setupFormulaAssist(deps: FormulaAssistDeps): { open(anchor: HTMLElement): void } {
  const { wrap } = deps;
  const pop = document.createElement("div");
  pop.className = "sheetedit-fxa-pop";
  pop.hidden = true;

  const title = mk("div", "sheetedit-fxa-title", t("fxAssistTitle"));
  const desc = document.createElement("textarea");
  desc.className = "sheetedit-fxa-desc";
  desc.placeholder = t("fxAssistPlaceholder");
  desc.rows = 2;
  const progress = mk("div", "sheetedit-fxa-progress", "");
  const rlabel = mk("label", "sheetedit-fxa-rlabel", t("fxAssistResult"));
  const result = document.createElement("input");
  result.type = "text";
  result.className = "sheetedit-fxa-result";
  result.spellcheck = false;
  result.autocomplete = "off";
  rlabel.appendChild(result);
  const actions = mk("div", "sheetedit-fxa-actions", "");
  const genBtn = btn(t("fxAssistGenerate"), () => void run());
  const acceptBtn = btn(t("fxAssistAccept"), () => accept());
  const cancelBtn = btn(t("fxAssistCancel"), () => close());
  acceptBtn.classList.add("is-primary");
  actions.append(genBtn, acceptBtn, cancelBtn);
  pop.append(title, desc, progress, rlabel, actions);
  wrap.appendChild(pop);

  let running: { cancel(): void } | null = null;
  let anchor: HTMLElement | null = null;

  function open(at: HTMLElement): void {
    anchor = at;
    desc.value = "";
    result.value = "";
    progress.textContent = "";
    acceptBtn.disabled = true;
    genBtn.disabled = false;
    pop.hidden = false;
    position();
    desc.focus();
  }
  function position(): void {
    if (!anchor) return;
    const a = anchor.getBoundingClientRect();
    const w = wrap.getBoundingClientRect();
    pop.style.top = `${a.bottom - w.top + 4}px`;
    pop.style.left = `${Math.max(4, Math.min(a.left - w.left, w.width - pop.offsetWidth - 4))}px`;
  }
  function close(): void {
    running?.cancel();
    running = null;
    pop.hidden = true;
  }

  async function run(): Promise<void> {
    const request = desc.value.trim();
    if (!request) {
      progress.textContent = t("fxAssistNeedInput");
      return;
    }
    running?.cancel();
    result.value = "";
    acceptBtn.disabled = true;
    genBtn.disabled = true;
    progress.textContent = t("fxAssistLoading");
    try {
      const { runGenerate } = await import("localml/generate");
      const r = runGenerate(request, { task: "write", system: system(deps.getContext()) }, {
        onProgress: (p) =>
          (progress.textContent = p.stage === "download" ? t("fxAssistDownloading", { pct: Math.round(p.ratio * 100) }) : t("fxAssistGenerating")),
        onPartial: (text) => (result.value = extractFormula(text)),
        onDevice: () => (progress.textContent = t("fxAssistGenerating")),
      });
      running = { cancel: r.cancel };
      const out = await r.done;
      result.value = extractFormula(out.text || result.value);
      progress.textContent = t("fxAssistDone");
      acceptBtn.disabled = !result.value.trim();
    } catch (e) {
      console.error("[sheetedit] formula assist failed", e);
      progress.textContent = t("fxAssistError", { msg: (e as Error).message });
    } finally {
      running = null;
      genBtn.disabled = false;
    }
  }
  function accept(): void {
    const f = result.value.trim();
    if (f) deps.onAccept(f);
    close();
  }

  document.addEventListener("mousedown", (e) => {
    if (!pop.isConnected) return; // editor destroyed
    if (!pop.hidden && !pop.contains(e.target as Node) && e.target !== anchor) close();
  });
  document.addEventListener("keydown", (e) => {
    if (pop.isConnected && e.key === "Escape" && !pop.hidden) close();
  });
  window.addEventListener("resize", () => {
    if (!pop.hidden) position();
  });

  return { open };
}

function mk(tag: string, cls: string, text: string): HTMLElement {
  const e = document.createElement(tag);
  e.className = cls;
  if (text) e.textContent = text;
  return e;
}
function btn(label: string, fn: () => void): HTMLButtonElement {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "sheetedit-fxa-btn";
  b.textContent = label;
  b.addEventListener("mousedown", (e) => e.preventDefault());
  b.addEventListener("click", (e) => {
    e.preventDefault();
    fn();
  });
  return b;
}
