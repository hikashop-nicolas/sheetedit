import { t } from "../i18n";
import type { Workbook } from "../model";
import { getCell } from "../model";

// Find and replace bar. Finding walks every sheet (navigation switches the
// active sheet and focuses the cell); replacing acts on the active sheet only,
// which keeps each replacement inside the per-sheet undo model.

export interface FindBarCtx {
  container: HTMLElement;
  beforeEl: HTMLElement | null;
  getWorkbook(): Workbook;
  getActiveSheet(): number;
  setActiveSheet(i: number): void;
  focusCell(r: number, c: number): void;
  commitValue(r: number, c: number, raw: string): void;
  /** Batched replacement on the active sheet (one undo step). */
  applyBatch(changes: { r: number; c: number; raw: string }[]): void;
}

interface Match {
  sheet: number;
  r: number;
  c: number;
}

const rawOf = (wb: Workbook, m: Match): string => {
  const cell = getCell(wb.sheets[m.sheet]!, m.r, m.c);
  return cell ? (cell.formula != null ? "=" + cell.formula : cell.value) : "";
};

/** Case-insensitive single replacement of `query` inside `raw`. */
export const replaceOnce = (raw: string, query: string, replacement: string): string => {
  const i = raw.toLowerCase().indexOf(query.toLowerCase());
  return i === -1 ? raw : raw.slice(0, i) + replacement + raw.slice(i + query.length);
};

/** Case-insensitive replacement of every occurrence, advancing past each
    replacement so "a" -> "aa" cannot loop. */
export const replaceEvery = (raw: string, query: string, replacement: string): string => {
  const lower = query.toLowerCase();
  let out = "";
  let i = 0;
  for (;;) {
    const j = raw.toLowerCase().indexOf(lower, i);
    if (j === -1) return out + raw.slice(i);
    out += raw.slice(i, j) + replacement;
    i = j + query.length;
  }
};

export function setupFindBar(ctx: FindBarCtx) {
  const bar = document.createElement("div");
  bar.className = "sheetedit-findbar";
  bar.hidden = true;
  const find = document.createElement("input");
  find.type = "text";
  find.placeholder = t("frFind");
  find.setAttribute("aria-label", t("frFind"));
  const replace = document.createElement("input");
  replace.type = "text";
  replace.placeholder = t("frReplace");
  replace.setAttribute("aria-label", t("frReplace"));
  const count = document.createElement("span");
  count.className = "sheetedit-findcount";
  count.setAttribute("aria-live", "polite");
  const mkBtn = (label: string, title: string, fn: () => void) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sheetedit-btn";
    b.textContent = label;
    b.title = title;
    b.addEventListener("click", fn);
    return b;
  };
  bar.append(
    find,
    replace,
    count,
    mkBtn("‹", t("frPrev"), () => step(-1)),
    mkBtn("›", t("frNext"), () => step(1)),
    mkBtn(t("frReplaceOne"), t("frReplaceOneTitle"), () => replaceCurrent()),
    mkBtn(t("frReplaceAll"), t("frReplaceAllTitle"), () => replaceAll()),
    mkBtn("✕", t("frClose"), () => hide()),
  );
  ctx.container.insertBefore(bar, ctx.beforeEl);

  let matches: Match[] = [];
  let current = -1;

  const search = (): void => {
    const wb = ctx.getWorkbook();
    const q = find.value.toLowerCase();
    matches = [];
    if (q) {
      wb.sheets.forEach((sheet, si) => {
        const cells = [...sheet.cells.values()].sort((a, b) => a.row - b.row || a.col - b.col);
        for (const cell of cells) {
          const raw = cell.formula != null ? "=" + cell.formula : cell.value;
          if (raw.toLowerCase().includes(q)) matches.push({ sheet: si, r: cell.row, c: cell.col });
        }
      });
    }
    current = matches.length ? 0 : -1;
    updateCount();
  };

  const updateCount = (): void => {
    count.textContent = matches.length ? t("frCount", { i: current + 1, n: matches.length }) : find.value ? t("frNone") : "";
  };

  const goTo = (m: Match): void => {
    if (m.sheet !== ctx.getActiveSheet()) ctx.setActiveSheet(m.sheet);
    ctx.focusCell(m.r, m.c);
  };

  const step = (dir: 1 | -1): void => {
    if (!matches.length) search();
    if (!matches.length) return;
    current = (current + dir + matches.length) % matches.length;
    updateCount();
    goTo(matches[current]!);
  };

  const replaceCurrent = (): void => {
    if (current === -1 || !matches.length) {
      step(1);
      return;
    }
    const m = matches[current]!;
    if (m.sheet !== ctx.getActiveSheet()) {
      goTo(m); // bring the match into the active sheet first; a second click replaces
      return;
    }
    const wb = ctx.getWorkbook();
    ctx.commitValue(m.r, m.c, replaceOnce(rawOf(wb, m), find.value, replace.value));
    search();
    if (matches.length) goTo(matches[Math.min(current, matches.length - 1)]!);
  };

  const replaceAll = (): void => {
    const q = find.value;
    if (!q) return;
    const wb = ctx.getWorkbook();
    const si = ctx.getActiveSheet();
    search();
    const changes: { r: number; c: number; raw: string }[] = [];
    for (const m of matches) {
      if (m.sheet !== si) continue;
      // Replace every occurrence within the cell, not just the first.
      changes.push({ r: m.r, c: m.c, raw: replaceEvery(rawOf(wb, m), q, replace.value) });
    }
    if (changes.length) ctx.applyBatch(changes);
    search();
  };

  find.addEventListener("input", () => search());
  find.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      step(e.shiftKey ? -1 : 1);
    } else if (e.key === "Escape") {
      e.preventDefault();
      hide();
    }
  });
  replace.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      replaceCurrent();
    } else if (e.key === "Escape") {
      e.preventDefault();
      hide();
    }
  });

  const show = (): void => {
    bar.hidden = false;
    find.focus();
    find.select();
  };
  const hide = (): void => {
    bar.hidden = true;
  };

  return {
    show,
    toggle: () => (bar.hidden ? show() : hide()),
    teardown: () => bar.remove(),
  };
}
