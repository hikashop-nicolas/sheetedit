import { t } from "../i18n";
import { nameError, refersFromLocal, refersToLocal, type NameError } from "../names";

// The Name Manager: every defined name with what it refers to (in the user's conventions), and
// an inline form to add, edit or delete one. Pure DOM; the editor applies each change to the model.

export interface NameRow {
  name: string;
  /** Invariant refers-to, as stored ("Sheet1!$A$1:$B$5", "0.5", "SUM(...)"). */
  ref: string;
  /** The sheet a sheet-scoped name belongs to; undefined for workbook scope. */
  sheet?: string;
}

export interface NameManagerCtx {
  wrap: HTMLElement;
  rows: () => NameRow[];
  /** The selection as an absolute, sheet-qualified reference: what a new name refers to. */
  selectionRef: () => string;
  /** Add (from null), edit (from a name) or delete (to null) a workbook-level name. */
  apply: (from: string | null, to: { name: string; ref: string } | null) => void;
}

const ERR_KEY: Record<NameError, string> = {
  empty: "nameErrEmpty",
  chars: "nameErrChars",
  cellRef: "nameErrCellRef",
  duplicate: "nameErrDuplicate",
};

let seq = 0;

export function openNameManager(ctx: NameManagerCtx, opts: { startNew?: boolean } = {}): HTMLElement {
  const restore = document.activeElement as HTMLElement | null;
  const id = `sheetedit-names-${++seq}`;
  const modal = document.createElement("div");
  modal.className = "sheetedit-modal sheetedit-names-modal";
  const card = document.createElement("div");
  card.className = "sheetedit-card is-wide sheetedit-names";
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");
  card.setAttribute("aria-labelledby", `${id}-title`);
  const h = document.createElement("h3");
  h.id = `${id}-title`;
  h.textContent = t("nameManager");
  card.appendChild(h);

  const btn = (label: string, role: string, primary = false): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.dataset.role = role;
    b.className = `sheetedit-dlg-btn${primary ? " is-primary" : ""}`;
    return b;
  };

  const list = document.createElement("table");
  list.className = "sheetedit-names-table";
  card.appendChild(list);

  // The add / edit form, shown under the list while a name is being written.
  const form = document.createElement("div");
  form.className = "sheetedit-names-form";
  form.hidden = true;
  const field = (key: string, label: string): HTMLInputElement => {
    const lbl = document.createElement("label");
    lbl.className = "sheetedit-field";
    const sp = document.createElement("span");
    sp.textContent = label;
    const inp = document.createElement("input");
    inp.type = "text";
    inp.className = "sheetedit-input";
    inp.dataset.field = key;
    inp.autocomplete = "off";
    inp.spellcheck = false;
    lbl.append(sp, inp);
    form.appendChild(lbl);
    return inp;
  };
  const nameIn = field("name", t("nameCol"));
  const refIn = field("refersTo", t("nameRefersTo"));
  const err = document.createElement("p");
  err.className = "sheetedit-names-error";
  err.id = `${id}-err`;
  err.setAttribute("role", "alert");
  form.appendChild(err);
  const formActions = document.createElement("div");
  formActions.className = "sheetedit-actions";
  const formCancel = btn(t("chartCancel"), "cancel"), formOk = btn(t("chartApply"), "ok", true);
  formActions.append(formCancel, formOk);
  form.appendChild(formActions);
  card.appendChild(form);

  const actions = document.createElement("div");
  actions.className = "sheetedit-actions";
  const addBtn = btn(t("nameNew"), "new"), closeBtn = btn(t("nameClose"), "close");
  actions.append(addBtn, closeBtn);
  card.appendChild(actions);

  let editing: string | null = null; // the name being edited, "" for a new one
  const showError = (msg: string, on: HTMLInputElement): void => {
    err.textContent = msg;
    for (const i of [nameIn, refIn]) i.removeAttribute("aria-invalid");
    on.setAttribute("aria-invalid", "true");
    on.setAttribute("aria-describedby", err.id);
    on.focus();
  };
  const closeForm = (): void => {
    form.hidden = true;
    actions.hidden = false;
    editing = null;
    err.textContent = "";
  };
  const openForm = (row: NameRow | null): void => {
    editing = row?.name ?? "";
    nameIn.value = row?.name ?? "";
    refIn.value = refersToLocal(row?.ref ?? ctx.selectionRef());
    err.textContent = "";
    for (const i of [nameIn, refIn]) i.removeAttribute("aria-invalid");
    form.hidden = false;
    actions.hidden = true;
    nameIn.focus();
  };

  const render = (): void => {
    list.textContent = "";
    const rows = ctx.rows();
    if (!rows.length) {
      const tr = list.insertRow();
      const td = tr.insertCell();
      td.className = "sheetedit-subtle";
      td.textContent = t("nameNone");
      return;
    }
    const head = list.createTHead().insertRow();
    for (const k of ["nameCol", "nameRefersTo", "nameScope", ""]) {
      const th = document.createElement("th");
      th.scope = "col";
      th.textContent = k ? t(k) : "";
      head.appendChild(th);
    }
    const body = list.createTBody();
    for (const row of rows) {
      const tr = body.insertRow();
      tr.dataset.name = row.name;
      tr.insertCell().textContent = row.name;
      const refCell = tr.insertCell();
      refCell.className = "sheetedit-names-ref";
      refCell.textContent = refersToLocal(row.ref);
      tr.insertCell().textContent = row.sheet ?? t("nameScopeWorkbook");
      const tools = tr.insertCell();
      tools.className = "sheetedit-names-tools";
      const edit = btn(t("nameEdit"), "edit"), del = btn(t("nameDelete"), "delete");
      edit.setAttribute("aria-label", t("nameEditAria", { name: row.name }));
      del.setAttribute("aria-label", t("nameDeleteAria", { name: row.name }));
      if (row.sheet != null) {
        // Sheet-scoped names are shown and kept, not edited here.
        edit.disabled = del.disabled = true;
        edit.title = del.title = t("nameSheetScoped");
      }
      edit.addEventListener("click", () => openForm(row));
      del.addEventListener("click", () => {
        ctx.apply(row.name, null);
        render();
        addBtn.focus();
      });
      tools.append(edit, del);
    }
  };

  const save = (): void => {
    if (editing == null) return;
    const name = nameIn.value.trim();
    const workbook = ctx.rows().filter((r) => r.sheet == null).map((r) => r.name);
    const bad = nameError(name, workbook, editing || undefined);
    if (bad) return showError(t(ERR_KEY[bad]), nameIn);
    const ref = refersFromLocal(refIn.value);
    if (!ref) return showError(t("nameErrRefersTo"), refIn);
    ctx.apply(editing || null, { name, ref });
    closeForm();
    render();
    const tr = Array.from(list.querySelectorAll<HTMLElement>("tr[data-name]")).find((r) => r.dataset.name === name);
    (tr?.querySelector<HTMLElement>('[data-role="edit"]') ?? addBtn).focus();
  };

  const close = (): void => {
    modal.remove();
    document.removeEventListener("keydown", onKey, true);
    restore?.focus?.();
  };
  const onKey = (e: KeyboardEvent): void => {
    if (!modal.isConnected) return;
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      if (!form.hidden) {
        closeForm();
        addBtn.focus();
      } else close();
    } else if (e.key === "Enter" && !form.hidden && (e.target === nameIn || e.target === refIn)) {
      e.preventDefault();
      save();
    } else if (e.key === "Tab") {
      // Keep focus inside the dialog.
      const items = Array.from(card.querySelectorAll<HTMLElement>("button, input")).filter((el) => !(el as HTMLButtonElement).disabled && el.offsetParent !== null);
      if (!items.length) return;
      const first = items[0]!, last = items[items.length - 1]!;
      if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
    }
  };
  document.addEventListener("keydown", onKey, true);
  addBtn.addEventListener("click", () => openForm(null));
  closeBtn.addEventListener("click", close);
  formCancel.addEventListener("click", () => { closeForm(); addBtn.focus(); });
  formOk.addEventListener("click", save);
  modal.addEventListener("mousedown", (e) => { if (e.target === modal) close(); });

  render();
  modal.appendChild(card);
  ctx.wrap.appendChild(modal);
  if (opts.startNew) openForm(null);
  else addBtn.focus();
  return modal;
}
