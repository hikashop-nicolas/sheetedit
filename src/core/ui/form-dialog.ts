import { t } from "../i18n";

// A small modal form: a titled card of labelled fields (text / checkbox / colour / select / a
// read-only note) with live show/hide gating (showFor) and OK / Cancel. Shared by every authoring
// dialog. Pure DOM; the only host dependency is the element to mount into. All styling lives in
// sheetedit.css.

export type FormField = (
  | { key: string; label: string; type: "text" | "checkbox" | "color"; value?: string | boolean }
  | { key: string; label: string; type: "select"; options: { value: string; label: string }[]; value?: string }
  | { key: string; label: string; type: "note" }
) & { showFor?: { key: string; values: string[] } };

export function formDialog(wrap: HTMLElement, title: string, fields: FormField[], onOk: (vals: Record<string, string | boolean>) => void): void {
  const modal = document.createElement("div");
  modal.className = "sheetedit-modal sheetedit-form-modal";
  const card = document.createElement("div");
  card.className = "sheetedit-card";
  const h = document.createElement("h3");
  h.textContent = title;
  card.appendChild(h);
  const inputs: Record<string, HTMLInputElement | HTMLSelectElement> = {};
  const rows: Record<string, HTMLElement> = {};
  for (const f of fields) {
    if (f.type === "note") {
      const p = document.createElement("p");
      p.textContent = f.label;
      p.className = "sheetedit-note";
      card.appendChild(p);
      rows[f.key] = p;
      continue;
    }
    const inline = f.type === "checkbox";
    const lbl = document.createElement("label");
    lbl.className = `sheetedit-field${inline ? " is-inline" : ""}`;
    let inp: HTMLInputElement | HTMLSelectElement;
    if (f.type === "select") {
      const s = document.createElement("select");
      s.className = "sheetedit-input";
      for (const o of f.options) {
        const op = document.createElement("option");
        op.value = o.value;
        op.textContent = o.label;
        s.appendChild(op);
      }
      if (f.value != null) s.value = f.value;
      inp = s;
    } else {
      const i = document.createElement("input");
      i.type = f.type;
      if (f.type === "checkbox") i.checked = !!f.value;
      else {
        i.value = (f.value as string) ?? "";
        i.className = f.type === "text" ? "sheetedit-input" : "sheetedit-color";
      }
      inp = i;
    }
    inp.dataset.field = f.key;
    const sp = document.createElement("span");
    sp.textContent = f.label;
    if (inline) lbl.append(inp, sp);
    else lbl.append(sp, inp);
    card.appendChild(lbl);
    inputs[f.key] = inp;
    rows[f.key] = lbl;
  }
  // Live show/hide of fields gated on another field's value.
  const applyVisibility = (): void => {
    for (const f of fields) {
      if (!f.showFor) continue;
      const driver = inputs[f.showFor.key];
      const row = rows[f.key];
      if (!row) continue;
      // The row carries display:flex from its class, which would override the hidden attribute.
      row.classList.toggle("is-hidden", !(driver && f.showFor.values.includes(driver.value)));
    }
  };
  for (const f of fields) {
    if (f.type === "select" && fields.some((o) => o.showFor?.key === f.key)) inputs[f.key]!.addEventListener("change", applyVisibility);
  }
  applyVisibility();
  const actions = document.createElement("div");
  actions.className = "sheetedit-actions";
  const btn = (label: string, primary: boolean): HTMLButtonElement => {
    const b = document.createElement("button");
    b.textContent = label;
    b.dataset.role = primary ? "ok" : "cancel";
    b.className = `sheetedit-dlg-btn${primary ? " is-primary" : ""}`;
    return b;
  };
  const cancel = btn(t("chartCancel"), false), ok = btn(t("chartApply"), true);
  const close = (): void => modal.remove();
  cancel.addEventListener("click", close);
  ok.addEventListener("click", () => {
    const vals: Record<string, string | boolean> = {};
    for (const f of fields) {
      const inp = inputs[f.key];
      if (!inp) continue;
      vals[f.key] = f.type === "checkbox" ? (inp as HTMLInputElement).checked : inp.value;
    }
    close();
    onOk(vals);
  });
  actions.append(cancel, ok);
  card.appendChild(actions);
  modal.appendChild(card);
  wrap.appendChild(modal);
  modal.addEventListener("mousedown", (e) => { if (e.target === modal) close(); });
  const firstText = fields.find((f) => f.type === "text");
  if (firstText) inputs[firstText.key]!.focus();
}
