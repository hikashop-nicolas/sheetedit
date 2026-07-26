import { t } from "../i18n";

// A small modal form: a titled card of labelled fields (text / checkbox / colour / select / a
// read-only note) with live show/hide gating (showFor) and OK / Cancel. Shared by every authoring
// dialog. Pure DOM; the only host dependency is the element to mount into.

export type FormField = (
  | { key: string; label: string; type: "text" | "checkbox" | "color"; value?: string | boolean }
  | { key: string; label: string; type: "select"; options: { value: string; label: string }[]; value?: string }
  | { key: string; label: string; type: "note" }
) & { showFor?: { key: string; values: string[] } };

export function formDialog(wrap: HTMLElement, title: string, fields: FormField[], onOk: (vals: Record<string, string | boolean>) => void): void {
  const modal = document.createElement("div");
  modal.className = "sheetedit-form-modal";
  modal.style.cssText = "position:fixed;inset:0;z-index:70;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.45)";
  const card = document.createElement("div");
  // max-height + scroll: a long form (page setup) must never push its OK button off screen.
  card.style.cssText = "width:min(420px,94%);max-height:90vh;overflow-y:auto;background:var(--sheetedit-chrome,#2b2f36);color:var(--sheetedit-text,#e6e6e6);border:1px solid var(--sheetedit-border,#1c1f24);border-radius:10px;box-shadow:0 14px 44px rgba(0,0,0,.5);padding:16px;font:13px system-ui,sans-serif";
  const h = document.createElement("h3"); h.textContent = title; h.style.cssText = "margin:0 0 12px;font-size:15px"; card.appendChild(h);
  const fieldStyle = "font:inherit;background:var(--sheetedit-border,#1c1f24);border:1px solid var(--sheetedit-btn,#3a4047);border-radius:5px;color:var(--sheetedit-text,#e7eaf0);padding:6px 8px";
  const inputs: Record<string, HTMLInputElement | HTMLSelectElement> = {};
  const rows: Record<string, HTMLElement> = {};
  for (const f of fields) {
    if (f.type === "note") {
      const p = document.createElement("p");
      p.textContent = f.label;
      p.style.cssText = "margin:0 0 12px;color:var(--sheetedit-muted,#aab2bf);font-size:13px";
      card.appendChild(p); rows[f.key] = p;
      continue;
    }
    const inline = f.type === "checkbox";
    const lbl = document.createElement("label");
    lbl.style.cssText = "display:flex;" + (inline ? "align-items:center;gap:7px;" : "flex-direction:column;gap:4px;") + "margin-bottom:10px;font-size:13px";
    let inp: HTMLInputElement | HTMLSelectElement;
    if (f.type === "select") { const s = document.createElement("select"); s.style.cssText = fieldStyle; for (const o of f.options) { const op = document.createElement("option"); op.value = o.value; op.textContent = o.label; s.appendChild(op); } if (f.value != null) s.value = f.value; inp = s; }
    else { const i = document.createElement("input"); i.type = f.type; if (f.type === "checkbox") i.checked = !!f.value; else { i.value = (f.value as string) ?? ""; if (f.type === "text") i.style.cssText = fieldStyle; else i.style.cssText = "width:34px;height:26px;padding:0;border:1px solid var(--sheetedit-btn,#3a4047);border-radius:5px;background:none;cursor:pointer"; } inp = i; }
    inp.dataset.field = f.key;
    const sp = document.createElement("span"); sp.textContent = f.label; sp.style.color = "var(--sheetedit-muted,#aab2bf)";
    if (inline) lbl.append(inp, sp); else lbl.append(sp, inp);
    card.appendChild(lbl); inputs[f.key] = inp; rows[f.key] = lbl;
  }
  // Live show/hide of fields gated on another field's value.
  const applyVisibility = (): void => {
    for (const f of fields) {
      if (!f.showFor) continue;
      const driver = inputs[f.showFor.key];
      if (!rows[f.key]) continue;
      // The label carries an inline display:flex, which would override the hidden attribute, so
      // toggle display directly.
      const show = !!driver && f.showFor.values.includes(driver.value);
      rows[f.key]!.style.display = show ? "flex" : "none";
    }
  };
  for (const f of fields) {
    if (f.type === "select" && fields.some((o) => o.showFor?.key === f.key)) inputs[f.key]!.addEventListener("change", applyVisibility);
  }
  applyVisibility();
  // Sticky so OK / Cancel stay reachable while a long form scrolls inside the card.
  const actions = document.createElement("div");
  actions.style.cssText = "display:flex;justify-content:flex-end;gap:8px;margin-top:6px;position:sticky;bottom:-16px;padding:10px 0;background:var(--sheetedit-chrome,#2b2f36)";
  const btn = (label: string, primary: boolean): HTMLButtonElement => { const b = document.createElement("button"); b.textContent = label; b.dataset.role = primary ? "ok" : "cancel"; b.style.cssText = "font:inherit;font-size:13px;padding:6px 14px;border:1px solid var(--sheetedit-btn-border,#4a4f57);border-radius:6px;cursor:pointer;" + (primary ? "background:var(--sheetedit-accent,#6e7bff);border-color:var(--sheetedit-accent,#6e7bff);color:#fff" : "background:var(--sheetedit-btn,#3a3f47);color:var(--sheetedit-text,#e6e6e6)"); return b; };
  const cancel = btn(t("chartCancel"), false), ok = btn(t("chartApply"), true);
  const close = (): void => modal.remove();
  cancel.addEventListener("click", close);
  ok.addEventListener("click", () => { const vals: Record<string, string | boolean> = {}; for (const f of fields) { const inp = inputs[f.key]; if (!inp) continue; vals[f.key] = f.type === "checkbox" ? (inp as HTMLInputElement).checked : inp.value; } close(); onOk(vals); });
  actions.append(cancel, ok); card.appendChild(actions);
  modal.appendChild(card); wrap.appendChild(modal);
  modal.addEventListener("mousedown", (e) => { if (e.target === modal) close(); });
  const firstText = fields.find((f) => f.type === "text");
  if (firstText) inputs[firstText.key]!.focus();
}
