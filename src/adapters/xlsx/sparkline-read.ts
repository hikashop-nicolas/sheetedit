import { parseA1Ref, type Sheet } from "../../core/model";

// Read sparklines (in-cell mini charts) from the worksheet's extLst:
//   <ext><x14:sparklineGroups><x14:sparklineGroup type="column">...<x14:sparklines>
//     <x14:sparkline><xm:f>Sheet1!B2:F2</xm:f><xm:sqref>H2</xm:sqref></x14:sparkline>
// Namespace-prefix-agnostic (elements matched by local name). Preserved on save (the extLst is
// kept verbatim); this only populates sheet.sparklines for rendering.

const kids = (el: Element, local: string): Element[] => Array.from(el.getElementsByTagName("*")).filter((e) => e.localName === local);
const childText = (el: Element, local: string): string => { const c = Array.from(el.children).find((x) => x.localName === local); return c?.textContent?.trim() ?? ""; };

/** Populate sheet.sparklines from the worksheet's x14 sparkline groups. */
export function readSparklines(sheet: Sheet, doc: Document): void {
  const groups = kids(doc.documentElement, "sparklineGroup");
  if (!groups.length) return;
  const out: NonNullable<Sheet["sparklines"]> = [];
  for (const g of groups) {
    const type = (g.getAttribute("type") as "line" | "column" | "stacked") || "line";
    const cs = Array.from(g.children).find((c) => c.localName === "colorSeries");
    const rgb = cs?.getAttribute("rgb");
    const color = rgb ? `#${rgb.slice(-6)}` : type === "line" ? "#376092" : "#376092";
    for (const sp of kids(g, "sparkline")) {
      const dataRef = childText(sp, "f");
      const host = childText(sp, "sqref");
      const p = parseA1Ref(host.replace(/\$/g, "").split(":")[0] ?? "");
      if (!dataRef || !p) continue;
      out.push({ type, color, host: { r: p.row, c: p.col }, dataRef });
    }
  }
  if (out.length) sheet.sparklines = out;
}
