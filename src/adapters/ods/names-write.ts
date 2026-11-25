import { colToLetters, type Workbook } from "../../core/model";
import { parseRangeRef } from "../../core/names";
import { ODS, a1ToOdf } from "./shared";

// Workbook-level names into <table:named-expressions>: a range becomes a table:named-range, a
// constant or formula a table:named-expression. Names inside a table (sheet-scoped) are untouched.

const odfSheet = (name: string): string => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);

/** "Sheet1!$A$1:$B$2" -> "$Sheet1.$A$1:.$B$2", or null when the target is not a plain range. */
function odfRange(ref: string, fallbackSheet: string): { addr: string; base: string } | null {
  const p = parseRangeRef(ref);
  if (!p) return null;
  const sh = "$" + odfSheet(p.sheet ?? fallbackSheet);
  const a = `$${colToLetters(p.rect.c1)}$${p.rect.r1}`, b = `$${colToLetters(p.rect.c2)}$${p.rect.r2}`;
  return { addr: `${sh}.${a}${a === b ? "" : ":." + b}`, base: `${sh}.${a}` };
}

export function writeOdsDefinedNames(wb: Workbook): void {
  const doc = wb.contentDoc;
  if (!wb.namesDirty || !doc) return;
  const ss = doc.getElementsByTagName("office:spreadsheet")[0];
  if (!ss) return;
  wb.namesDirty = false;
  const want = wb.definedNames ?? new Map<string, string>();
  let box = Array.from(ss.children).find((e) => e.localName === "named-expressions") ?? null;
  if (box) while (box.firstChild) box.removeChild(box.firstChild);
  if (!want.size) {
    box?.parentNode?.removeChild(box);
    return;
  }
  if (!box) {
    box = doc.createElementNS(ODS.table, "table:named-expressions");
    const tables = Array.from(ss.children).filter((e) => e.localName === "table");
    const last = tables[tables.length - 1];
    ss.insertBefore(box, last ? last.nextSibling : ss.firstChild);
  }
  const first = wb.sheets[0]?.name ?? "Sheet1";
  for (const [name, ref] of want) {
    const range = odfRange(ref, first);
    const el = doc.createElementNS(ODS.table, range ? "table:named-range" : "table:named-expression");
    el.setAttributeNS(ODS.table, "table:name", name);
    if (range) {
      el.setAttributeNS(ODS.table, "table:base-cell-address", range.base);
      el.setAttributeNS(ODS.table, "table:cell-range-address", range.addr);
    } else {
      el.setAttributeNS(ODS.table, "table:base-cell-address", `$${odfSheet(first)}.$A$1`);
      el.setAttributeNS(ODS.table, "table:expression", a1ToOdf(ref));
    }
    box.appendChild(el);
  }
}
