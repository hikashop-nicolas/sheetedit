// Shared ODS (ODF spreadsheet) namespaces, formula syntax conversion and
// length/colour parsing, used by the read, write and style modules.

export const ODS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
};
export const REPEAT_CAP = 1024;

/** Replace text outside single-quoted string literals. */
export function replaceOutsideStrings(s: string, fn: (chunk: string) => string): string {
  let out = "";
  let i = 0;
  while (i < s.length) {
    const q = s.indexOf('"', i);
    const qq = s.indexOf("'", i);
    let next = -1;
    let quote = '"';
    if (q === -1 && qq === -1) next = -1;
    else if (q === -1) {
      next = qq;
      quote = "'";
    } else if (qq === -1) {
      next = q;
      quote = '"';
    } else if (q < qq) {
      next = q;
      quote = '"';
    } else {
      next = qq;
      quote = "'";
    }
    if (next === -1) {
      out += fn(s.slice(i));
      break;
    }
    out += fn(s.slice(i, next));
    const end = s.indexOf(quote, next + 1);
    if (end === -1) {
      out += s.slice(next);
      break;
    }
    out += s.slice(next, end + 1);
    i = end + 1;
  }
  return out;
}

/** ODF formula (`of:=[.A1]+[.B1]`) -> A1 (`A1+B1`). */
export function odfToA1(odf: string): string {
  let core = odf.replace(/^of:=/, "").replace(/^=/, "");
  core = core.replace(/\[([^\]]*)\]/g, (_, inner: string) => {
    if (!inner || inner.includes("#")) return inner.replace(/\./g, ""); // #REF! etc.
    const parts = inner.split(":");
    const mapped = parts.map((part) => {
      const dot = part.lastIndexOf(".");
      const sheet = dot >= 0 ? part.slice(0, dot) : "";
      const ref = dot >= 0 ? part.slice(dot + 1) : part;
      return { sheet, ref };
    });
    const sheet = mapped[0]!.sheet;
    const cells = mapped.map((m) => m.ref).join(":");
    return (sheet ? sheet + "!" : "") + cells;
  });
  return replaceOutsideStrings(core, (chunk) => chunk.replace(/;/g, ","));
}

/** A1 (`A1+B1`) -> ODF formula (`of:=[.A1]+[.B1]`). Used only for user-typed formulas. */
export function a1ToOdf(a1: string): string {
  const refRe =
    /(?<![A-Za-z0-9_.$])(?:('[^']*'|[A-Za-z_][A-Za-z0-9_]*)!)?(\$?[A-Z]{1,3}\$?[0-9]+)(?::(\$?[A-Z]{1,3}\$?[0-9]+))?(?![A-Za-z0-9_(])/g;
  const converted = replaceOutsideStrings(a1, (chunk) => {
    const semi = chunk.replace(/,/g, ";");
    return semi.replace(refRe, (_m, sheet: string | undefined, c1: string, c2?: string) => {
      const sh = sheet ?? "";
      const range = c2 ? `${c1}:.${c2}` : c1;
      return `[${sh}.${range}]`;
    });
  });
  return "of:=" + converted;
}

export function odsCellText(cell: Element): string {
  return Array.from(cell.getElementsByTagName("text:p"))
    .map((p) => p.textContent ?? "")
    .join("\n");
}

// Convert an ODF length ("2.5cm", "96pt", "1in", "0.45mm", "12px") to CSS px.
export function odsLenToPx(len: string | null | undefined): number | undefined {
  if (!len) return undefined;
  const m = /^([\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(len.trim());
  if (!m) return undefined;
  const v = parseFloat(m[1]!);
  switch (m[2]) {
    case "cm":
      return Math.round((v * 96) / 2.54);
    case "mm":
      return Math.round((v * 96) / 25.4);
    case "in":
      return Math.round(v * 96);
    case "pt":
      return Math.round((v * 96) / 72);
    case "pc":
      return Math.round(v * 16);
    default:
      return Math.round(v); // px or unitless
  }
}

export const odsColorOf = (v: string | null): string | undefined =>
  v && v !== "transparent" && v !== "none" ? (/^#[0-9a-fA-F]{6}$/.test(v) ? v.toLowerCase() : v) : undefined;

// A border value is like "0.5pt solid #000000" or "none"; keep the colour if it draws.
export function odsBorderColor(v: string | null): string | undefined {
  if (!v || v === "none" || /(^|\s)0(\.0+)?(pt|cm|mm|in|px)?\s/.test(" " + v + " ")) return undefined;
  const m = /#[0-9a-fA-F]{6}/.exec(v);
  return m ? m[0].toLowerCase() : "#000000";
}
