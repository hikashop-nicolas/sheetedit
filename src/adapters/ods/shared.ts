// Shared ODS (ODF spreadsheet) namespaces, formula syntax conversion and
// length/colour parsing, used by the read, write and style modules.
import type { Cell, Phonetic } from "../../core/model";

export const ODS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  number: "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  dc: "http://purl.org/dc/elements/1.1/",
  calcext: "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
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

// A cell hyperlink from an ODF <text:a xlink:href>. Internal targets are "#Sheet1.A1"
// (converted to the "Sheet1!A1" form the editor follows); everything else is external.
export function odsCellLink(cell: Element): Cell["link"] {
  const a = cell.getElementsByTagName("text:a")[0];
  if (!a) return undefined;
  const href = a.getAttribute("xlink:href") ?? "";
  if (!href) return undefined;
  if (href.startsWith("#")) return { href: href.slice(1).replace(".", "!"), internal: true };
  return { href };
}

// Cell notes from ODF <office:annotation> (dc:creator + text:p lines), newest last.
export function odsCellComments(cell: Element): Cell["comments"] {
  const out: NonNullable<Cell["comments"]> = [];
  for (const an of Array.from(cell.getElementsByTagName("office:annotation"))) {
    const author = an.getElementsByTagName("dc:creator")[0]?.textContent?.trim() || undefined;
    const text = Array.from(an.getElementsByTagName("text:p")).map((p) => p.textContent ?? "").join("\n").trim();
    if (text) out.push({ author, text });
  }
  return out.length ? out : undefined;
}

// Ruby-aware cell text: the base text (plain text + <text:ruby-base>) plus the phonetic
// (furigana) runs from <text:ruby-text>. Without this, textContent folds the reading into
// the value (東京 + トウキョウ run together), the ODF analogue of the xlsx rPh bug.
export function odsCellRich(cell: Element): { text: string; phonetic?: Phonetic[] } {
  const out = { text: "", phonetic: [] as Phonetic[] };
  const walk = (node: Node): void => {
    for (const ch of Array.from(node.childNodes)) {
      if (ch.nodeType === 3) {
        out.text += ch.textContent ?? "";
        continue;
      }
      if (ch.nodeType !== 1) continue;
      const el = ch as Element;
      switch (el.localName) {
        case "ruby": {
          const base = el.getElementsByTagName("text:ruby-base")[0]?.textContent ?? "";
          const reading = el.getElementsByTagName("text:ruby-text")[0]?.textContent ?? "";
          const sb = out.text.length;
          out.text += base;
          if (reading) out.phonetic.push({ sb, eb: out.text.length, reading });
          break;
        }
        case "s":
          out.text += " ".repeat(Math.max(1, Number(el.getAttribute("text:c") || "1")));
          break;
        case "tab":
          out.text += "\t";
          break;
        case "line-break":
          out.text += "\n";
          break;
        default:
          walk(el); // spans and other inline wrappers: descend
      }
    }
  };
  Array.from(cell.getElementsByTagName("text:p")).forEach((p, i) => {
    if (i > 0) out.text += "\n";
    walk(p);
  });
  return out.phonetic.length ? { text: out.text, phonetic: out.phonetic } : { text: out.text };
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
