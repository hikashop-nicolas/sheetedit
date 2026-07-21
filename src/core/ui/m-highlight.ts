// Lightweight, never-throws tokenizer for Power Query M, producing highlighted HTML for the
// editor's overlay layer. Deliberately lexical-only (no parser): it must always return
// something sensible for mid-edit, syntactically-invalid input, and run on every keystroke.

const KEYWORDS = new Set([
  "and", "as", "each", "else", "error", "false", "if", "in", "is", "let", "meta", "not",
  "otherwise", "or", "section", "shared", "then", "true", "try", "type", "nullable", "optional",
]);

// #-prefixed intrinsic keywords (constructors and special literals).
const HASH_KEYWORDS = new Set([
  "#binary", "#date", "#datetime", "#datetimezone", "#duration", "#table", "#time",
  "#sections", "#shared", "#infinity", "#nan",
]);

const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const span = (cls: string, s: string): string => `<span class="mtok-${cls}">${esc(s)}</span>`;

const isIdStart = (c: string): boolean => /[A-Za-z_]/.test(c);
const isIdChar = (c: string): boolean => /[A-Za-z0-9_.]/.test(c);
const isDigit = (c: string): boolean => c >= "0" && c <= "9";

/** Tokenize M and return HTML (spans with mtok-* classes). Escapes all text content. */
export function highlightM(src: string): string {
  let out = "";
  let i = 0;
  const n = src.length;
  while (i < n) {
    const c = src[i]!;

    if (/\s/.test(c)) {
      let j = i + 1;
      while (j < n && /\s/.test(src[j]!)) j++;
      out += esc(src.slice(i, j));
      i = j;
      continue;
    }
    // line comment
    if (c === "/" && src[i + 1] === "/") {
      let j = i + 2;
      while (j < n && src[j] !== "\n") j++;
      out += span("com", src.slice(i, j));
      i = j;
      continue;
    }
    // block comment (tolerate unterminated)
    if (c === "/" && src[i + 1] === "*") {
      let j = i + 2;
      while (j < n && !(src[j] === "*" && src[j + 1] === "/")) j++;
      j = Math.min(n, j + 2);
      out += span("com", src.slice(i, j));
      i = j;
      continue;
    }
    // text literal (tolerate unterminated; "" is an escaped quote)
    if (c === '"') {
      let j = i + 1;
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            j += 2;
            continue;
          }
          j++;
          break;
        }
        j++;
      }
      out += span("str", src.slice(i, j));
      i = j;
      continue;
    }
    // quoted identifier #"..."
    if (c === "#" && src[i + 1] === '"') {
      let j = i + 2;
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            j += 2;
            continue;
          }
          j++;
          break;
        }
        j++;
      }
      out += span("id", src.slice(i, j));
      i = j;
      continue;
    }
    // #-keyword (#date, #table, ...)
    if (c === "#" && isIdStart(src[i + 1] ?? "")) {
      let j = i + 1;
      while (j < n && /[A-Za-z]/.test(src[j]!)) j++;
      const w = src.slice(i, j);
      out += span(HASH_KEYWORDS.has(w) ? "kw" : "id", w);
      i = j;
      continue;
    }
    // number
    if (isDigit(c) || (c === "." && isDigit(src[i + 1] ?? ""))) {
      let j = i;
      if (c === "0" && (src[i + 1] === "x" || src[i + 1] === "X")) {
        j = i + 2;
        while (j < n && /[0-9a-fA-F]/.test(src[j]!)) j++;
      } else {
        while (j < n && isDigit(src[j]!)) j++;
        if (src[j] === ".") {
          j++;
          while (j < n && isDigit(src[j]!)) j++;
        }
        if (src[j] === "e" || src[j] === "E") {
          j++;
          if (src[j] === "+" || src[j] === "-") j++;
          while (j < n && isDigit(src[j]!)) j++;
        }
      }
      out += span("num", src.slice(i, j));
      i = j;
      continue;
    }
    // identifier / keyword / library call (Table.SelectRows, List.Sum, ...)
    if (isIdStart(c)) {
      let j = i + 1;
      while (j < n && isIdChar(src[j]!)) j++;
      const w = src.slice(i, j);
      const cls = KEYWORDS.has(w) ? "kw" : /^[A-Z][A-Za-z0-9]*\.[A-Za-z]/.test(w) ? "fn" : "id";
      out += span(cls, w);
      i = j;
      continue;
    }
    // operators / punctuation
    if (/[=<>+\-*/&@?,;:(){}[\].]/.test(c)) {
      out += span("op", c);
      i++;
      continue;
    }
    out += esc(c);
    i++;
  }
  // Trailing newline keeps the overlay's scroll height aligned with the textarea.
  return out + "\n";
}
