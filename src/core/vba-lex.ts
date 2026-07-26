// ---------------------------------------------------------------------------
// VBA lexer (Stage 1 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// Built from [MS-VBAL]'s lexical rules. The awkward parts, all of which the parser would rather not
// know about:
//   - a line ends a statement, and so does ":", so newlines are tokens rather than whitespace
//   - "_" at the end of a line continues it, joining two physical lines into one logical one
//   - "'" and REM run to the end of the line; a comment does NOT continue over "_"
//   - identifiers and keywords are case-insensitive
//   - a string doubles its quotes to escape them, and there is no backslash escape

export type TokenKind = "ident" | "keyword" | "number" | "string" | "date" | "op" | "eol" | "eof";

export interface Token {
  kind: TokenKind;
  /** The text as written, for error messages. */
  text: string;
  /** Identifiers and keywords, upper-cased, so comparisons need not care about case. */
  upper: string;
  /** Decoded value for a number or string literal. */
  value?: number | string;
  line: number;
}

/** Words the parser treats structurally. Everything else is an identifier, VBA-style. */
export const KEYWORDS = new Set([
  "AND", "AS", "BYREF", "BYVAL", "CALL", "CASE", "CONST", "DIM", "DO", "EACH", "ELSE", "ELSEIF",
  "END", "EQV", "ERROR", "EXIT", "FALSE", "FOR", "FUNCTION", "GET", "GOTO", "IF", "IMP", "IN", "IS",
  "LET", "LIKE", "LOOP", "MOD", "NEW", "NEXT", "NOT", "NOTHING", "ON", "OPTION", "OPTIONAL", "OR",
  "PARAMARRAY", "PRESERVE", "PRIVATE", "PROPERTY", "PUBLIC", "REDIM", "RESUME", "RETURN", "SELECT",
  "SET", "STATIC", "STEP", "SUB", "THEN", "TO", "TRUE", "UNTIL", "WEND", "WHILE", "WITH", "XOR",
  "ATTRIBUTE", "DECLARE", "ENUM", "TYPE", "FRIEND", "IMPLEMENTS", "WITHEVENTS", "EMPTY", "NULL",
]);

/** Multi-character operators, longest first so ">=" never lexes as ">" then "=". */
const OPERATORS = [">=", "<=", "<>", "&", "*", "/", "\\", "^", "+", "-", "=", "<", ">", "(", ")", ",", ".", ":", ";", "!", "#"];

const isDigit = (c: string): boolean => c >= "0" && c <= "9";
const isIdentStart = (c: string): boolean => /[A-Za-z_À-￿]/.test(c);
const isIdentChar = (c: string): boolean => /[A-Za-z0-9_À-￿]/.test(c);

export class VbaSyntaxError extends Error {
  constructor(message: string, readonly line: number) {
    super(`line ${line}: ${message}`);
    this.name = "VbaSyntaxError";
  }
}

/** Turn VBA source into tokens. Comments are dropped; line breaks are kept as "eol". */
export function lex(source: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  let line = 1;
  const src = source.replace(/\r\n?/g, "\n");
  const push = (kind: TokenKind, text: string, value?: number | string): void => {
    // `upper` exists only for matching names and symbols, so a literal gets none: otherwise the
    // string "-" would look to the parser like the minus operator, and "End" like the keyword.
    const matchable = kind === "ident" || kind === "keyword" || kind === "op";
    out.push({ kind, text, upper: matchable ? text.toUpperCase() : "", value, line });
  };

  while (i < src.length) {
    const c = src[i]!;

    // A line continuation is an underscore that ends the line: it and the newline both vanish.
    if (c === "_" && /^_[ \t]*\n/.test(src.slice(i))) {
      const nl = src.indexOf("\n", i);
      i = nl + 1;
      line++;
      continue;
    }
    if (c === " " || c === "\t") { i++; continue; }
    if (c === "\n") {
      push("eol", "\n");
      i++;
      line++;
      continue;
    }
    // A comment runs to the end of the line and never continues over an underscore.
    if (c === "'") {
      while (i < src.length && src[i] !== "\n") i++;
      continue;
    }
    if (/^REM\b/i.test(src.slice(i))) {
      while (i < src.length && src[i] !== "\n") i++;
      continue;
    }
    if (c === '"') {
      let s = "";
      i++;
      while (i < src.length) {
        if (src[i] === '"') {
          // A doubled quote is one literal quote; a single one ends the string.
          if (src[i + 1] === '"') { s += '"'; i += 2; continue; }
          i++;
          break;
        }
        if (src[i] === "\n") throw new VbaSyntaxError("unterminated string", line);
        s += src[i++];
      }
      push("string", s, s);
      continue;
    }
    // A date literal is fenced by #, e.g. #1/1/2020#. Kept as text; the runtime parses it.
    if (c === "#" && /^#[^#\n]*#/.test(src.slice(i))) {
      const end = src.indexOf("#", i + 1);
      const text = src.slice(i, end + 1);
      push("date", text, text.slice(1, -1));
      i = end + 1;
      continue;
    }
    // &H hex and &O octal come before the "&" operator, or they lex as concatenation.
    const amp = /^&([HO])([0-9A-Fa-f]+)&?/.exec(src.slice(i));
    if (amp) {
      const value = parseInt(amp[2]!, amp[1]!.toUpperCase() === "H" ? 16 : 8);
      push("number", amp[0], value);
      i += amp[0].length;
      continue;
    }
    if (isDigit(c) || (c === "." && isDigit(src[i + 1] ?? ""))) {
      const m = /^\d*\.?\d+(?:[eE][-+]?\d+)?[#!@&%]?/.exec(src.slice(i))!;
      // A trailing type character (#, !, @, &, %) declares the literal's type, not its value.
      push("number", m[0], Number(m[0].replace(/[#!@&%]$/, "")));
      i += m[0].length;
      continue;
    }
    if (isIdentStart(c)) {
      let j = i;
      while (j < src.length && isIdentChar(src[j]!)) j++;
      const text = src.slice(i, j);
      const upper = text.toUpperCase();
      push(KEYWORDS.has(upper) ? "keyword" : "ident", text);
      i = j;
      continue;
    }
    const op = OPERATORS.find((o) => src.startsWith(o, i));
    if (op) {
      push("op", op);
      i += op.length;
      continue;
    }
    throw new VbaSyntaxError(`unexpected character ${JSON.stringify(c)}`, line);
  }
  push("eof", "");
  return out;
}
