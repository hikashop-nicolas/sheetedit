// Text decoding and content sniffing for spreadsheet exports that are HTML or Excel 2003 XML.

export type SheetTextKind = "html" | "spreadsheetml";

const latin1 = (bytes: Uint8Array, n: number): string => {
  let s = "";
  const end = Math.min(bytes.length, n);
  for (let i = 0; i < end; i++) s += String.fromCharCode(bytes[i]!);
  return s;
};

function decodeWith(label: string, bytes: Uint8Array, fatal = false): string | null {
  try {
    return new TextDecoder(label, { fatal }).decode(bytes);
  } catch {
    return null;
  }
}

/** The BOM's encoding and length, if the bytes start with one. */
function bomOf(b: Uint8Array): { enc: string; len: number } | null {
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return { enc: "utf-8", len: 3 };
  if (b[0] === 0xff && b[1] === 0xfe) return { enc: "utf-16le", len: 2 };
  if (b[0] === 0xfe && b[1] === 0xff) return { enc: "utf-16be", len: 2 };
  return null;
}

/** UTF-16 without a BOM shows as "<" next to a zero byte. */
function bomlessUtf16(b: Uint8Array): string | null {
  if (b.length < 4) return null;
  if (b[0] === 0x3c && b[1] === 0 && b[3] === 0) return "utf-16le";
  if (b[0] === 0 && b[1] === 0x3c && b[2] === 0) return "utf-16be";
  return null;
}

/** The charset a document declares in its head (<meta charset>, http-equiv or an XML declaration). */
export function declaredCharset(head: string): string | undefined {
  const xml = /^\s*<\?xml[^>]*\bencoding\s*=\s*["']([\w.:-]+)["']/i.exec(head);
  if (xml) return xml[1];
  const meta = /<meta\b[^>]*?\bcharset\s*=\s*["']?\s*([\w.:-]+)/i.exec(head);
  if (meta) return meta[1];
  const mime = /^Content-Type:[^\r\n]*charset\s*=\s*"?([\w.:-]+)/im.exec(head);
  return mime?.[1];
}

/** Decode a text export: BOM first, then the declared charset, then UTF-8, then Windows-1252. */
export function decodeSheetText(bytes: Uint8Array): string {
  const bom = bomOf(bytes);
  if (bom) return decodeWith(bom.enc, bytes.subarray(bom.len)) ?? "";
  const u16 = bomlessUtf16(bytes);
  if (u16) return decodeWith(u16, bytes) ?? "";
  let label = declaredCharset(latin1(bytes, 4096))?.toLowerCase();
  // Pages often declare Latin-1 or ASCII while carrying Windows-1252 punctuation; browsers agree.
  if (label && /^(iso-8859-1|latin1|us-ascii|ascii|iso_8859-1)$/.test(label)) label = "windows-1252";
  if (label && !/^utf-?8$/.test(label)) {
    const out = decodeWith(label, bytes);
    if (out != null) return out;
  }
  return decodeWith("utf-8", bytes, true) ?? decodeWith("windows-1252", bytes) ?? "";
}

/** A cheap decode of the first bytes, enough to recognize the format. */
function headText(bytes: Uint8Array, n: number): string {
  const bom = bomOf(bytes);
  const enc = bom?.enc ?? bomlessUtf16(bytes);
  if (enc && enc !== "utf-8") return decodeWith(enc, bytes.subarray(bom?.len ?? 0, Math.min(bytes.length, n * 2))) ?? "";
  return latin1(bytes.subarray(bom?.len ?? 0), n);
}

const SSML_NS = "urn:schemas-microsoft-com:office:spreadsheet";
const EXCEL_NS = "urn:schemas-microsoft-com:office:excel";

/** Whether the text starts like an HTML document or fragment (after a BOM, doctype and comments). */
function startsLikeHtml(head: string): boolean {
  const s = head.replace(/^﻿/, "").replace(/^(?:\s+|<!--[\s\S]*?-->|<\?xml[^>]*\?>)*/, "");
  return /^<(?:!doctype\s+html|html|head|body|meta|title|style|table|link|div|p|h[1-6]|span|center|font|b|br|form)\b/i.test(s);
}

const startsLikeMime = (head: string): boolean => /^\s*(?:MIME-Version:|Content-Type:\s*multipart\/related)/i.test(head);

/**
 * What a text file holds, for routing: "spreadsheetml" for an Excel 2003 XML Spreadsheet, "html" for
 * an HTML table export (or Excel's single-file web page), null for anything else. Pass the file's
 * extension: ".xls" claims any HTML page with a table, while ".htm", ".html", ".xml" and no extension
 * claim only files that are clearly spreadsheet exports.
 */
export function sniffSheetText(bytes: Uint8Array, ext?: string): SheetTextKind | null {
  if (bytes.length < 6) return null;
  // Zips and compound files are binary workbooks, never these.
  if ((bytes[0] === 0x50 && bytes[1] === 0x4b) || (bytes[0] === 0xd0 && bytes[1] === 0xcf)) return null;
  const head = headText(bytes, 65536);
  const lenient = (ext ?? "").toLowerCase().replace(/^\./, "") === "xls";
  const trimmed = head.replace(/^﻿/, "").replace(/^(?:\s+|<!--[\s\S]*?-->)*/, "");
  if (/^<\?xml\b/i.test(trimmed) || /^<(?:[\w-]+:)?Workbook\b/.test(trimmed)) {
    const root = /<(?:([\w-]+):)?Workbook\b([^>]*)>/.exec(head);
    if (root && root[2]!.includes(SSML_NS)) return "spreadsheetml";
    if (/<\?mso-application\s+progid\s*=\s*["']Excel\.Sheet["']/i.test(head) && root) return "spreadsheetml";
    if (!/<html\b/i.test(head)) return null;
  }
  if (startsLikeMime(head)) return head.includes(EXCEL_NS) || /Excel\.Sheet/i.test(head) ? "html" : null;
  if (!startsLikeHtml(head)) return null;
  const strong = head.includes(EXCEL_NS) || /<meta\b[^>]*ProgId[^>]*Excel\.Sheet/i.test(head) || /mso-number-format|<x:ExcelWorkbook|\sx:(?:num|str|fmla)\b/i.test(head);
  if (strong) return "html";
  if (!lenient) return null;
  // A .xls named HTML page is a spreadsheet export when it carries a table anywhere.
  if (/<table\b/i.test(head)) return "html";
  return bytes.length > 65536 && hasTableTag(bytes) ? "html" : null;
}

/** A case-insensitive search for "<table" in the raw bytes (ASCII-compatible encodings). */
function hasTableTag(b: Uint8Array): boolean {
  const word = [0x74, 0x61, 0x62, 0x6c, 0x65];
  for (let i = b.indexOf(0x3c); i >= 0 && i < b.length - 6; i = b.indexOf(0x3c, i + 1)) {
    let k = 0;
    while (k < 5 && (b[i + 1 + k]! | 0x20) === word[k]) k++;
    if (k === 5 && !/[\w-]/.test(String.fromCharCode(b[i + 6]!))) return true;
  }
  return false;
}
