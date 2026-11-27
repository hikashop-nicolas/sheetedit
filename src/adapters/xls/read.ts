// BIFF reader ([MS-XLS]): the Workbook stream of a legacy .xls (BIFF8, BIFF5/7, and the
// single-sheet BIFF2-4 files) into an XlsBook. Unknown records are skipped; a damaged stream
// fails with a clear error rather than looping.
import { colToLetters } from "../../core/model";
import { type BiffRecord, Cursor, readRichString, readString, readStringNoCch, records, utf16le, type XlsString } from "./biff";
import { isCfb, readCfbStorage, readCfbStreams } from "./cfb";
import { type BlockCipher, rc4CryptoApi, rc4Standard, Rc4, xorKeyAndArray } from "./crypto";
import { type Blip, type DrawNode, readBlipStore, readDrawingTree, readSheetShapes } from "./escher";
import { type ChartEnv, chartRecords, readChart } from "./chart-read";
import { icvCss } from "./palette";
import { decodeFormula, ERRORS, type FormulaContext } from "./formula-read";
import { DEFAULT_PALETTE } from "./palette";
import type { XlsBook, XlsCell, XlsCf, XlsFont, XlsName, XlsRange, XlsSheet, XlsSupbook, XlsTxo, XlsXf } from "./types";

export interface XlsReadOptions {
  /** The password for an encrypted workbook. */
  password?: string;
}

/** Thrown for an encrypted workbook when no password, or a wrong one, was given. */
export class XlsPasswordError extends Error {
  constructor(public code: "password-required" | "password-wrong") {
    super(code === "password-required" ? "this workbook is password-protected" : "wrong password");
    this.name = "XlsPasswordError";
  }
}

const BOF_TYPES = new Set([0x0009, 0x0209, 0x0409, 0x0809]);
const DEFAULT_PASSWORD = "VelvetSweatshop";

/** Whether these bytes are a raw BIFF2-4 stream (not wrapped in a compound file). */
export const isRawBiff = (b: Uint8Array): boolean => b.length >= 8 && BOF_TYPES.has(b[0]! | (b[1]! << 8)) && (b[2]! | (b[3]! << 8)) >= 4 && (b[2]! | (b[3]! << 8)) <= 20;

/** Read a .xls (compound file or raw BIFF stream) into an XlsBook. */
export function readXlsBook(bytes: Uint8Array, opts: XlsReadOptions = {}): XlsBook {
  let stream: Uint8Array | undefined;
  let summary: Map<string, Uint8Array> | undefined;
  let vba: { name: string; data: Uint8Array }[] | undefined;
  if (isCfb(bytes)) {
    const { streams } = readCfbStreams(bytes);
    stream = streams.get("Workbook") ?? streams.get("Book") ?? [...streams.entries()].find(([k]) => /^(workbook|book)$/i.test(k))?.[1];
    if (!stream) {
      if (streams.has("EncryptionInfo") || streams.has("EncryptedPackage")) throw new Error("password-protected .xlsx (OOXML encryption) is not supported");
      throw new Error("not an Excel workbook: no Workbook stream in the compound file");
    }
    summary = new Map();
    for (const [k, v] of streams) if (/SummaryInformation$|^\u0001CompObj$/.test(k)) summary.set(k, v);
    vba = readCfbStorage(bytes, "_VBA_PROJECT_CUR");
    // The project's signature sits at the root; it travels with the project as _VBA_SIGNATURE.
    const sig = streams.get("\u0005DigitalSignature");
    if (vba.length && sig) vba.push({ name: "_VBA_SIGNATURE", data: sig });
  } else if (isRawBiff(bytes)) stream = bytes;
  else throw new Error("not an .xls workbook");
  const book = parseWorkbookStream(stream, opts);
  if (summary?.size) book.summary = summary;
  if (vba?.length) book.vba = vba;
  return book;
}

function newSheet(name: string, state: number): XlsSheet {
  return { name, state, cells: new Map(), rows: new Map(), cols: [], merges: [], links: [], notes: [], dvs: [], condFmts: [], images: [], print: {}, charts: 0, shapes: 0, chartDrawings: 0, chartObjs: [] };
}

/** Records whose payload stays clear in an encrypted stream. */
const CLEAR = new Set([0x0809, 0x002f, 0x00e1, 0x0194, 0x0195, 0x0196, 0x0138]);

/** Decrypt every record payload after FILEPASS in place. */
function decryptStream(s: Uint8Array, from: number, rc4: BlockCipher | null, xor: Uint8Array | null): void {
  let block = -1, rc: Rc4 | null = null, rcPos = 0;
  const rcRange = (a: number, b: number, skip: boolean): void => {
    while (a < b) {
      const blk = Math.floor(a / 1024);
      if (blk !== block || !rc || rcPos > a) { block = blk; rc = new Rc4(rc4!.blockKey(blk)); rcPos = blk * 1024; }
      if (rcPos < a) { rc.apply(s, rcPos, a, true); rcPos = a; }
      const end = Math.min(b, (blk + 1) * 1024);
      rc.apply(s, a, end, skip);
      rcPos = end;
      a = end;
    }
  };
  let p = from;
  while (p + 4 <= s.length) {
    const type = s[p]! | (s[p + 1]! << 8);
    const len = s[p + 2]! | (s[p + 3]! << 8);
    const d = p + 4, e = Math.min(s.length, d + len);
    if (rc4) {
      rcRange(p, d, true);
      if (CLEAR.has(type)) rcRange(d, e, true);
      else if (type === 0x85) { rcRange(d, Math.min(e, d + 4), true); rcRange(Math.min(e, d + 4), e, false); }
      else rcRange(d, e, false);
    } else if (xor && !CLEAR.has(type)) {
      const start = type === 0x85 ? d + 4 : d;
      for (let i = start; i < e; i++) {
        const v = s[i]! ^ xor[(i + len) & 15]!; // XorArrayIndex = (data offset + data length) % 16, then +1 per byte
        s[i] = ((v >> 5) | (v << 3)) & 0xff;
      }
    }
    p = e;
  }
}

function setupDecryption(r: BiffRecord, biff: number, password: string | undefined): { rc4: BlockCipher | null; xor: Uint8Array | null; kind: XlsBook["encrypted"] } {
  const c = new Cursor(r.data);
  const tryPw = <T>(f: (pw: string) => T | null): T => {
    const viaDefault = f(DEFAULT_PASSWORD);
    if (viaDefault) return viaDefault;
    if (password == null) throw new XlsPasswordError("password-required");
    const hit = f(password);
    if (!hit) throw new XlsPasswordError("password-wrong");
    return hit;
  };
  const xorCheck = (key: number, verifier: number) => tryPw((pw) => {
    const x = xorKeyAndArray(pw);
    return x.key === key && x.verifier === verifier ? x.array : null;
  });
  if (biff < 8 || r.data.length === 4) {
    const key = c.u16(), ver = c.u16();
    return { rc4: null, xor: xorCheck(key, ver), kind: "xor" };
  }
  const type = c.u16();
  if (type === 0) {
    const key = c.u16(), ver = c.u16();
    return { rc4: null, xor: xorCheck(key, ver), kind: "xor" };
  }
  const major = c.u16();
  c.u16(); // minor
  if (major === 1) {
    const salt = c.bytes(16).slice(), ver = c.bytes(16).slice(), hash = c.bytes(16).slice();
    return { rc4: tryPw((pw) => rc4Standard(pw, salt, ver, hash)), xor: null, kind: "rc4" };
  }
  // RC4 CryptoAPI: EncryptionHeader then EncryptionVerifier.
  c.u32(); // flags
  const hdrSize = c.u32();
  const hdr = new Cursor(c.bytes(hdrSize));
  hdr.u32(); hdr.u32(); hdr.u32(); // flags, sizeExtra, algId
  hdr.u32(); // algIdHash
  const keyBits = hdr.u32();
  const saltSize = c.u32();
  const salt = c.bytes(saltSize).slice(), ver = c.bytes(16).slice();
  c.u32(); // verifier hash size
  const hash = c.bytes(Math.min(20, c.left)).slice();
  return { rc4: tryPw((pw) => rc4CryptoApi(pw, salt, ver, hash, keyBits)), xor: null, kind: "cryptoapi" };
}

function parseWorkbookStream(input: Uint8Array, opts: XlsReadOptions): XlsBook {
  let stream = input;
  const first = records(stream).next().value as BiffRecord | undefined;
  if (!first || !BOF_TYPES.has(first.type)) throw new Error("damaged .xls: the Workbook stream does not start with BOF");
  let biff = 8;
  const bofVersion = (r: BiffRecord): number => {
    if (r.type === 0x0009) return 2;
    if (r.type === 0x0209) return 3;
    if (r.type === 0x0409) return 4;
    const v = r.data.length >= 2 ? r.data[0]! | (r.data[1]! << 8) : 0x600;
    return v === 0x500 ? 5 : v === 0x400 ? 4 : v === 0x300 ? 3 : v === 0x200 || v === 0x7 ? 2 : 8;
  };
  biff = bofVersion(first);
  const bofDt = first.data.length >= 4 ? first.data[2]! | (first.data[3]! << 8) : 0x10;
  const book: XlsBook = {
    biff, codepage: biff >= 8 ? 1200 : 1252, date1904: false, fonts: [], formats: new Map(), xfs: [], palette: DEFAULT_PALETTE.slice(),
    sheets: [], names: [], supbooks: [], xti: [], activeTab: 0, sst: [], skippedSheets: [],
  };
  // BIFF2-4 (and a BIFF5 file that is a single worksheet) hold one sheet in the stream itself.
  const single = biff < 5 || bofDt === 0x10;
  const bound: { name: string; pos: number; state: number; type: number }[] = [];
  const names5: { sheet: string; names: string[] }[] = []; // BIFF5 EXTERNSHEET entries and their EXTERNNAMEs
  let dggParts: Uint8Array[] = [];
  let globalsEnd = 0;
  for (let it = records(stream, 0), step = it.next(); !step.done; step = it.next()) {
    const r = step.value;
    const c = new Cursor(r, book.codepage);
    try {
      switch (r.type) {
        case 0x002f: { // FILEPASS
          const dec = setupDecryption(r, biff, opts.password);
          book.encrypted = dec.kind;
          stream = stream.slice();
          const from = r.pos + 4 + r.data.length;
          decryptStream(stream, from, dec.rc4, dec.xor);
          it = records(stream, from);
          break;
        }
        case 0x0042: book.codepage = c.u16(); if (book.codepage === 0x8000) book.codepage = 10000; if (biff >= 8) book.codepage = 1200; break;
        case 0x0022: book.date1904 = c.u16() === 1; break;
        case 0x0031: case 0x0231: book.fonts.push(readFont(c, biff, r.type)); break;
        case 0x0045: { const f = book.fonts[book.fonts.length - 1]; if (f) f.color = c.u16(); break; } // BIFF2 FONTCOLOR
        case 0x001e: case 0x041e: {
          if (biff >= 5 || r.type === 0x041e) {
            const id = c.u16();
            book.formats.set(id, readString(c, biff, biff >= 8 ? 2 : 1));
          } else {
            if (biff === 4) c.u16();
            book.formats.set(book.formats.size, readString(c, biff, 1));
          }
          break;
        }
        case 0x00e0: case 0x0243: case 0x0443: case 0x0043: book.xfs.push(readXf(c, biff, r.type)); break;
        case 0x0092: {
          const n = c.u16();
          for (let i = 0; i < n && i < 56; i++) {
            const [rr, g, b] = [c.u8(), c.u8(), c.u8()];
            c.u8();
            book.palette[i] = "#" + [rr, g, b].map((v) => v.toString(16).padStart(2, "0")).join("");
          }
          break;
        }
        case 0x0085: { // BOUNDSHEET
          const pos = c.u32();
          const state = c.u8() & 3;
          const type = c.u8();
          const name = readString(c, biff, 1);
          bound.push({ name, pos, state, type });
          break;
        }
        case 0x01ae: book.supbooks.push(readSupbook(c, r)); break;
        case 0x0059: case 0x005a: { // XCT / CRN: an external book's cached values
          const sb = book.supbooks[book.supbooks.length - 1];
          if (sb?.kind === "external") (sb.cache ??= []).push({ type: r.type, data: r.data.slice() });
          break;
        }
        case 0x0023: { // EXTERNNAME
          const sb = book.supbooks[book.supbooks.length - 1];
          if (biff >= 8) {
            c.u16(); c.u32();
            const nm = readString(c, 8, 1);
            if (sb) { sb.names.push(nm); (sb.rawNames ??= []).push(concatAll([r.data, ...r.conts])); }
          } else {
            c.u16(); c.u32();
            names5[names5.length - 1]?.names.push(readString(c, biff, 1));
          }
          break;
        }
        case 0x0017: { // EXTERNSHEET
          if (biff >= 8) {
            const n = c.u16();
            for (let i = 0; i < n && c.left >= 6; i++) book.xti.push({ book: c.u16(), first: c.i16(), last: c.i16() });
          } else {
            const cch = c.u8();
            const raw = c.bytes(Math.min(cch, c.left));
            names5.push({ sheet: new TextDecoder("windows-1252").decode(raw), names: [] });
          }
          break;
        }
        case 0x0018: case 0x0218: book.names.push(readName(c, biff)); break;
        case 0x00fc: book.sst = readSst(r); break;
        case 0x003d: if (r.data.length >= 12) { c.skip(10); book.activeTab = c.u16(); } break;
        case 0x0012: book.protect = c.u16() === 1; break;
        case 0x0013: book.password = c.u16(); break;
        case 0x0019: book.windowProtect = c.u16() === 1; break;
        case 0x00eb: dggParts.push(r.data, ...r.conts); break;
        case 0x01ba: if (biff >= 8) book.codeName = readString(c, 8, 2); break; // CODENAME
      }
    } catch (e) {
      if (e instanceof XlsPasswordError) throw e;
      // A truncated record is skipped; the rest of the stream may still be readable.
    }
    if (r.type === 0x000a) { globalsEnd = r.pos + 4 + r.data.length; break; }
    if (BOF_TYPES.has(r.type) && r.pos > 0 && !single) { globalsEnd = r.pos; break; } // globals with no EOF
  }
  // No BOUNDSHEET at all: find the worksheets that follow the globals, as LibreOffice does.
  if (!single && !bound.length && globalsEnd > 0) {
    let depth = 0;
    for (const r of records(stream, globalsEnd)) {
      if (BOF_TYPES.has(r.type)) {
        const dt = r.data.length >= 4 ? r.data[2]! | (r.data[3]! << 8) : 0x10;
        if (depth === 0 && (dt === 0x10 || dt === 0x40)) bound.push({ name: `Sheet${bound.length + 1}`, pos: r.pos, state: 0, type: 0 });
        depth++;
      } else if (r.type === 0x000a) depth = Math.max(0, depth - 1);
    }
  }
  const blips = dggParts.length ? readBlipStore(concatAll(dggParts)) : [];
  dggParts = [];
  // Sheet names for formulas follow BOUNDSHEET order, charts included.
  const allNames = bound.map((b) => b.name);
  const ctxBase = makeContext(book, allNames, names5);
  if (single) {
    const sh = newSheet("Sheet1", 0);
    readSheet(stream, 0, sh, book, ctxBase, blips, true);
    book.sheets.push(sh);
  } else {
    for (const [bi, b] of bound.entries()) {
      if (b.type === 2 && biff >= 8 && b.pos > 0 && b.pos < stream.length) {
        // A chart sheet: one chart filling the sheet.
        const recs = chartRecords(stream, b.pos);
        const chart = readChart(recs, chartEnv(book, ctxBase));
        if (chart) {
          const sh = newSheet(b.name, b.state);
          sh.boundIndex = bi;
          sh.chartSheet = chart;
          sh.chartDrawings = recs.filter((r) => r.type === 0x005d).length;
          const cn = recs.find((r) => r.type === 0x01ba);
          if (cn) try { sh.codeName = readString(new Cursor(cn, book.codepage), 8, 2); } catch { /* no code name */ }
          book.sheets.push(sh);
          continue;
        }
      }
      if (b.type !== 0 && b.type !== 1) { book.skippedSheets.push(b.name); continue; }
      if (b.pos <= 0 || b.pos >= stream.length) throw new Error(`damaged .xls: sheet "${b.name}" points outside the stream`);
      const sh = newSheet(b.name, b.state);
      sh.boundIndex = bi;
      readSheet(stream, b.pos, sh, book, ctxBase, blips, false);
      book.sheets.push(sh);
    }
  }
  if (!book.sheets.length) throw new Error(book.skippedSheets.length ? "this .xls holds only chart sheets, which sheetedit cannot open" : "damaged .xls: no worksheet found");
  for (const n of book.names) {
    if (!n.rgce || biff < 5) continue;
    n.formula = decodeFormula(n.rgce, n.rgcb, { ...ctxBase, row: 0, col: 0 }) ?? undefined;
    n.rgce = n.rgcb = undefined;
    // The autofilter range is the sheet's hidden _FilterDatabase name.
    if (n.builtin === 0x0d && n.itab && n.formula) {
      const sh = book.sheets.find((s) => s.boundIndex === n.itab - 1);
      const g = parseArea(n.formula);
      if (sh?.hasAutoFilter && g) sh.autoFilter = g;
    }
  }
  // The active tab counts skipped sheets; map it onto the loaded ones.
  const loadedBefore = single ? 0 : book.sheets.filter((s) => (s.boundIndex ?? 0) <= book.activeTab).length - 1;
  book.activeTab = Math.max(0, Math.min(book.sheets.length - 1, loadedBefore));
  return book;
}

const concatAll = (parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
};

const quoteSheet = (name: string): string =>
  /^[A-Za-z_¡-￿][A-Za-z0-9_.¡-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name) ? name : "'" + name.replace(/'/g, "''") + "'";

/** Formula context shared by the sheets, minus the cell position. */
function makeContext(book: XlsBook, sheetNames: string[], names5: { sheet: string; names: string[] }[]): Omit<FormulaContext, "row" | "col"> {
  const sheetRange = (a: string, b: string): string => (a === b ? quoteSheet(a) : quoteSheet(a + ":" + b).replace(/^'(.*)'$/, (m) => (/[^A-Za-z0-9_.:]/.test(a + b) ? m : a + ":" + b)));
  return {
    biff: book.biff,
    xti: (i) => {
      const x = book.xti[i];
      if (!x) return null;
      const sb = book.supbooks[x.book];
      if (!sb) return null;
      if (sb.kind === "self") {
        if (x.first < 0 || x.first >= sheetNames.length) return null;
        return sheetRange(sheetNames[x.first]!, sheetNames[Math.max(x.first, Math.min(x.last, sheetNames.length - 1))]!) + "!";
      }
      if (sb.kind === "external") {
        const ext = book.supbooks.filter((s) => s.kind === "external").indexOf(sb) + 1;
        const s1 = sb.sheets[x.first], s2 = sb.sheets[x.last];
        if (s1 == null) return "[" + ext + "]!";
        return "'[" + ext + "]" + (s2 && s2 !== s1 ? s1 + ":" + s2 : s1).replace(/'/g, "''") + "'!";
      }
      return null;
    },
    tabs: (t1, t2, ixals) => {
      if (ixals > 0) {
        const e = names5[ixals - 1]?.sheet;
        if (e && (e.charCodeAt(0) === 3 || e.charCodeAt(0) === 2)) return quoteSheet(e.slice(1)) + "!";
      }
      if (t1 < 0 || t1 >= sheetNames.length) return ixals < 0 && t1 < 0 ? "" : null;
      return sheetRange(sheetNames[t1]!, sheetNames[Math.max(t1, Math.min(t2, sheetNames.length - 1))]!) + "!";
    },
    name: (i) => {
      return (book.names[i - 1]?.name ?? "#NAME?").replace(/^_xlfn\./, "");
    },
    nameX: (ixti, i) => {
      if (book.biff < 8) return names5[Math.abs(ixti) - 1]?.names[i - 1] ?? "#NAME?";
      const x = book.xti[ixti];
      const sb = x && book.supbooks[x.book];
      if (!sb) return "#NAME?";
      if (sb.kind === "self") {
        const n = book.names[i - 1];
        if (!n) return "#NAME?";
        // A name scoped to a sheet, referenced from elsewhere, carries its sheet.
        if (n.itab && n.builtin == null && sheetNames[n.itab - 1] != null) return quoteSheet(sheetNames[n.itab - 1]!) + "!" + n.name;
        return n.name.replace(/^_xlfn\./, "");
      }
      const nm = (sb.names[i - 1] ?? "#NAME?").replace(/^_xlfn\./, "");
      if (sb.kind === "external") {
        const ext = book.supbooks.filter((s) => s.kind === "external").indexOf(sb) + 1;
        return "[" + ext + "]!" + nm;
      }
      return nm;
    },
  };
}

function readFont(c: Cursor, biff: number, type: number): XlsFont {
  const height = c.u16();
  const grbit = c.u16();
  if (biff <= 4 || type === 0x0031 && biff === 2) {
    let color = 0x7fff;
    if (type === 0x0231 || biff >= 3) color = c.u16();
    return { height, italic: !!(grbit & 2), strike: !!(grbit & 8), weight: grbit & 1 ? 700 : 400, underline: grbit & 4 ? 1 : 0, escapement: 0, color, name: readString(c, biff, 1) };
  }
  const color = c.u16();
  const weight = c.u16();
  const escapement = c.u16();
  const underline = c.u8();
  c.u8(); c.u8(); c.u8(); // family, charset, reserved
  return { height, italic: !!(grbit & 2), strike: !!(grbit & 8), weight, underline, escapement, color, name: readString(c, biff, 1) };
}

const NO_BORDER = { style: 0, color: 64 };

function readXf(c: Cursor, biff: number, type: number): XlsXf {
  const xf: XlsXf = {
    font: 0, fmt: 0, locked: true, hidden: false, isStyle: false, parent: 0, hAlign: 0, wrap: false, vAlign: 2, rot: 0, indent: 0, shrink: false,
    left: { ...NO_BORDER }, right: { ...NO_BORDER }, top: { ...NO_BORDER }, bottom: { ...NO_BORDER }, pattern: 0, fore: 64, back: 65,
  };
  if (biff >= 8) {
    xf.font = c.u16(); xf.fmt = c.u16();
    const f = c.u16();
    xf.locked = !!(f & 1); xf.hidden = !!(f & 2); xf.isStyle = !!(f & 4); xf.parent = f >> 4;
    const al = c.u8();
    xf.hAlign = al & 7; xf.wrap = !!(al & 8); xf.vAlign = (al >> 4) & 7;
    xf.rot = c.u8();
    const ind = c.u8();
    xf.indent = ind & 15; xf.shrink = !!(ind & 16);
    c.u8();
    const b1 = c.u32(), b2 = c.u32(), p = c.u16();
    xf.left = { style: b1 & 15, color: (b1 >>> 16) & 0x7f };
    xf.right = { style: (b1 >>> 4) & 15, color: (b1 >>> 23) & 0x7f };
    xf.top = { style: (b1 >>> 8) & 15, color: b2 & 0x7f };
    xf.bottom = { style: (b1 >>> 12) & 15, color: (b2 >>> 7) & 0x7f };
    const diag = (b1 >>> 30) & 3, dgDiag = (b2 >>> 21) & 15;
    if (diag && dgDiag) xf.diag = { style: dgDiag, color: (b2 >>> 14) & 0x7f, down: !!(diag & 1), up: !!(diag & 2) };
    xf.pattern = (b2 >>> 26) & 0x3f;
    xf.fore = p & 0x7f; xf.back = (p >> 7) & 0x7f;
  } else if (biff >= 5) {
    xf.font = c.u16(); xf.fmt = c.u16();
    const f = c.u16();
    xf.locked = !!(f & 1); xf.hidden = !!(f & 2); xf.isStyle = !!(f & 4); xf.parent = f >> 4;
    const al = c.u8();
    xf.hAlign = al & 7; xf.wrap = !!(al & 8); xf.vAlign = (al >> 4) & 7;
    const ori = c.u8() & 3;
    xf.rot = ori === 1 ? 255 : ori === 2 ? 90 : ori === 3 ? 180 : 0;
    const a = c.u32(), b = c.u32();
    xf.fore = a & 0x7f; xf.back = (a >>> 7) & 0x7f; xf.pattern = (a >>> 16) & 0x3f;
    xf.bottom = { style: (a >>> 22) & 7, color: (a >>> 25) & 0x7f };
    xf.top = { style: b & 7, color: (b >>> 9) & 0x7f };
    xf.left = { style: (b >>> 3) & 7, color: (b >>> 16) & 0x7f };
    xf.right = { style: (b >>> 6) & 7, color: (b >>> 23) & 0x7f };
  } else if (type === 0x0043 && biff === 2) {
    xf.font = c.u8(); c.u8();
    const f = c.u8();
    xf.fmt = f & 0x3f; xf.locked = !!(f & 0x40); xf.hidden = !!(f & 0x80);
    const al = c.u8();
    xf.hAlign = al & 7;
  } else {
    // BIFF3 / BIFF4: font, format, protection/type, alignment, area, borders.
    xf.font = c.u8(); xf.fmt = c.u8();
    const f = c.u8();
    xf.locked = !!(f & 1); xf.hidden = !!(f & 2); xf.isStyle = !!(f & 4);
    if (biff === 4) {
      c.u8();
      const al = c.u8();
      xf.hAlign = al & 7; xf.wrap = !!(al & 8); xf.vAlign = (al >> 4) & 3;
      c.u8();
    } else {
      c.u8();
      const al = c.u16();
      xf.hAlign = al & 7; xf.wrap = !!(al & 8);
    }
    const area = c.u16();
    xf.pattern = area & 0x3f; xf.fore = (area >> 6) & 0x1f; xf.back = (area >> 11) & 0x1f;
    const bd = c.u32();
    xf.top = { style: bd & 7, color: (bd >>> 3) & 0x1f };
    xf.left = { style: (bd >>> 8) & 7, color: (bd >>> 11) & 0x1f };
    xf.bottom = { style: (bd >>> 16) & 7, color: (bd >>> 19) & 0x1f };
    xf.right = { style: (bd >>> 24) & 7, color: (bd >>> 27) & 0x1f };
  }
  return xf;
}

function readSupbook(c: Cursor, r: BiffRecord): XlsSupbook {
  const ctab = c.u16();
  const cch = c.u16();
  if (cch === 0x0401) return { kind: "self", sheets: [], names: [] };
  if (cch === 0x3a01) return { kind: "addin", sheets: [], names: [] };
  void r;
  let path = readStringNoCch(c, cch, 8);
  const sheets: string[] = [];
  for (let i = 0; i < ctab && c.left >= 3; i++) sheets.push(readString(c, 8, 2));
  // Encoded virtual path: a leading 0x01 starts the encoding; strip the control characters.
  if (path.charCodeAt(0) === 1) path = path.slice(1).replace(/[\u0001-\u0008]/g, (ch) => (ch === "\u0003" ? "/" : ch === "\u0002" ? "/" : ""));
  return { kind: path === "\u0000" || !path ? "ole" : "external", path, sheets, names: [] };
}

function readName(c: Cursor, biff: number): XlsName {
  const grbit = c.u16();
  c.u8(); // keyboard shortcut
  const cch = c.u8();
  const cce = c.u16();
  let itab = 0;
  if (biff >= 5) {
    c.u16();
    itab = c.u16();
    c.skip(4);
  }
  const builtin = (grbit & 0x20) !== 0;
  let name: string;
  let code: number | undefined;
  if (biff >= 8) {
    const flags = c.u8();
    name = c.chars8(cch, (flags & 1) === 1);
  } else name = c.chars7(cch);
  if (builtin && name.length >= 1) {
    code = name.charCodeAt(0);
    const BUILTIN = ["Consolidate_Area", "Auto_Open", "Auto_Close", "Extract", "Database", "Criteria", "Print_Area", "Print_Titles", "Recorder", "Data_Form", "Auto_Activate", "Auto_Deactivate", "Sheet_Title", "_FilterDatabase"];
    name = BUILTIN[code] ?? name;
  }
  const rgce = c.bytes(Math.min(cce, c.left)).slice();
  const rgcb = c.left > 0 ? c.bytes(c.left).slice() : undefined;
  return { name, itab, hidden: (grbit & 1) !== 0, func: (grbit & 2) !== 0, builtin: code, rgce, rgcb };
}

function readSst(r: BiffRecord): XlsString[] {
  const c = new Cursor(r);
  c.u32();
  const unique = c.u32();
  const out: XlsString[] = [];
  for (let i = 0; i < unique && c.left > 0; i++) {
    try { out.push(readRichString(c)); } catch { break; }
  }
  return out;
}

const rk = (v: number): number => {
  let n: number;
  if (v & 2) n = v >> 2;
  else {
    const b = new DataView(new ArrayBuffer(8));
    b.setUint32(4, v & 0xfffffffc, true);
    n = b.getFloat64(0, true);
  }
  return v & 1 ? n / 100 : n;
};

const range8 = (c: Cursor): XlsRange => {
  const r1 = c.u16(), r2 = c.u16(), c1 = c.u16(), c2 = c.u16();
  return { r1, r2, c1: c1 & 0xff, c2: Math.min(c2, 0xff) };
};

/** "Sheet1!$A$1:$C$9" -> a 0-based range. */
function parseArea(f: string): XlsRange | null {
  const m = /(?:^|!)\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(f);
  if (!m) return null;
  const col = (l: string): number => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const r1 = Number(m[2]) - 1, c1 = col(m[1]!);
  return { r1, c1, r2: m[4] ? Number(m[4]) - 1 : r1, c2: m[3] ? col(m[3]) : c1 };
}

const a1 = (r: number, c: number): string => colToLetters(c + 1) + String(r + 1);
const rangeA1 = (g: XlsRange): string => (g.r1 === g.r2 && g.c1 === g.c2 ? a1(g.r1, g.c1) : a1(g.r1, g.c1) + ":" + a1(g.r2, g.c2));
export { rangeA1 };

/** Read one sheet substream from `pos` (the sheet's BOF). */
function readSheet(stream: Uint8Array, pos: number, sh: XlsSheet, book: XlsBook, base: Omit<FormulaContext, "row" | "col">, blips: (Blip | undefined)[], single: boolean): void {
  const biff = book.biff;
  const b8 = biff >= 8;
  const ctx = (row: number, col: number): FormulaContext => ({ ...base, row, col });
  const shared: { range: XlsRange; rgce: Uint8Array; rgcb?: Uint8Array }[] = [];
  const arrays: { range: XlsRange; rgce: Uint8Array; rgcb?: Uint8Array }[] = [];
  const pendingExp: XlsCell[] = [];
  let lastFormula: XlsCell | null = null;
  let depth = 0;
  let started = false;
  const drawing: Uint8Array[] = [];
  const objs: { ot: number; id: number }[] = [];
  const txos: XlsTxo[] = [];
  const txoByObj = new Map<number, string>();
  const chartAt = new Map<number, number>();
  let cf: { ranges: XlsRange[]; rules: XlsCf[] } | null = null;
  let note5: { r: number; c: number; text: string } | null = null;
  const put = (cell: XlsCell): void => {
    if (cell.r > 0xffff || cell.c > 0xff) return;
    sh.cells.set(cell.r * 256 + cell.c, cell);
  };
  const cellHead = (c: Cursor): [number, number, number] => {
    const r = c.u16(), col = c.u16();
    if (biff === 2 && !b8) {
      // BIFF2 cell attributes: 3 bytes; the XF index lives in the low 6 bits of the first.
      const a = c.u8(); c.u8(); c.u8();
      return [r, col, a & 0x3f];
    }
    return [r, col, c.u16()];
  };
  let count = 0;
  for (const r of records(stream, pos)) {
    if (++count > 5_000_000) throw new Error("damaged .xls: too many records");
    if (BOF_TYPES.has(r.type)) {
      if (started) {
        depth++;
        // An embedded chart's substream follows its OBJ.
        if (depth === 1 && b8 && r.data.length >= 4 && (r.data[2]! | (r.data[3]! << 8)) === 0x20) chartAt.set(objs.length - 1, r.pos);
      }
      started = true;
      continue;
    }
    if (r.type === 0x000a) {
      if (depth === 0) break;
      // The next shape's drawing data can follow an embedded chart's EOF as CONTINUEs.
      if (depth === 1 && b8) drawing.push(...r.conts);
      depth--;
      continue;
    }
    if (depth > 0) { if (r.type === 0x005d) sh.chartDrawings++; continue; } // embedded chart substream
    const c = new Cursor(r, book.codepage);
    try {
      switch (r.type) {
        case 0x00fd: { // LABELSST
          const [row, col, xf] = cellHead(c);
          const s = book.sst[c.u32()] ?? { text: "" };
          put({ r: row, c: col, xf, val: { k: "s", v: s.text, runs: s.runs } });
          break;
        }
        case 0x0204: case 0x0004: case 0x00d6: { // LABEL / RSTRING
          const [row, col, xf] = cellHead(c);
          let text: string;
          if (b8) text = readString(c, 8, 2);
          else text = r.type === 0x0004 && biff === 2 ? c.chars7(c.u8()) : c.chars7(c.u16());
          put({ r: row, c: col, xf, val: { k: "s", v: text } });
          break;
        }
        case 0x0203: case 0x0003: { const [row, col, xf] = cellHead(c); put({ r: row, c: col, xf, val: { k: "n", v: c.f64() } }); break; }
        case 0x0002: { const [row, col, xf] = cellHead(c); put({ r: row, c: col, xf, val: { k: "n", v: c.u16() } }); break; }
        case 0x027e: case 0x007e: { const [row, col, xf] = cellHead(c); put({ r: row, c: col, xf, val: { k: "n", v: rk(c.u32()) } }); break; }
        case 0x00bd: { // MULRK
          const row = c.u16();
          let col = c.u16();
          while (c.left >= 8) { const xf = c.u16(); put({ r: row, c: col++, xf, val: { k: "n", v: rk(c.u32()) } }); }
          break;
        }
        case 0x0205: case 0x0005: {
          const [row, col, xf] = cellHead(c);
          const v = c.u8(), isErr = c.u8();
          put({ r: row, c: col, xf, val: isErr ? { k: "e", v: ERRORS[v] ?? "#N/A" } : { k: "b", v: v !== 0 } });
          break;
        }
        case 0x0201: case 0x0001: { const [row, col, xf] = cellHead(c); put({ r: row, c: col, xf, val: { k: "blank" } }); break; }
        case 0x00be: {
          const row = c.u16();
          let col = c.u16();
          while (c.left >= 4) { const xf = c.u16(); put({ r: row, c: col++, xf, val: { k: "blank" } }); }
          break;
        }
        case 0x0006: case 0x0206: case 0x0406: { // FORMULA
          const [row, col, xf] = cellHead(c);
          const res = c.bytes(8).slice();
          let val: XlsCell["val"];
          let wantString = false;
          if (res[6] === 0xff && res[7] === 0xff) {
            if (res[0] === 0) { val = { k: "s", v: "" }; wantString = true; }
            else if (res[0] === 1) val = { k: "b", v: res[2] !== 0 };
            else if (res[0] === 2) val = { k: "e", v: ERRORS[res[2]!] ?? "#N/A" };
            else val = { k: "s", v: "" };
          } else val = { k: "n", v: new DataView(res.buffer).getFloat64(0, true) };
          if (biff === 2) c.skip(1); else c.u16(); // grbit
          if (biff >= 5) c.u32(); // chn
          const cce = biff === 2 ? c.u8() : c.u16();
          const rgce = c.bytes(Math.min(cce, c.left)).slice();
          const rgcb = c.left > 0 ? c.bytes(c.left).slice() : undefined;
          const cell: XlsCell = { r: row, c: col, xf, val };
          if (rgce[0] === 0x01 && rgce.length >= 5) {
            cell.exp = { r: rgce[1]! | (rgce[2]! << 8), c: biff >= 3 && rgce.length >= 5 ? rgce[3]! | (rgce[4]! << 8) : rgce[3]! };
            pendingExp.push(cell);
          } else if (rgce[0] === 0x02) {
            // Data table: keep the value only.
          } else if (biff >= 3) {
            const f = decodeFormula(rgce, rgcb, ctx(row, col));
            if (f != null) cell.formula = f;
          }
          put(cell);
          lastFormula = wantString ? cell : null;
          break;
        }
        case 0x0207: case 0x0007: { // STRING (a formula's text result)
          if (lastFormula) {
            const text = b8 ? readString(c, 8, 2) : biff === 2 ? c.chars7(c.u8()) : c.chars7(c.u16());
            lastFormula.val = { k: "s", v: text };
            lastFormula = null;
          }
          break;
        }
        case 0x04bc: { // SHRFMLA
          const g = { r1: c.u16(), r2: c.u16(), c1: c.u8(), c2: c.u8() };
          c.u16();
          const cce = c.u16();
          const rgce = c.bytes(Math.min(cce, c.left)).slice();
          shared.push({ range: g, rgce, rgcb: c.left ? c.bytes(c.left).slice() : undefined });
          break;
        }
        case 0x0221: case 0x0021: { // ARRAY
          const g = { r1: c.u16(), r2: c.u16(), c1: c.u8(), c2: c.u8() };
          c.u16();
          if (biff >= 5) c.u32();
          const cce = biff === 2 ? c.u8() : c.u16();
          const rgce = c.bytes(Math.min(cce, c.left)).slice();
          arrays.push({ range: g, rgce, rgcb: c.left ? c.bytes(c.left).slice() : undefined });
          break;
        }
        case 0x0208: case 0x0008: { // ROW
          const row = c.u16();
          c.u16(); c.u16();
          const h = c.u16();
          const info = sh.rows.get(row) ?? {};
          if (biff === 2) { if (!(h & 0x8000)) { info.height = h & 0x7fff; info.custom = true; } sh.rows.set(row, info); break; }
          c.u16(); c.u16();
          const f = c.u16();
          info.height = h & 0x7fff;
          info.level = f & 7;
          info.collapsed = !!(f & 0x10);
          info.hidden = !!(f & 0x20);
          info.custom = !!(f & 0x40);
          if (f & 0x80) info.xf = c.u16() & 0xfff;
          sh.rows.set(row, info);
          break;
        }
        case 0x007d: { // COLINFO
          const first = c.u16(), last = c.u16(), width = c.u16(), xf = c.u16(), f = c.u16();
          if (first > 0xff) break;
          sh.cols.push({ first, last: Math.min(last, 0xff), width, xf, hidden: !!(f & 1), level: (f >> 8) & 7, collapsed: !!(f & 0x1000) });
          break;
        }
        case 0x0024: { // BIFF2 COLWIDTH
          const first = c.u8(), last = c.u8(), width = c.u16();
          sh.cols.push({ first, last, width, xf: 15, hidden: false, level: 0, collapsed: false });
          break;
        }
        case 0x0055: sh.defColWidth = c.u16(); break;
        case 0x0099: sh.stdWidth = c.u16(); break;
        case 0x0225: c.u16(); sh.defRowHeight = c.u16() & 0x7fff; break;
        case 0x0025: sh.defRowHeight = c.u16() & 0x7fff; break;
        case 0x0081: { const f = c.u16(); sh.summaryBelow = !!(f & 0x40); sh.summaryRight = !!(f & 0x80); sh.print.fitToPage = !!(f & 0x100); break; }
        case 0x00e5: { const n = c.u16(); for (let i = 0; i < n && c.left >= 8; i++) sh.merges.push(range8(c)); break; }
        case 0x01b8: { const l = readHyperlink(c); if (l) sh.links.push(l); break; }
        case 0x0800: { // HLINKTOOLTIP
          c.u16();
          const g = range8(c);
          const tip = utf16le(c.bytes(c.left)).replace(/\u0000.*$/, "");
          const l = sh.links.find((x) => x.range.r1 === g.r1 && x.range.c1 === g.c1);
          if (l && tip) l.tip = tip;
          break;
        }
        case 0x001c: { // NOTE
          const row = c.u16(), col = c.u16();
          if (b8) {
            c.u16();
            const obj = c.u16();
            const author = readString(c, 8, 2);
            sh.notes.push({ r: row, c: col, author, text: "", obj });
          } else if (row !== 0xffff) {
            const cch = c.u16();
            note5 = { r: row, c: col, text: c.chars7(Math.min(cch, c.left)) };
            sh.notes.push({ r: row, c: col, author: "", text: note5.text, obj: -1 });
          } else if (note5) {
            const cch = c.u16();
            note5.text += c.chars7(Math.min(cch, c.left));
            sh.notes[sh.notes.length - 1]!.text = note5.text;
          }
          break;
        }
        case 0x005d: { // OBJ
          if (!b8) { sh.shapes++; break; }
          let ot = -1, id = -1;
          while (c.left >= 4) {
            const ft = c.u16(), cb = c.u16();
            if (ft === 0x15) { ot = c.u16(); id = c.u16(); c.skip(Math.max(0, cb - 4)); break; }
            if (ft === 0) break;
            c.skip(Math.min(cb, c.left));
          }
          objs.push({ ot, id });
          drawing.push(...r.conts); // the next shapes' drawing data can ride in its CONTINUEs
          break;
        }
        case 0x01b6: { // TXO
          const txo = readTxo(r);
          if (b8) drawing.push(...r.conts.slice(txoConts(r)));
          txos.push(txo);
          const o = objs[objs.length - 1];
          if (o) txoByObj.set(o.id, txo.text);
          break;
        }
        case 0x00ec: drawing.push(r.data, ...r.conts); break;
        case 0x023e: case 0x003e: {
          if (biff >= 5) sh.window2 = c.u16();
          break;
        }
        case 0x0041: { const x = c.u16(), y = c.u16(), top = c.u16(), left = c.u16(); sh.pane = { x, y, top, left }; break; }
        case 0x0012: sh.protect = c.u16() === 1; break;
        case 0x0013: sh.password = c.u16(); break;
        case 0x0063: sh.objProtect = c.u16() === 1; break;
        case 0x00dd: sh.scenProtect = c.u16() === 1; break;
        case 0x0867: { // FEATHEADR
          c.skip(12);
          const isf = c.u16();
          c.u8();
          const cb = c.u32();
          if (isf === 2 && cb !== 0xffffffff && c.left >= 4) sh.allow = c.u32();
          break;
        }
        case 0x0862: { // SHEETEXT
          c.skip(12);
          c.u32();
          const icv = c.u32() & 0x7f;
          if (icv !== 0x7f) sh.tabColor = icv;
          break;
        }
        case 0x0014: { const s = b8 ? (r.data.length ? readString(c, 8, 2) : "") : r.data.length ? c.chars7(c.u8()) : ""; if (s) sh.print.header = s; break; }
        case 0x0015: { const s = b8 ? (r.data.length ? readString(c, 8, 2) : "") : r.data.length ? c.chars7(c.u8()) : ""; if (s) sh.print.footer = s; break; }
        case 0x0026: sh.print.left = c.f64(); break;
        case 0x0027: sh.print.right = c.f64(); break;
        case 0x0028: sh.print.top = c.f64(); break;
        case 0x0029: sh.print.bottom = c.f64(); break;
        case 0x002a: sh.print.headings = c.u16() === 1; break;
        case 0x002b: sh.print.gridLines = c.u16() === 1; break;
        case 0x0083: sh.print.hCenter = c.u16() === 1; break;
        case 0x0084: sh.print.vCenter = c.u16() === 1; break;
        case 0x00a1: {
          const p = sh.print;
          p.paperSize = c.u16(); p.scale = c.u16(); p.pageStart = c.i16(); p.fitWidth = c.u16(); p.fitHeight = c.u16();
          const g = c.u16();
          p.overThenDown = !!(g & 1); p.portrait = !!(g & 2); p.noPls = !!(g & 4); p.noOrient = !!(g & 0x40); p.usePage = !!(g & 0x80);
          if (biff >= 5 && c.left >= 20) { c.u16(); c.u16(); p.headerMargin = c.f64(); p.footerMargin = c.f64(); }
          break;
        }
        case 0x001b: case 0x001a: {
          const n = c.u16();
          const out: number[] = [];
          for (let i = 0; i < n && c.left >= 2; i++) { out.push(c.u16()); if (b8) c.skip(4); }
          if (r.type === 0x001b) sh.print.rowBreaks = out; else sh.print.colBreaks = out;
          break;
        }
        case 0x01b0: { // CONDFMT
          c.u16(); c.u16();
          range8(c);
          const n = c.u16();
          const ranges: XlsRange[] = [];
          for (let i = 0; i < n && c.left >= 8; i++) ranges.push(range8(c));
          cf = { ranges, rules: [] };
          sh.condFmts.push(cf);
          break;
        }
        case 0x01b1: { // CF
          if (!cf) break;
          const top = cf.ranges[0] ?? { r1: 0, c1: 0 };
          const rule = readCf(c, ctx(top.r1, top.c1));
          if (rule) cf.rules.push(rule);
          break;
        }
        case 0x009d: sh.hasAutoFilter = true; break;
        case 0x01ba: if (b8) sh.codeName = readString(c, 8, 2); break; // CODENAME
        case 0x01be: { // DV
          const dv = readDv(c, ctx);
          if (dv) sh.dvs.push(dv);
          break;
        }
      }
    } catch {
      // A truncated record is skipped.
    }
  }
  if (!single && !started) throw new Error(`damaged .xls: sheet "${sh.name}" has no BOF`);
  // Shared and array formulas, now that every SHRFMLA / ARRAY record is known.
  const findIn = (list: typeof shared, r: number, c: number, anchorR: number, anchorC: number) =>
    list.find((s) => s.range.r1 === anchorR && s.range.c1 === anchorC) ?? list.find((s) => r >= s.range.r1 && r <= s.range.r2 && c >= s.range.c1 && c <= s.range.c2);
  for (const cell of pendingExp) {
    const e = cell.exp!;
    const arr = arrays.find((a) => a.range.r1 === e.r && a.range.c1 === e.c);
    if (arr) {
      if (cell.r === e.r && cell.c === e.c) {
        const f = decodeFormula(arr.rgce, arr.rgcb, ctx(cell.r, cell.c));
        if (f != null) { cell.formula = f; cell.arrayRef = rangeA1(arr.range); }
      }
    } else {
      const s = findIn(shared, cell.r, cell.c, e.r, e.c);
      if (s) {
        const f = decodeFormula(s.rgce, s.rgcb, ctx(cell.r, cell.c));
        if (f != null) cell.formula = f;
      }
    }
    delete cell.exp;
  }
  // Comments: NOTE -> its OBJ id -> the TXO that followed that OBJ.
  for (const n of sh.notes) if (n.obj >= 0) n.text = txoByObj.get(n.obj) ?? n.text;
  // Pictures from the drawing.
  if (drawing.length && b8) {
    try {
      for (const shape of readSheetShapes(concatAll(drawing))) {
        if (shape.pib == null || !shape.anchor) continue;
        const o = shape.obj != null ? objs[shape.obj] : undefined;
        if (o && o.ot !== 8) continue;
        const blip = blips[shape.pib];
        if (blip) sh.images.push({ anchor: shape.anchor, blip, name: shape.name });
      }
    } catch {
      // A damaged drawing loses its pictures, not the sheet.
    }
  }
  // Shapes, text boxes and embedded charts.
  const seen = new Set<number>();
  if (drawing.length && b8) {
    try {
      const nodes = readDrawingTree(concatAll(drawing));
      const keep: DrawNode[] = [];
      const visit = (n: DrawNode, top: boolean): boolean => {
        if (n.obj != null) seen.add(n.obj);
        const ot = n.obj != null ? objs[n.obj]?.ot ?? -1 : -1;
        if (ot === 5) {
          const at = n.obj != null ? chartAt.get(n.obj) : undefined;
          const ch = at != null ? readChart(chartRecords(stream, at), chartEnv(book, base)) : null;
          if (ch && top) sh.chartObjs.push({ anchor: n.anchor, chart: ch });
          else sh.charts++;
          return false;
        }
        if (ot === 8 || ot === 0x19 || ot === 0x14) return false; // pictures, notes, filter arrows
        if (FORM_CONTROLS.has(ot)) { sh.shapes++; return false; }
        if (n.kids) n.kids = n.kids.filter((k) => visit(k, false));
        return !n.kids || n.kids.length > 0;
      };
      for (const n of nodes) if (visit(n, true)) keep.push(n);
      if (keep.length) sh.drawing = { nodes: keep, objs, txos };
    } catch {
      // A damaged drawing loses its shapes, not the sheet.
    }
  }
  // Objects no drawing describes (old-style OBJ records) cannot be kept either.
  objs.forEach((o, i) => { if (!seen.has(i) && o.ot !== 0x14 && o.ot !== 0x19 && o.ot !== 8) { if (o.ot === 5) sh.charts++; else sh.shapes++; } });
}

/** OBJ types that are form controls (buttons, check boxes, lists...), still not kept. */
const FORM_CONTROLS = new Set([7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 21, 22, 23, 24]);

/** What the chart reader needs from the book. */
export function chartEnv(book: XlsBook, base: Omit<FormulaContext, "row" | "col">): ChartEnv {
  return {
    formula: (rgce) => decodeFormula(rgce, undefined, { ...base, row: 0, col: 0 }) ?? undefined,
    font: (i) => book.fonts[i < 4 ? i : i - 1],
    format: (i) => book.formats.get(i) ?? BUILTIN_FMT[i],
    color: (icv) => icvCss(book.palette, icv),
  };
}

const BUILTIN_FMT: Record<number, string> = { 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 14: "m/d/yyyy" };

function readHyperlink(c: Cursor): XlsSheet["links"][number] | null {
  const range = range8(c);
  c.skip(16);
  c.u32();
  const flags = c.u32();
  const str = (): string => {
    const n = c.u32();
    if (n > 0x100000) throw new RangeError("hyperlink string");
    return utf16le(c.bytes(n * 2)).replace(/\u0000+$/, "");
  };
  if (flags & 0x10) str(); // display name
  if (flags & 0x80) str(); // frame
  let url: string | undefined;
  if (flags & 0x01) {
    if (flags & 0x100) url = str();
    else {
      const clsid = c.bytes(16);
      const hex = Array.from(clsid, (b) => b.toString(16).padStart(2, "0")).join("");
      if (hex === "e0c9ea79f9bace118c8200aa004ba90b") {
        const len = c.u32();
        url = utf16le(c.bytes(Math.min(len, c.left))).replace(/\u0000.*$/, "");
      } else if (hex === "0303000000000000c000000000000046") {
        const up = c.u16();
        const alen = c.u32();
        let path = new TextDecoder("windows-1252").decode(c.bytes(alen)).replace(/\u0000.*$/, "");
        c.skip(24);
        const sz = c.u32();
        if (sz > 0) {
          const ub = c.u32();
          c.u16();
          path = utf16le(c.bytes(ub));
        }
        url = "../".repeat(up) + path.replace(/\\/g, "/");
      }
    }
  }
  let location: string | undefined;
  if (flags & 0x08) location = str();
  if (!url && !location) return null;
  return { range, url, location };
}

/** How many of a TXO's CONTINUEs hold its text and runs; any after them are drawing data. */
function txoConts(r: BiffRecord): number {
  const d = r.data;
  const cch = d.length >= 14 ? d[10]! | (d[11]! << 8) : 0, cbRuns = d.length >= 14 ? d[12]! | (d[13]! << 8) : 0;
  let i = 0, left = cch;
  while (left > 0 && i < r.conts.length) { const c = r.conts[i++]!; left -= Math.floor((c.length - 1) / (c[0]! & 1 ? 2 : 1)); }
  for (let b = cbRuns; b > 0 && i < r.conts.length;) b -= r.conts[i++]!.length;
  return i;
}

function readTxo(r: BiffRecord): XlsTxo {
  const c = new Cursor(r.data);
  const grbit = c.u16(), rot = c.u16();
  c.skip(6);
  const cch = c.u16(), cbRuns = c.u16();
  if (!cch || !r.conts.length) return { text: "", runs: [], grbit, rot };
  const own = r.conts.slice(0, txoConts(r));
  const t = new Cursor({ data: own[0]!, conts: own.slice(1) });
  const flags = t.u8();
  const raw = t.chars8(cch, (flags & 1) === 1);
  // The formatting runs fill the last cbRuns bytes: ich, font, 4 reserved; the last marks the end.
  const all = concatAll(own);
  const runs: XlsTxo["runs"] = [];
  for (let o = Math.max(0, all.length - cbRuns); o + 8 <= all.length; o += 8) {
    const ich = all[o]! | (all[o + 1]! << 8), ifnt = all[o + 2]! | (all[o + 3]! << 8);
    if (ich < cch) runs.push({ ich, ifnt });
  }
  // A CR LF pair is two characters in the runs; keep their indexes in step with the text.
  let text = "";
  const map: number[] = [];
  for (let i = 0; i < raw.length; i++) {
    map[i] = text.length;
    if (raw[i] === "\r") { if (raw[i + 1] !== "\n") text += "\n"; continue; }
    text += raw[i];
  }
  return { text, runs: runs.map((x) => ({ ich: map[x.ich] ?? text.length, ifnt: x.ifnt })), grbit, rot };
}

function readCf(c: Cursor, ctx: FormulaContext): XlsCf | null {
  const type = c.u8(), op = c.u8();
  const cce1 = c.u16(), cce2 = c.u16();
  const flags = c.u32();
  const flags2 = c.u16();
  const rule: XlsCf = { type, op };
  if (flags & 0x02000000) {
    if (flags2 & 1) c.skip(c.u16() - 2);
    else c.skip(2);
  }
  if (flags & 0x04000000) {
    // DXFFntD
    const start = c.pos;
    c.skip(64);
    c.i32(); // height
    const ts = c.u32();
    const bls = c.i16();
    c.u16();
    const uls = c.u8();
    c.skip(3);
    const icv = c.i32();
    c.u32();
    const tsNinch = c.u32();
    c.u32();
    const ulsNinch = c.u32();
    const blsNinch = c.u32();
    c.skip(c.pos - start < 118 ? 118 - (c.pos - start) : 0);
    if (!(tsNinch & 2)) rule.italic = !!(ts & 2);
    if (!(tsNinch & 0x80)) rule.strike = !!(ts & 0x80);
    if (!blsNinch && bls !== -1) rule.bold = bls >= 700;
    if (!ulsNinch) rule.underline = uls > 0;
    if (icv !== -1 && icv >= 0 && icv < 0x7fff) rule.color = icv;
  }
  if (flags & 0x08000000) c.skip(8);
  if (flags & 0x10000000) c.skip(8);
  if (flags & 0x20000000) {
    const a = c.u16(), b = c.u16();
    const fls = (a >> 10) & 0x3f;
    const fore = b & 0x7f, back = (b >> 7) & 0x7f;
    // A solid CF fill shows its background colour; a pattern is approximated by its foreground.
    if (!(flags & 0x10000) && fls === 0) { /* explicitly no fill */ }
    else rule.fill = (fls <= 1 && !(flags & 0x40000)) || fore >= 64 ? back : fore;
  }
  if (flags & 0x40000000) c.skip(2);
  const r1 = c.bytes(cce1), r2 = c.bytes(cce2);
  if (cce1) rule.formula1 = decodeFormula(r1, undefined, ctx) ?? undefined;
  if (cce2) rule.formula2 = decodeFormula(r2, undefined, ctx) ?? undefined;
  return rule;
}

function readDv(c: Cursor, ctx: (r: number, c: number) => FormulaContext): XlsSheet["dvs"][number] | null {
  const f = c.u32();
  const str = (): string => { const s = readString(c, 8, 2); return s === "\u0000" ? "" : s; };
  const promptTitle = str(), errorTitle = str(), prompt = str(), error = str();
  const cce1 = c.u16(); c.u16();
  const rg1 = c.bytes(cce1).slice();
  const cce2 = c.u16(); c.u16();
  const rg2 = c.bytes(cce2).slice();
  const n = c.u16();
  const ranges: XlsRange[] = [];
  for (let i = 0; i < n && c.left >= 8; i++) ranges.push(range8(c));
  if (!ranges.length) return null;
  const at = ctx(ranges[0]!.r1, ranges[0]!.c1);
  const dv: XlsSheet["dvs"][number] = {
    type: f & 15, errStyle: (f >> 4) & 7, strList: !!(f & 0x80), allowBlank: !!(f & 0x100), op: (f >>> 20) & 15,
    promptTitle, errorTitle, prompt, error, ranges,
  };
  if (cce1) dv.formula1 = decodeFormula(rg1, undefined, at) ?? undefined;
  if (cce2) dv.formula2 = decodeFormula(rg2, undefined, at) ?? undefined;
  if (dv.strList && dv.formula1) dv.formula1 = dv.formula1.replace(/\u0000/g, ",");
  return dv;
}

/** Whether a .xls needs a password, and whether `password` opens it, without reading the book. */
export function xlsPasswordState(bytes: Uint8Array, password?: string): "none" | "ok" | "required" | "wrong" {
  let stream: Uint8Array | undefined;
  try {
    if (isCfb(bytes)) {
      const { streams } = readCfbStreams(bytes);
      stream = streams.get("Workbook") ?? streams.get("Book");
    } else if (isRawBiff(bytes)) stream = bytes;
  } catch {
    return "none";
  }
  if (!stream) return "none";
  let biff = 8;
  for (const r of records(stream)) {
    if (r.type === 0x0809 && r.data.length >= 2) biff = (r.data[0]! | (r.data[1]! << 8)) === 0x500 ? 5 : 8;
    if (r.type === 0x0085 || r.type === 0x000a || r.type === 0x00fc) return "none";
    if (r.type === 0x002f) {
      try {
        setupDecryption(r, biff, password);
        return password != null ? "ok" : "none"; // no password needed: the default one opened it
      } catch (e) {
        return e instanceof XlsPasswordError ? (e.code === "password-wrong" ? "wrong" : "required") : "none";
      }
    }
  }
  return "none";
}
