import { appendFileSync, existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../core/workbook";
import { writeXls } from "./write";
import { isDateFmt } from "../../core/dates";
import { readXlsBook } from "./read";
import { xlsBookToXlsx } from "./to-xlsx";
import type { XlsBook } from "./types";

// A local corpus run, off unless XLS_CORPUS names a folder of .xls files. XLS_LO names a folder of
// LibreOffice per-sheet CSV exports ("<base>-<sheet>.csv") to compare every sheet's values against;
// XLS_WRITE round-trips each file through the editor model and the BIFF8 writer (XLS_OUT keeps the
// written files); XLS_REPORT writes the per-file results as JSON. The corpus is never committed.
const dir = process.env.XLS_CORPUS ?? "";
const loDir = process.env.XLS_LO ?? "";
const files = dir && existsSync(dir) ? readdirSync(dir).filter((f) => /\.xls$/i.test(f)).sort() : [];
const only = process.env.XLS_ONLY;
// jsdom parses large XML very slowly; bigger books skip the model round trip in this harness.
const MAX_XML = Number(process.env.XLS_MAX_XML ?? 1_500_000);

function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [], cur = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === ",") { row.push(cur); cur = ""; }
    else if (ch === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; }
    else if (ch !== "\r") cur += ch;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

const norm = (s: string): string => s.replace(/\s+/g, " ").trim();
// LibreOffice's unformatted export still writes dates as MM/DD/YYYY [HH:MM:SS] and percents with %.
function loNumber(s: string): number | null {
  const d = /^(\d\d)\/(\d\d)\/(-?\d{4})(?: (\d\d):(\d\d):(\d\d(?:\.\d+)?))?$/.exec(s);
  if (d) {
    const ms = Date.UTC(Number(d[3]), Number(d[1]) - 1, Number(d[2])) - Date.UTC(1899, 11, 30);
    let serial = ms / 86400000 + (Number(d[4] ?? 0) * 3600 + Number(d[5] ?? 0) * 60 + Number(d[6] ?? 0)) / 86400;
    if (serial < 61 && serial >= 1) serial -= 1; // Excel's phantom 29 Feb 1900
    return serial;
  }
  const t = /^(\d\d):(\d\d):(\d\d(?:\.\d+)?)$/.exec(s);
  if (t) return (Number(t[1]) * 3600 + Number(t[2]) * 60 + Number(t[3])) / 86400;
  const pct = /^(-?[\d.]+(?:E[-+]?\d+)?)%$/i.exec(s);
  if (pct) return Number(pct[1]) / 100;
  const n = Number(s.replace(/[,$€£¥\s]/g, ""));
  return s !== "" && Number.isFinite(n) ? n : null;
}
function same(a: string, b: string): boolean {
  a = norm(a); b = norm(b);
  if (a === b) return true;
  const la = loNumber(a), nb = Number(b);
  if (la != null && b !== "" && Number.isFinite(nb) && Math.abs(la - nb) <= 1e-6 * Math.max(1, Math.abs(nb))) return true;
  if ((a === "TRUE" && b === "1") || (a === "FALSE" && b === "0")) return true;
  return a.toUpperCase() === b.toUpperCase();
}

function textOf(book: XlsBook, s: XlsBook["sheets"][number], r: number, c: number): string {
  const cell = s.cells.get(r * 256 + c);
  if (!cell) return "";
  const v = cell.val;
  if (v.k === "s") return v.v;
  if (v.k === "b") return v.v ? "TRUE" : "FALSE";
  if (v.k === "e") return v.v;
  if (v.k === "blank") return "";
  const id = book.xfs[cell.xf]?.fmt;
  const fmt = id ? book.formats.get(id) ?? (book.biff >= 5 ? id : undefined) : undefined;
  return String(book.date1904 && fmt != null && isDateFmt(fmt) ? v.v + 1462 : v.v);
}

function compareSheets(book: XlsBook, base: string): { sheets: number; matched: number; cells: number; bad: number; sample?: string } {
  let sheets = 0, matched = 0, cells = 0, bad = 0;
  let sample: string | undefined;
  for (const s of book.sheets) {
    const path = join(loDir, `${base}-${s.name}.csv`);
    if (!existsSync(path)) continue;
    sheets++;
    const rows = parseCsv(readFileSync(path, "utf8"));
    let sheetBad = 0;
    rows.forEach((row, r) => row.forEach((v, c) => {
      const mine = c < 256 ? textOf(book, s, r, c) : "";
      if (!v && !mine) return;
      cells++;
      if (!same(v, mine)) { sheetBad++; if (!sample) sample = `${s.name}!R${r + 1}C${c + 1} lo=${JSON.stringify(v).slice(0, 60)} mine=${JSON.stringify(mine).slice(0, 60)}`; }
    }));
    bad += sheetBad;
    if (sheetBad === 0) matched++;
  }
  return { sheets, matched, cells, bad, sample };
}

/** Every cell's value and formula, every sheet, as one comparable string. */
function sig(book: XlsBook): string {
  const isDate = (xf: number): boolean => { const id = book.xfs[xf]?.fmt; return !!id && /[dmyhs]/i.test((book.formats.get(id) ?? (id >= 14 && id <= 22 ? "d" : "")).replace(/"[^"]*"|\[[^\]]*\]/g, "")); };
  return book.sheets.map((s) => s.name + ":" + [...s.cells.values()]
    // An empty string and a non-number are written as blank cells.
    .filter((c) => (c.val.k !== "blank" && !(c.val.k === "s" && c.val.v === "") && !(c.val.k === "n" && !Number.isFinite(c.val.v))) || c.formula)
    .map((c) => {
      const v = c.val;
      const n = v.k === "n" && book.date1904 && isDate(c.xf) ? v.v + 1462 : v.k === "n" ? v.v : 0;
      const val = v.k === "n" ? String(parseFloat(n.toPrecision(15))) : v.k === "blank" ? "" : String(v.v);
      return `${c.r},${c.c}=${v.k === "blank" ? "s" : v.k}:${val}${c.formula ? "|" + c.formula : ""}`;
    })
    .sort().join(";")).join("\n") + "\nnames:" + book.names.filter((n) => n.formula && !n.func && !n.name.startsWith("_xlfn.")).map((n) => `${n.name}@${n.itab}=${n.formula}`).sort().join(";");
}

describe.skipIf(!process.env.XLS_DUMP)("xls dump", () => {
  it("dumps a book", () => {
    const book = readXlsBook(new Uint8Array(readFileSync(process.env.XLS_DUMP!)), { password: process.env.XLS_PW });
    const lines: string[] = [JSON.stringify({ biff: book.biff, enc: book.encrypted, supbooks: book.supbooks, xti: book.xti, names: book.names.map((n) => [n.name, n.itab, n.formula]) })];
    for (const s of book.sheets) for (const c of s.cells.values()) lines.push(`${s.name}!${c.r},${c.c} ${JSON.stringify(c.val)} ${c.formula ?? ""}`);
    writeFileSync(process.env.XLS_DUMP + ".txt", lines.join("\n"));
  });
});

describe.skipIf(!files.length)("xls corpus", () => {
  it("opens, matches LibreOffice and round-trips", () => {
    const report: Record<string, unknown>[] = [];
    for (const f of files) {
      if (only && !f.includes(only)) continue;
      const bytes = new Uint8Array(readFileSync(join(dir, f)));
      if (process.env.XLS_PROGRESS) appendFileSync(process.env.XLS_PROGRESS, f + "\n");
      const t0 = performance.now();
      const entry: Record<string, unknown> = { f };
      try {
        const book = readXlsBook(bytes);
        entry.ok = true;
        entry.ms = Math.round(performance.now() - t0);
        entry.biff = book.biff;
        entry.sheetCount = book.sheets.length;
        if (loDir) Object.assign(entry, compareSheets(book, f.replace(/\.xls$/i, "")));
        if (process.env.XLS_WRITE) {
          const xml = Object.values(xlsBookToXlsx(book)).reduce((n, b) => n + b.length, 0);
          if (xml > MAX_XML) entry.rt = "skipped (large)";
          else {
            try {
              const out = writeXls(readWorkbook(bytes)); // no recalc: the cached values are what round-trips
              const again = readXlsBook(out);
              const a = sig(book), b = sig(again);
              entry.rt = a === b ? "same" : "diff";
              if (a !== b) {
                const la = a.split(/[;\n]/), lb = new Set(b.split(/[;\n]/));
                entry.rtSample = la.filter((x) => !lb.has(x)).slice(0, 3);
              }
              if (process.env.XLS_OUT) writeFileSync(join(process.env.XLS_OUT, f), out);
            } catch (e) {
              entry.rt = "error: " + String((e as Error).stack).slice(0, 300);
            }
          }
        }
      } catch (e) {
        entry.ok = false;
        entry.err = String((e as Error).message).slice(0, 200);
      }
      report.push(entry);
    }
    if (process.env.XLS_REPORT) writeFileSync(process.env.XLS_REPORT, JSON.stringify(report, null, 1));
    expect(report.length).toBeGreaterThan(0);
  }, 3_600_000);
});
