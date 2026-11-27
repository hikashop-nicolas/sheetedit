import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../core/workbook";
import { getCell, type Workbook } from "../../core/model";
import { sniffSheetText } from "./text";
import { readSheetText } from ".";

// A local corpus run, off unless TAB_CORPUS names a folder of HTML / Excel 2003 XML exports.
// TAB_LO names a folder of LibreOffice per-sheet CSV exports ("<base>-<sheet>.csv") to compare
// values against; TAB_REPORT writes the per-file results as JSON. The corpus is never committed.
const dir = process.env.TAB_CORPUS ?? "";
const loDir = process.env.TAB_LO ?? "";
const files = dir && existsSync(dir) ? readdirSync(dir).filter((f) => /\.(xls|xml|html?|mht)$/i.test(f)).sort() : [];

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
function loNumber(s: string): number | null {
  const d = /^(\d\d)\/(\d\d)\/(-?\d{4})(?: (\d\d):(\d\d):(\d\d(?:\.\d+)?))?$/.exec(s);
  if (d) {
    const ms = Date.UTC(Number(d[3]), Number(d[1]) - 1, Number(d[2])) - Date.UTC(1899, 11, 30);
    let serial = ms / 86400000 + (Number(d[4] ?? 0) * 3600 + Number(d[5] ?? 0) * 60 + Number(d[6] ?? 0)) / 86400;
    if (serial < 61 && serial >= 1) serial -= 1;
    return serial;
  }
  const iso = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d\d)(?::(\d\d))?)?$/.exec(s);
  if (iso) {
    const ms = Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])) - Date.UTC(1899, 11, 30);
    return ms / 86400000 + (Number(iso[4] ?? 0) * 3600 + Number(iso[5] ?? 0) * 60 + Number(iso[6] ?? 0)) / 86400;
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
  return a.toUpperCase() === b.toUpperCase();
}

function grid(wb: Workbook, si: number): string[][] {
  const s = wb.sheets[si]!;
  const out: string[][] = [];
  for (const c of s.cells.values()) {
    const v = c.kind === "blank" ? "" : c.kind === "b" ? (c.value === "TRUE" || c.value === "1" ? "TRUE" : "FALSE") : String(c.value);
    if (v === "") continue;
    (out[c.row - 1] ??= [])[c.col - 1] = v;
  }
  return out;
}

describe.skipIf(!files.length)("text spreadsheet corpus", () => {
  const report: Record<string, unknown>[] = [];
  for (const f of files) {
    it(f, () => {
      const bytes = new Uint8Array(readFileSync(join(dir, f)));
      const ext = f.slice(f.lastIndexOf("."));
      const kind = sniffSheetText(bytes, ext);
      const row: Record<string, unknown> = { file: f, kind };
      report.push(row);
      // Strictly named pages (.html) that are not claimed are still read, to compare them.
      if (!kind && !/\.html?$/i.test(f)) return;
      const t0 = Date.now();
      let wb: Workbook;
      try { wb = readWorkbook(bytes); } catch (e) { row.error = String(e); return; }
      row.ms = Date.now() - t0;
      row.sheets = wb.sheets.length;
      row.cells = wb.sheets.reduce((n, s) => n + [...s.cells.values()].filter((c) => c.kind !== "blank").length, 0);
      // The tokenizer fallback must read what DOMParser reads.
      if (kind) {
        const alt = readSheetText(bytes, kind, { useDom: false })!;
        const flat = (w: Workbook): string => w.sheets.map((_, i) => JSON.stringify(grid(w, i))).join("|");
        row.fallbackSame = flat(alt) === flat(wb);
        if (!row.fallbackSame) {
          const a = flat(alt), b = flat(wb);
          let i = 0;
          while (a[i] === b[i]) i++;
          row.fallbackDiff = `dom=${b.slice(Math.max(0, i - 40), i + 40)} tok=${a.slice(Math.max(0, i - 40), i + 40)}`;
        }
      }
      if (!loDir) return;
      const base = f.slice(0, f.lastIndexOf("."));
      const all = readdirSync(loDir).filter((x) => x.startsWith(base + "-") && x.endsWith(".csv"));
      // LibreOffice files are named after the sheets; line them up with ours by name.
      const byName = wb.sheets.map((s) => all.find((x) => x === `${base}-${s.name}.csv`));
      const lo = byName.every((x) => x) && byName.length === all.length ? (byName as string[]) : all.sort();
      row.loSheets = lo.length;
      // Value coverage: every LibreOffice value found among ours (as a multiset), whatever the layout.
      const ours: string[] = [];
      wb.sheets.forEach((_, i) => grid(wb, i).forEach((r) => r?.forEach((v) => { if (v != null) ours.push(v); })));
      const loVals: string[] = [];
      const loGrids = lo.map((x) => parseCsv(readFileSync(join(loDir, x), "utf8")));
      for (const g of loGrids) for (const r of g) for (const v of r) if (norm(v)) loVals.push(v);
      const pool = new Map<string, number>();
      const key = (v: string): string => { const n = loNumber(norm(v)); return n != null ? "n:" + Number(n.toPrecision(10)) : "s:" + norm(v).toUpperCase(); };
      for (const v of ours) pool.set(key(v), (pool.get(key(v)) ?? 0) + 1);
      let found = 0;
      const missing: string[] = [];
      for (const v of loVals) { const k = key(v); const n = pool.get(k) ?? 0; if (n > 0) { found++; pool.set(k, n - 1); } else if (missing.length < 4) missing.push(v); }
      row.loValues = loVals.length;
      row.found = found;
      row.extra = [...pool.values()].reduce((a, b) => a + b, 0);
      if (missing.length) row.missing = missing;
      // Cell by cell when the sheets line up one to one.
      if (lo.length === wb.sheets.length) {
        let compared = 0, diff = 0;
        const diffs: string[] = [];
        loGrids.forEach((g, si) => {
          const mine = grid(wb, si);
          const rows = Math.max(g.length, mine.length);
          for (let r = 0; r < rows; r++) {
            const cols = Math.max(g[r]?.length ?? 0, mine[r]?.length ?? 0);
            for (let c = 0; c < cols; c++) {
              const a = g[r]?.[c] ?? "", b = mine[r]?.[c] ?? "";
              if (!norm(a) && !norm(b)) continue;
              compared++;
              if (!same(a, b)) { diff++; if (diffs.length < 4) diffs.push(`${si}:R${r + 1}C${c + 1} lo=${JSON.stringify(a)} ours=${JSON.stringify(b)}`); }
            }
          }
        });
        row.compared = compared;
        row.diff = diff;
        if (diffs.length) row.diffs = diffs;
      }
      expect(wb.sheets.length).toBeGreaterThan(0);
    }, 60000);
  }
  it("writes the report", () => {
    if (process.env.TAB_REPORT) writeFileSync(process.env.TAB_REPORT, JSON.stringify(report, null, 1));
  });
});

// TAB_LOX names a folder of LibreOffice xlsx conversions ("<base>.xlsx"): formatting is compared
// cell by cell (fonts, fills, borders, alignment, number formats) with merges, widths and panes.
const loxDir = process.env.TAB_LOX ?? "";
describe.skipIf(!files.length || !loxDir)("text spreadsheet formatting", () => {
  const totals: Record<string, [number, number]> = {};
  const samples: Record<string, string[]> = {};
  let where = "";
  const tally = (k: string, ok: boolean, what = ""): void => {
    const t = (totals[k] ??= [0, 0]);
    t[0]++;
    if (!ok) { t[1]++; const s = (samples[k] ??= []); if (s.length < 6) s.push(`${where} ${what}`); }
  };
  for (const f of files) {
    const base = f.slice(0, f.lastIndexOf("."));
    const lox = join(loxDir, base + ".xlsx");
    if (!existsSync(lox)) continue;
    it(f, () => {
      const bytes = new Uint8Array(readFileSync(join(dir, f)));
      if (!sniffSheetText(bytes, ".xls")) return;
      const ours = readWorkbook(bytes), lo = readWorkbook(new Uint8Array(readFileSync(lox)));
      ours.sheets.forEach((s, si) => {
        const l = lo.sheets.find((x) => x.name === s.name) ?? (lo.sheets.length === ours.sheets.length ? lo.sheets[si] : undefined);
        if (!l) return;
        for (const c of s.cells.values()) {
          if (c.kind === "blank" && !c.cellStyle) continue;
          const o = getCell(l, c.row, c.col);
          // Only cells whose values agree, so a shifted layout does not count as formatting.
          if (!o || !same(String(o.value ?? ""), String(c.value ?? ""))) continue;
          const a = c.cellStyle ?? {}, b = o.cellStyle ?? {};
          where = `${f.slice(0, 40)} ${s.name} R${c.row}C${c.col}`;
          tally("bold", !!a.bold === !!b.bold, `ours=${!!a.bold} lo=${!!b.bold}`);
          tally("italic", !!a.italic === !!b.italic);
          tally("underline", !!a.underline === !!b.underline, `ours=${!!a.underline} lo=${!!b.underline}`);
          tally("color", (a.color ?? "#000000").toLowerCase() === (b.color ?? "#000000").toLowerCase(), `ours=${a.color} lo=${b.color}`);
          tally("fill", (a.bg ?? "").toLowerCase() === (b.bg ?? "").toLowerCase(), `ours=${a.bg} lo=${b.bg}`);
          const sides = (x: typeof a): string => ["top", "right", "bottom", "left"].filter((k) => x.borders?.[k as "top"]).join(",");
          tally("borders", sides(a) === sides(b), `ours=${sides(a)} lo=${sides(b)}`);
          tally("align", (a.align ?? "") === (b.align ?? ""), `ours=${a.align} lo=${b.align}`);
          tally("wrap", !!a.wrap === !!b.wrap, `ours=${!!a.wrap} lo=${!!b.wrap}`);
          if (c.kind === "n") tally("numFmt", String(c.display ?? c.value) === String(o.display ?? o.value), `ours=${c.display ?? c.value} (${c.numFmt}) lo=${o.display ?? o.value} (${o.numFmt})`);
        }
        const mk = (x: typeof s): string => (x.merges ?? []).map((m) => `${m.r1},${m.c1},${m.r2},${m.c2}`).sort().join(";");
        where = `${f.slice(0, 40)} ${s.name}`;
        tally("merges", mk(s) === mk(l), `ours=${mk(s).slice(0, 60)} lo=${mk(l).slice(0, 60)}`);
        tally("hiddenRows", [...(s.hiddenRows ?? [])].sort().join() === [...(l.hiddenRows ?? [])].sort().join());
        tally("hiddenCols", [...(s.hiddenCols ?? [])].sort().join() === [...(l.hiddenCols ?? [])].sort().join());
        tally("freeze", `${s.freeze?.rows ?? 0},${s.freeze?.cols ?? 0}` === `${l.freeze?.rows ?? 0},${l.freeze?.cols ?? 0}`, `ours=${JSON.stringify(s.freeze)} lo=${JSON.stringify(l.freeze)}`);
        for (const [c, w] of s.colWidths ?? []) { const lw = l.colWidths?.get(c); if (lw) tally("colWidth", Math.abs(lw - w) <= Math.max(4, lw * 0.15), `col ${c} ours=${w} lo=${lw}`); }
      });
    }, 60000);
  }
  it("prints the formatting tally", () => {
    for (const [k, [n, bad]] of Object.entries(totals)) console.log(`STYLE ${k.padEnd(10)} ${n - bad}/${n} agree${process.env.TAB_SAMPLES ? "\n  " + (samples[k] ?? []).join("\n  ") : ""}`);
  });
});

// TAB_NEG names a folder of ordinary pages and XML files: none may be claimed under its own name.
const negDir = process.env.TAB_NEG ?? "";
const negs = negDir && existsSync(negDir) ? readdirSync(negDir).sort() : [];
describe.skipIf(!negs.length)("text spreadsheet negatives", () => {
  for (const f of negs) {
    it(f, () => {
      const bytes = new Uint8Array(readFileSync(join(negDir, f)));
      const ext = f.includes(".") ? f.slice(f.lastIndexOf(".")) : "";
      const own = sniffSheetText(bytes, ext), asXls = sniffSheetText(bytes, ".xls");
      console.log(`NEG ${f} own=${own} asXls=${asXls}`);
      expect(own).toBeNull();
    });
  }
});
