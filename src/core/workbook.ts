import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type { Sheet, Workbook } from "./model";
import { ensureCell, firstByLocal, formatNumber, getCell, isNumeric, parseXmlOpt, serializeXml } from "./model";
import { readCsv, writeCsv } from "../adapters/csv";
import { readOds, writeOds } from "../adapters/ods";
import { recalc } from "./recalc";
import { readXlsx, writeXlsx } from "../adapters/xlsx";
// ---------------------------------------------------------------------------
// Public read / write
// ---------------------------------------------------------------------------

export interface ReadOptions {
  /** Route plain-text input explicitly (a .csv/.tsv extension known to the host). */
  formatHint?: "csv" | "tsv";
}

const readCsvBytes = (bytes: Uint8Array, hint?: "csv" | "tsv"): Workbook => {
  const text = new TextDecoder("utf-8", { fatal: false }).decode(bytes);
  const wb = readCsv(text, hint === "tsv" ? "\t" : undefined);
  // CSV formulas carry no cached results; compute them once on open.
  const hasFormula = wb.sheets.some((s) => [...s.cells.values()].some((c) => c.formula != null));
  if (hasFormula) recalc(wb);
  return wb;
};

/** Text-looking bytes (no control chars beyond tab/newline in the head) may be
    CSV; zips, CFB containers and binary junk never qualify. */
const looksTextual = (bytes: Uint8Array): boolean => {
  const n = Math.min(bytes.length, 4096);
  for (let i = 0; i < n; i++) {
    const b = bytes[i]!;
    if (b < 0x20 && b !== 0x09 && b !== 0x0a && b !== 0x0d) return false;
  }
  return n > 0;
};

export function readWorkbook(bytes: Uint8Array, opts: ReadOptions = {}): Workbook {
  if (opts.formatHint === "csv" || opts.formatHint === "tsv") return readCsvBytes(bytes, opts.formatHint);
  // Encrypted .xlsx and legacy binary .xls are CFB containers, not zips.
  if (bytes.length >= 4 && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0)
    throw new Error("password-protected or legacy binary workbook");
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch {
    if (looksTextual(bytes)) return readCsvBytes(bytes);
    throw new Error("not a valid workbook file (unreadable archive)");
  }
  if (files["xl/workbook.xml"]) return readXlsx(files);
  if (files["content.xml"]) {
    const mt = files["mimetype"] ? strFromU8(files["mimetype"]) : "";
    if (!mt || mt.includes("spreadsheet")) return readOds(files);
  }
  throw new Error("unrecognized workbook: expected .xlsx or .ods");
}

// After a formula was added/changed/removed, xl/calcChain.xml is stale (a known
// Excel "repair" trigger) and cached values may be too: drop the chain, detach its
// content-type/relationship entries, and ask for a full recalc on open.
export function dropStaleCalcChain(wb: Workbook): void {
  let formulasChanged = false;
  outer: for (const s of wb.sheets)
    for (const c of s.cells.values())
      if (c.fDirty) {
        formulasChanged = true;
        break outer;
      }
  if (!formulasChanged) return;
  if (wb.files["xl/calcChain.xml"]) {
    delete wb.files["xl/calcChain.xml"];
    const ct = wb.files["[Content_Types].xml"];
    const ctDoc = ct ? parseXmlOpt(ct) : undefined;
    if (ctDoc) {
      const doc = ctDoc;
      for (const o of Array.from(doc.getElementsByTagName("Override")))
        if (o.getAttribute("PartName") === "/xl/calcChain.xml") o.parentNode?.removeChild(o);
      wb.files["[Content_Types].xml"] = serializeXml(doc);
    }
    const relsPath = "xl/_rels/workbook.xml.rels";
    const relsDoc2 = wb.files[relsPath] ? parseXmlOpt(wb.files[relsPath]!) : undefined;
    if (relsDoc2) {
      const doc = relsDoc2;
      for (const r of Array.from(doc.getElementsByTagName("Relationship")))
        if ((r.getAttribute("Target") ?? "").endsWith("calcChain.xml")) r.parentNode?.removeChild(r);
      wb.files[relsPath] = serializeXml(doc);
    }
  }
  const wbXml = wb.files["xl/workbook.xml"];
  const wbDoc2 = wbXml ? parseXmlOpt(wbXml) : undefined;
  if (wbDoc2) {
    const doc = wbDoc2;
    const root = doc.documentElement;
    let calcPr = firstByLocal(root, "calcPr");
    if (!calcPr) {
      calcPr = doc.createElementNS(root.namespaceURI, "calcPr");
      // CT_Workbook is a strict sequence: insert before the first element that follows calcPr.
      const after = ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];
      let anchor: Element | null = null;
      for (const ch of Array.from(root.children))
        if (after.includes(ch.localName)) {
          anchor = ch;
          break;
        }
      root.insertBefore(calcPr, anchor);
    }
    calcPr.setAttribute("fullCalcOnLoad", "1");
    wb.files["xl/workbook.xml"] = serializeXml(doc);
  }
}

export function writeWorkbook(wb: Workbook): Uint8Array {
  recalc(wb);
  if (wb.kind === "csv") return strToU8(writeCsv(wb));
  if (wb.kind === "xlsx") {
    writeXlsx(wb);
    dropStaleCalcChain(wb);
    return zipSync(wb.files);
  }
  writeOds(wb);
  // ODF requires the "mimetype" entry first and stored (uncompressed).
  const repacked: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {};
  if (wb.files["mimetype"]) repacked["mimetype"] = [wb.files["mimetype"], { level: 0 }];
  for (const [name, data] of Object.entries(wb.files)) {
    if (name === "mimetype") continue;
    repacked[name] = data;
  }
  return zipSync(repacked as Record<string, Uint8Array>);
}

/** Commit a raw grid edit (the text a user typed) into the model. */
export function setCellInput(sheet: Sheet, row: number, col: number, raw: string): void {
  const existing = getCell(sheet, row, col);
  if (raw.startsWith("=")) {
    const cell = ensureCell(sheet, row, col);
    const nf = raw.slice(1).trim();
    if (cell.formula !== nf) cell.fDirty = true;
    cell.formula = nf;
    cell.odfFormula = undefined;
    cell.edited = true;
    return;
  }
  if (existing == null && raw === "") return;
  const cell = ensureCell(sheet, row, col);
  if (cell.formula != null) cell.fDirty = true;
  cell.formula = undefined;
  cell.odfFormula = undefined;
  cell.edited = true;
  if (raw === "") {
    cell.value = "";
    cell.kind = "blank";
  } else if (isNumeric(raw)) {
    cell.value = raw.trim();
    cell.kind = "n";
  } else if (raw.toUpperCase() === "TRUE" || raw.toUpperCase() === "FALSE") {
    cell.value = raw.toUpperCase();
    cell.kind = "b";
  } else {
    cell.value = raw;
    cell.kind = "s";
  }
  // Re-apply the cell's number format (a typed value keeps the cell's format, like Excel).
  cell.display =
    cell.kind === "n" && cell.numFmt != null ? formatNumber(cell.numFmt, cell.value) ?? undefined : undefined;
}

