import type { Cell, Sheet, Workbook } from "../../core/model";
import { isNumeric, key, noteExtent } from "../../core/model";
// ---------------------------------------------------------------------------
// csv read: delimiter sniffing and span-preserving parsing. Each physical row
// keeps its exact original text (sheet.csvRows) so the writer can emit
// untouched rows byte-for-byte; only the parsed cells feed the grid.
// ---------------------------------------------------------------------------

export const CSV_DELIMITERS = [",", ";", "\t", "|"] as const;
const SNIFF_LINES = 20;
const SNIFF_CHARS = 65536;

/** Pick the delimiter whose per-line count (outside quotes) is most consistent. */
export function sniffCsvDelimiter(text: string): string {
  const counts = new Map<string, number[]>(CSV_DELIMITERS.map((d) => [d, [0]]));
  let inQuotes = false;
  let line = 0;
  const limit = Math.min(text.length, SNIFF_CHARS);
  for (let i = 0; i < limit && line < SNIFF_LINES; i++) {
    const ch = text[i]!;
    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && (ch === "\n" || ch === "\r")) {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      line++;
      for (const arr of counts.values()) arr.push(0);
      continue;
    }
    if (!inQuotes) {
      const arr = counts.get(ch);
      if (arr) arr[line]!++;
    }
  }
  let lines = counts.get(",")!.length;
  const arrays = [...counts.values()];
  while (lines > 1 && arrays.every((a) => a[lines - 1] === 0)) lines--;
  let best: { d: string; consistent: number } | null = null;
  for (const d of CSV_DELIMITERS) {
    const perLine = counts.get(d)!;
    const freq = new Map<number, number>();
    for (let i = 0; i < lines; i++) freq.set(perLine[i]!, (freq.get(perLine[i]!) ?? 0) + 1);
    let modal = 0;
    let modalFreq = 0;
    for (const [n, f] of freq) if (n > 0 && (f > modalFreq || (f === modalFreq && n > modal))) [modal, modalFreq] = [n, f];
    if (modal > 0 && (!best || modalFreq > best.consistent)) best = { d, consistent: modalFreq };
  }
  return best?.d ?? ",";
}

/** Parse one raw CSV row (no terminator) into unquoted field values. */
export function splitCsvRow(raw: string, delimiter: string): string[] {
  const cells: string[] = [];
  let field = "";
  let inQuotes = false;
  let fieldStart = true;
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i]!;
    if (inQuotes) {
      if (ch === '"') {
        if (raw[i + 1] === '"') {
          field += '"';
          i++;
        } else inQuotes = false;
      } else field += ch;
      continue;
    }
    if (ch === '"' && fieldStart) {
      inQuotes = true;
      fieldStart = false;
      continue;
    }
    if (ch === delimiter) {
      cells.push(field);
      field = "";
      fieldStart = true;
      continue;
    }
    field += ch;
    fieldStart = false;
  }
  cells.push(field);
  return cells;
}

/** Type a parsed CSV field into a model cell (formulas start with "="). */
function makeCell(row: number, col: number, text: string): Cell {
  if (text.startsWith("=") && text.length > 1) {
    return { row, col, value: "", kind: "blank", formula: text.slice(1) };
  }
  if (isNumeric(text)) return { row, col, value: text.trim(), kind: "n" };
  const up = text.toUpperCase();
  if (up === "TRUE" || up === "FALSE") return { row, col, value: up, kind: "b" };
  return { row, col, value: text, kind: "s" };
}

/**
 * Read CSV/TSV text into a single-sheet workbook. `delimiterHint` (a .tsv
 * extension, a host override) wins over sniffing. Formula cells have no cached
 * value in a CSV, so the caller recalculates after reading.
 */
export function readCsv(text: string, delimiterHint?: string, sheetName = "Sheet1"): Workbook {
  const delimiter = delimiterHint ?? sniffCsvDelimiter(text);
  const sheet: Sheet = { name: sheetName, cells: new Map(), maxRow: 0, maxCol: 0 };
  const csvRows: { raw: string; terminator: string; dirty: boolean; width: number }[] = [];
  let i = 0;
  let rowStart = 0;
  let inQuotes = false;
  const pushRow = (end: number, terminator: string) => {
    const raw = text.slice(rowStart, end);
    const cells = splitCsvRow(raw, delimiter);
    csvRows.push({ raw, terminator, dirty: false, width: cells.length });
    const r = csvRows.length;
    cells.forEach((v, idx) => {
      if (v === "") return; // blank fields stay implicit (no model cell)
      const cell = makeCell(r, idx + 1, v);
      sheet.cells.set(key(r, idx + 1), cell);
      noteExtent(sheet, r, idx + 1);
    });
    if (cells.length > 1) noteExtent(sheet, r, cells.length); // keep the grid width honest
  };
  while (i < text.length) {
    const ch = text[i]!;
    if (ch === '"') inQuotes = !inQuotes;
    if (!inQuotes && (ch === "\n" || ch === "\r")) {
      const terminator = ch === "\r" && text[i + 1] === "\n" ? "\r\n" : ch;
      pushRow(i, terminator);
      i += terminator.length;
      rowStart = i;
      continue;
    }
    i++;
  }
  if (rowStart < text.length) pushRow(text.length, "");
  sheet.csvRows = csvRows;
  return { kind: "csv", sheets: [sheet], files: {}, csvDelimiter: delimiter };
}
