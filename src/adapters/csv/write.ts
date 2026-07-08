import type { Cell, Sheet, Workbook } from "../../core/model";
// ---------------------------------------------------------------------------
// csv write: untouched rows are emitted byte-for-byte from sheet.csvRows;
// only rows whose cells were edited (or whose formulas were rewritten, or
// that structural ops dirtied) are re-serialized. A recalculated cached value
// does not dirty a row: a CSV stores the formula text, not the result.
// ---------------------------------------------------------------------------

/** Serialize one cell for CSV output, quoting only where the format needs it. */
function fieldOf(cell: Cell | undefined, delimiter: string): string {
  if (!cell) return "";
  const text = cell.formula != null ? "=" + cell.formula : cell.value;
  if (text.includes(delimiter) || text.includes('"') || text.includes("\n") || text.includes("\r")) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}

export function writeCsv(wb: Workbook): string {
  const sheet: Sheet = wb.sheets[0]!;
  const delimiter = wb.csvDelimiter ?? ",";
  const rows = sheet.csvRows ?? [];
  const dominant = rows.find((r) => r.terminator)?.terminator ?? "\n";
  const endsWithNewline = rows.length === 0 || rows[rows.length - 1]!.terminator !== "";

  const dirtyRows = new Set<number>();
  const rowCells = new Map<number, Cell[]>();
  for (const cell of sheet.cells.values()) {
    if (cell.edited || cell.fDirty) dirtyRows.add(cell.row);
    let list = rowCells.get(cell.row);
    if (!list) rowCells.set(cell.row, (list = []));
    list.push(cell);
  }

  const total = Math.max(rows.length, sheet.maxRow);
  let out = "";
  for (let r = 1; r <= total; r++) {
    const orig = rows[r - 1];
    const last = r === total;
    // A row must end with a terminator unless it is the (newline-less) final row.
    const terminator = last ? (endsWithNewline ? orig?.terminator || dominant : "") : orig?.terminator || dominant;
    if (orig && !orig.dirty && !dirtyRows.has(r)) {
      out += orig.raw + terminator;
      continue;
    }
    const cells = (rowCells.get(r) ?? []).sort((a, b) => a.col - b.col);
    const width = Math.max(orig?.width ?? 0, cells.length ? cells[cells.length - 1]!.col : 0);
    const byCol = new Map(cells.map((c) => [c.col, c]));
    const vals: string[] = [];
    for (let c = 1; c <= width; c++) {
      const cell = byCol.get(c);
      // A cleared cell (blank) is an empty field; fully empty rows become blank lines.
      vals.push(cell && cell.kind !== "blank" ? fieldOf(cell, delimiter) : cell?.formula != null ? fieldOf(cell, delimiter) : "");
    }
    out += vals.join(delimiter) + terminator;
  }
  return out;
}
