import { strToU8, zipSync } from "fflate";
import type { Cell, Workbook } from "../../core/model";
import { colToLetters } from "../../core/model";
import { SS_MAIN } from "../xlsx/shared";
// ---------------------------------------------------------------------------
// csv -> xlsx conversion: build a minimal real workbook from the model
// (values, formulas, column widths). Formulas carry their computed result as
// the cached value; fullCalcOnLoad still asks the target app to recompute.
// ---------------------------------------------------------------------------

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function csvToXlsx(wb: Workbook): Uint8Array {
  const sheet = wb.sheets[0]!;
  const doc = document.implementation.createDocument(SS_MAIN, "worksheet", null);
  const ws = doc.documentElement;
  const ce = (name: string) => doc.createElementNS(SS_MAIN, name);

  if (sheet.colWidths?.size) {
    const cols = ce("cols");
    for (const [c, px] of [...sheet.colWidths.entries()].sort((a, b) => a[0] - b[0])) {
      const col = ce("col");
      col.setAttribute("min", String(c));
      col.setAttribute("max", String(c));
      col.setAttribute("width", String((px - 5) / 7));
      col.setAttribute("customWidth", "1");
      cols.appendChild(col);
    }
    ws.appendChild(cols);
  }

  const sheetData = ce("sheetData");
  const byRow = new Map<number, Cell[]>();
  for (const cell of sheet.cells.values()) {
    if (cell.kind === "blank" && cell.formula == null) continue;
    let list = byRow.get(cell.row);
    if (!list) byRow.set(cell.row, (list = []));
    list.push(cell);
  }
  for (const r of [...byRow.keys()].sort((a, b) => a - b)) {
    const rowEl = ce("row");
    rowEl.setAttribute("r", String(r));
    for (const cell of byRow.get(r)!.sort((a, b) => a.col - b.col)) {
      const c = ce("c");
      c.setAttribute("r", colToLetters(cell.col) + r);
      if (cell.formula != null) {
        const f = ce("f");
        f.textContent = cell.formula;
        c.appendChild(f);
        // The computed result rides along as the cached value, like Excel writes it,
        // so any reader that trusts caches shows the number immediately.
        if (cell.kind === "b") c.setAttribute("t", "b");
        else if (cell.kind === "e") c.setAttribute("t", "e");
        else if (cell.kind === "s") c.setAttribute("t", "str");
        if (cell.value !== "") {
          const v = ce("v");
          v.textContent = cell.kind === "b" ? (cell.value === "TRUE" ? "1" : "0") : cell.value;
          c.appendChild(v);
        }
      } else if (cell.kind === "n") {
        const v = ce("v");
        v.textContent = cell.value;
        c.appendChild(v);
      } else if (cell.kind === "b") {
        c.setAttribute("t", "b");
        const v = ce("v");
        v.textContent = cell.value === "TRUE" ? "1" : "0";
        c.appendChild(v);
      } else {
        c.setAttribute("t", "inlineStr");
        const is = ce("is");
        const t = ce("t");
        t.setAttribute("xml:space", "preserve");
        t.textContent = cell.value;
        is.appendChild(t);
        c.appendChild(is);
      }
      rowEl.appendChild(c);
    }
    sheetData.appendChild(rowEl);
  }
  ws.appendChild(sheetData);
  const sheetXml = XML_DECL + new XMLSerializer().serializeToString(doc);

  const workbookXml =
    XML_DECL +
    `<workbook xmlns="${SS_MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>`;
  const relsXml =
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';
  const rootRels =
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
  const contentTypes =
    XML_DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';

  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(relsXml),
    "xl/worksheets/sheet1.xml": strToU8(sheetXml),
  });
}
