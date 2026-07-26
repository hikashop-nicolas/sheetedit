import { colToLetters, firstByLocal, parseXmlOpt, serializeXml, type Sheet, type Workbook } from "../../core/model";
import { formatHeaderFooter, hasPrintSetup, type PrintSetup } from "../../core/print";
import { SS_MAIN } from "./shared";
import { insertWsChild } from "./write";

// Writing print setup means five worksheet elements plus one flag on sheetPr, and the print area /
// print titles as sheet-scoped defined names in workbook.xml. Each element is patched in place when
// it exists, so the attributes sheetedit does not model (dpi, copies, cellComments, ...) survive.

/** Find-or-create a worksheet child, inserted in canonical order. */
function ensureChild(doc: Document, ws: Element, name: string): Element {
  const found = firstByLocal(ws, name);
  if (found) return found;
  const el = doc.createElementNS(ws.namespaceURI || SS_MAIN, name);
  insertWsChild(ws, el);
  return el;
}

const setOrDrop = (el: Element, name: string, v: string | undefined): void => {
  if (v == null) el.removeAttribute(name);
  else el.setAttribute(name, v);
};
const boolAttr = (v: boolean | undefined): string | undefined => (v === undefined ? undefined : v ? "true" : "false");

/** Rewrite a <rowBreaks> / <colBreaks> element from the model's 1-based lines. */
function writeBreaks(doc: Document, ws: Element, name: string, lines: number[] | undefined, maxOther: number): void {
  const existing = firstByLocal(ws, name);
  if (!lines?.length) {
    if (existing) existing.parentNode?.removeChild(existing);
    return;
  }
  const el = existing ?? ensureChild(doc, ws, name);
  while (el.firstChild) el.removeChild(el.firstChild);
  for (const line of lines) {
    const brk = doc.createElementNS(el.namespaceURI || SS_MAIN, "brk");
    // The file counts the break 0-based from the origin; the model names the line that starts the page.
    brk.setAttribute("id", String(line - 1));
    brk.setAttribute("max", String(maxOther));
    brk.setAttribute("man", "1");
    el.appendChild(brk);
  }
  el.setAttribute("count", String(lines.length));
  el.setAttribute("manualBreakCount", String(lines.length));
}

/** Write (or clear) one sheet's print setup into its worksheet DOM. */
export function writeXlsxPrintSetup(sheet: Sheet): void {
  const doc = sheet.doc;
  const ws = doc?.documentElement;
  if (!doc || !ws) return;
  const p: PrintSetup | undefined = sheet.printSetup;
  sheet.layoutDirty = true;
  if (!hasPrintSetup(p)) {
    for (const name of ["printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks"]) {
      const el = firstByLocal(ws, name);
      if (el) el.parentNode?.removeChild(el);
    }
    return;
  }
  const setup = p!;

  // <pageSetup>: paper, orientation, scale and the fit-to counts.
  if (setup.orientation || setup.paperSize || setup.scale || setup.fitToWidth != null || setup.fitToHeight != null || setup.pageOrder || setup.firstPageNumber) {
    const el = ensureChild(doc, ws, "pageSetup");
    setOrDrop(el, "paperSize", setup.paperSize != null ? String(setup.paperSize) : undefined);
    setOrDrop(el, "orientation", setup.orientation);
    setOrDrop(el, "scale", setup.scale != null ? String(setup.scale) : undefined);
    setOrDrop(el, "fitToWidth", setup.fitToWidth != null ? String(setup.fitToWidth) : undefined);
    setOrDrop(el, "fitToHeight", setup.fitToHeight != null ? String(setup.fitToHeight) : undefined);
    setOrDrop(el, "pageOrder", setup.pageOrder);
    if (setup.firstPageNumber != null) {
      el.setAttribute("firstPageNumber", String(setup.firstPageNumber));
      el.setAttribute("useFirstPageNumber", "true");
    } else {
      el.removeAttribute("firstPageNumber");
      el.removeAttribute("useFirstPageNumber");
    }
  }
  // fitToPage is what makes the fit-to counts win over scale, and it lives on sheetPr.
  if (setup.fitToPage !== undefined) {
    const sheetPr = ensureChild(doc, ws, "sheetPr");
    let setUpPr = firstByLocal(sheetPr, "pageSetUpPr");
    if (!setUpPr) {
      setUpPr = doc.createElementNS(sheetPr.namespaceURI || SS_MAIN, "pageSetUpPr");
      sheetPr.appendChild(setUpPr);
    }
    setUpPr.setAttribute("fitToPage", setup.fitToPage ? "true" : "false");
  }

  if (setup.margins) {
    const el = ensureChild(doc, ws, "pageMargins");
    for (const side of ["left", "right", "top", "bottom", "header", "footer"] as const) el.setAttribute(side, String(setup.margins[side]));
  }

  if (setup.gridLines !== undefined || setup.headings !== undefined || setup.horizontalCentered !== undefined || setup.verticalCentered !== undefined) {
    const el = ensureChild(doc, ws, "printOptions");
    setOrDrop(el, "gridLines", boolAttr(setup.gridLines));
    setOrDrop(el, "headings", boolAttr(setup.headings));
    setOrDrop(el, "horizontalCentered", boolAttr(setup.horizontalCentered));
    setOrDrop(el, "verticalCentered", boolAttr(setup.verticalCentered));
  }

  const headText = formatHeaderFooter(setup.header), footText = formatHeaderFooter(setup.footer);
  const hfExisting = firstByLocal(ws, "headerFooter");
  if (headText || footText) {
    const el = hfExisting ?? ensureChild(doc, ws, "headerFooter");
    const region = (name: string, text: string | undefined): void => {
      const cur = firstByLocal(el, name);
      if (!text) { if (cur) el.removeChild(cur); return; }
      const target = cur ?? (el.appendChild(doc.createElementNS(el.namespaceURI || SS_MAIN, name)) as Element);
      target.textContent = text;
    };
    region("oddHeader", headText);
    region("oddFooter", footText);
  } else if (hfExisting) {
    hfExisting.parentNode?.removeChild(hfExisting);
  }

  // A break spans the whole of the other axis; the max is that axis's last index.
  writeBreaks(doc, ws, "rowBreaks", setup.rowBreaks, 16383);
  writeBreaks(doc, ws, "colBreaks", setup.colBreaks, 1048575);
}

const absArea = (a: { r1: number; c1: number; r2: number; c2: number }, sheetName: string): string =>
  `${quoteSheet(sheetName)}!$${colToLetters(a.c1)}$${a.r1}:$${colToLetters(a.c2)}$${a.r2}`;

/** A sheet name needs quoting in a reference when it is not a bare identifier. */
export const quoteSheet = (name: string): string =>
  /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;

/** Rewrite the sheet-scoped _xlnm.Print_Area / Print_Titles names in workbook.xml. */
export function writeXlsxPrintNames(wb: Workbook): void {
  if (!wb.sheets.some((s) => s.printDirty)) return;
  const file = wb.files["xl/workbook.xml"];
  if (!file) return;
  const doc = parseXmlOpt(file);
  const root = doc?.documentElement;
  if (!doc || !root) return;
  const ns = root.namespaceURI || SS_MAIN;
  let names = firstByLocal(root, "definedNames");

  const wanted = new Map<string, string>(); // "name|sheetIdx" -> value
  wb.sheets.forEach((sheet, idx) => {
    const p = sheet.printSetup;
    if (p?.printArea?.length) wanted.set(`_xlnm.Print_Area|${idx}`, p.printArea.map((a) => absArea(a, sheet.name)).join(","));
    const parts: string[] = [];
    if (p?.titleCols) parts.push(`${quoteSheet(sheet.name)}!$${colToLetters(p.titleCols.from)}:$${colToLetters(p.titleCols.to)}`);
    if (p?.titleRows) parts.push(`${quoteSheet(sheet.name)}!$${p.titleRows.from}:$${p.titleRows.to}`);
    if (parts.length) wanted.set(`_xlnm.Print_Titles|${idx}`, parts.join(","));
  });

  // Drop every print name belonging to a sheet whose setup changed, then re-add what it now wants.
  const dirty = new Set(wb.sheets.map((s, i) => (s.printDirty ? i : -1)).filter((i) => i >= 0));
  if (names) {
    for (const dn of Array.from(names.children)) {
      const name = dn.getAttribute("name");
      if (name !== "_xlnm.Print_Area" && name !== "_xlnm.Print_Titles") continue;
      if (dirty.has(Number(dn.getAttribute("localSheetId") ?? "-1"))) names.removeChild(dn);
    }
  }
  for (const [k, value] of wanted) {
    const [name, idxStr] = k.split("|");
    const idx = Number(idxStr);
    if (!dirty.has(idx)) continue; // an untouched sheet keeps whatever the file already said
    if (!names) {
      names = doc.createElementNS(ns, "definedNames");
      // CT_Workbook order: definedNames sits after externalReferences and before calcPr.
      const before = Array.from(root.children).find((e) => ["calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "extLst"].includes(e.localName)) ?? null;
      root.insertBefore(names, before);
    }
    const dn = doc.createElementNS(ns, "definedName");
    dn.setAttribute("name", name!);
    dn.setAttribute("localSheetId", String(idx));
    dn.textContent = value;
    names.appendChild(dn);
  }
  if (names && !names.children.length) names.parentNode?.removeChild(names);
  wb.files["xl/workbook.xml"] = serializeXml(doc);
}

/** Persist every sheet whose print setup changed (called from the save path). */
export function writeXlsxPrintSetups(wb: Workbook): void {
  writeXlsxPrintNames(wb); // reads printDirty, so it must run before the flags are cleared
  for (const sheet of wb.sheets) {
    if (!sheet.printDirty) continue;
    writeXlsxPrintSetup(sheet);
    sheet.printDirty = false;
  }
}
