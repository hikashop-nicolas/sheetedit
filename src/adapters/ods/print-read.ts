import { lettersToCol, parseXmlOpt, type Workbook } from "../../core/model";
import { paperSizeFor, type HeaderFooter, type PrintSetup } from "../../core/print";

// ODF states page setup as a <style:page-layout> (paper, margins, orientation, scale) paired with a
// <style:master-page> (the header and footer), both in styles.xml. A table reaches its master page
// through its own table style's style:master-page-name.
//
// The rest is on the table itself: table:print-ranges for the print area, <table:table-header-rows>
// for the rows repeated on each page, and fo:break-before="page" on a row or column style for a
// manual break.

const MM_PER_IN = 25.4;

/** An ODF length ("8.2681in", "210mm", "21cm") in millimetres. */
export function odsLenToMm(v: string | null | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^(-?[\d.]+)\s*(mm|cm|in|pt|pc)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  switch (m[2]) {
    case "cm": return n * 10;
    case "in": return n * MM_PER_IN;
    case "pt": return (n / 72) * MM_PER_IN;
    case "pc": return (n / 6) * MM_PER_IN;
    default: return n; // mm
  }
}
const mmToIn = (mm: number): number => Math.round((mm / MM_PER_IN) * 10000) / 10000;

/** Flatten one header/footer region into the &-code text the model stores. */
function regionText(region: Element | undefined): string | undefined {
  if (!region) return undefined;
  let out = "";
  const walk = (node: Node): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3) { out += child.nodeValue ?? ""; continue; }
      if (child.nodeType !== 1) continue;
      const el = child as Element;
      switch (el.localName) {
        // The field elements both formats have in common, back into Excel's codes.
        case "page-number": out += "&P"; break;
        case "page-count": out += "&N"; break;
        case "date": out += "&D"; break;
        case "time": out += "&T"; break;
        case "sheet-name": out += "&A"; break;
        case "file-name": case "title": out += "&F"; break;
        case "s": out += " "; break;
        case "tab": out += "\t"; break;
        default: walk(el);
      }
    }
  };
  walk(region);
  return out || undefined;
}

/** The three regions of a <style:header> / <style:footer>. */
function readHeaderFooter(el: Element | undefined): HeaderFooter | undefined {
  if (!el) return undefined;
  const region = (name: string): Element | undefined => Array.from(el.children).find((e) => e.localName === name);
  const left = regionText(region("region-left"));
  const center = regionText(region("region-center"));
  const right = regionText(region("region-right"));
  // A header with no explicit regions is one centred paragraph.
  if (!left && !center && !right) {
    const plain = regionText(el);
    return plain ? { center: plain } : undefined;
  }
  const hf: HeaderFooter = {};
  if (left) hf.left = left;
  if (center) hf.center = center;
  if (right) hf.right = right;
  return hf;
}

/** ODF page setup for every sheet, from styles.xml plus the table elements in content.xml. */
export function readOdsPrintSetup(wb: Workbook, files: Record<string, Uint8Array>): void {
  const stylesDoc = files["styles.xml"] ? parseXmlOpt(files["styles.xml"]) : undefined;
  const contentDoc = wb.contentDoc;
  const byLocal = (doc: Document | undefined, local: string): Element[] =>
    doc ? Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === local) : [];

  // page-layout name -> its properties, and master-page name -> (layout name, header, footer).
  const layouts = new Map<string, Element>();
  for (const l of byLocal(stylesDoc, "page-layout")) {
    const name = l.getAttribute("style:name");
    const props = Array.from(l.children).find((e) => e.localName === "page-layout-properties");
    if (name && props) layouts.set(name, props);
  }
  const masters = new Map<string, { layout?: string; header?: Element; footer?: Element }>();
  for (const m of byLocal(stylesDoc, "master-page")) {
    const name = m.getAttribute("style:name");
    if (!name) continue;
    masters.set(name, {
      layout: m.getAttribute("style:page-layout-name") ?? undefined,
      header: Array.from(m.children).find((e) => e.localName === "header"),
      footer: Array.from(m.children).find((e) => e.localName === "footer"),
    });
  }
  // table style name -> master page it points at (the styles can live in either document).
  const tableMaster = new Map<string, string>();
  for (const doc of [contentDoc, stylesDoc]) {
    for (const st of byLocal(doc, "style")) {
      if (st.getAttribute("style:family") !== "table") continue;
      const name = st.getAttribute("style:name"), mp = st.getAttribute("style:master-page-name");
      if (name && mp) tableMaster.set(name, mp);
    }
  }
  // row / column style name -> whether it starts a new page.
  const breakStyles = new Set<string>();
  for (const doc of [contentDoc, stylesDoc]) {
    for (const st of byLocal(doc, "style")) {
      const fam = st.getAttribute("style:family");
      if (fam !== "table-row" && fam !== "table-column") continue;
      const props = Array.from(st.children).find((e) => e.localName.endsWith("-properties"));
      if (props?.getAttribute("fo:break-before") !== "page") continue;
      const name = st.getAttribute("style:name");
      if (name) breakStyles.add(name);
    }
  }

  for (const sheet of wb.sheets) {
    const table = sheet.tableEl;
    if (!table) continue;
    const p: PrintSetup = {};
    const master = masters.get(tableMaster.get(table.getAttribute("table:style-name") ?? "") ?? "");
    const props = master?.layout ? layouts.get(master.layout) : undefined;
    if (props) {
      const orient = props.getAttribute("style:print-orientation");
      if (orient === "landscape" || orient === "portrait") p.orientation = orient;
      const w = odsLenToMm(props.getAttribute("fo:page-width")), h = odsLenToMm(props.getAttribute("fo:page-height"));
      if (w != null && h != null) p.paperSize = paperSizeFor(w, h);
      // ODF's page margins sit inside the header/footer blocks, xlsx's enclose them, so the top and
      // bottom are the ODF margin plus that block's reserved height.
      const mm = (n: string): number | undefined => odsLenToMm(props.getAttribute(n));
      // The reserved height sits on the sibling <style:header-style> / <style:footer-style>.
      const blockMm = (styleLocal: string): number => {
        const holder = Array.from(props.parentElement?.children ?? []).find((e) => e.localName === styleLocal);
        const hp = holder ? Array.from(holder.children).find((e) => e.localName === "header-footer-properties") : undefined;
        return odsLenToMm(hp?.getAttribute("fo:min-height")) ?? 0;
      };
      const top = mm("fo:margin-top"), bottom = mm("fo:margin-bottom"), left = mm("fo:margin-left"), right = mm("fo:margin-right");
      if (top != null || bottom != null || left != null || right != null) {
        const headH = blockMm("header-style"), footH = blockMm("footer-style");
        p.margins = {
          left: mmToIn(left ?? 0),
          right: mmToIn(right ?? 0),
          top: mmToIn((top ?? 0) + headH),
          bottom: mmToIn((bottom ?? 0) + footH),
          header: mmToIn(top ?? 0),
          footer: mmToIn(bottom ?? 0),
        };
      }
      const order = props.getAttribute("style:print-page-order");
      if (order) p.pageOrder = order === "ltr" ? "overThenDown" : "downThenOver";
      const first = props.getAttribute("style:first-page-number");
      if (first && first !== "continue") { const n = Number(first); if (Number.isFinite(n) && n !== 1) p.firstPageNumber = n; }
      const scaleTo = props.getAttribute("style:scale-to");
      if (scaleTo) { const n = Number(scaleTo.replace("%", "")); if (Number.isFinite(n)) p.scale = n; }
      const fitX = props.getAttribute("style:scale-to-X"), fitY = props.getAttribute("style:scale-to-Y");
      const pages = props.getAttribute("style:scale-to-pages");
      if (fitX || fitY || pages) {
        p.fitToPage = true;
        if (fitX) p.fitToWidth = Number(fitX) || 0;
        if (fitY) p.fitToHeight = Number(fitY) || 0;
        // scale-to-pages constrains the total, which xlsx cannot state; take it as pages across.
        if (!fitX && !fitY && pages) p.fitToWidth = Number(pages) || 0;
        if (p.fitToWidth === undefined) p.fitToWidth = 0;
        if (p.fitToHeight === undefined) p.fitToHeight = 0;
      }
      const centering = props.getAttribute("style:table-centering");
      if (centering) {
        p.horizontalCentered = centering === "horizontal" || centering === "both";
        p.verticalCentered = centering === "vertical" || centering === "both";
      }
      // style:print is a space-separated list of what gets printed.
      const printList = props.getAttribute("style:print");
      if (printList != null) {
        const parts = printList.split(/\s+/);
        p.gridLines = parts.includes("grid");
        p.headings = parts.includes("headers");
      }
    }
    const header = readHeaderFooter(master?.header);
    const footer = readHeaderFooter(master?.footer);
    if (header) p.header = header;
    if (footer) p.footer = footer;

    // Print ranges: "Sheet1.A1:Sheet1.C5 Sheet1.E1:Sheet1.E9" (space separated).
    const ranges = table.getAttribute("table:print-ranges");
    if (ranges) {
      const areas = ranges.split(/\s+/).map(parseOdsRange).filter((r): r is NonNullable<typeof r> => !!r);
      if (areas.length) p.printArea = areas;
    }
    // Repeated rows are the header-rows group the reader already located.
    if (sheet.odsHeaderRows) p.titleRows = { from: sheet.odsHeaderRows.from, to: sheet.odsHeaderRows.to };
    const rowBreaks = linesWithBreak(sheet.odsRowStyles, breakStyles);
    if (rowBreaks.length) p.rowBreaks = rowBreaks;
    const colBreaks = columnBreaks(table, breakStyles);
    if (colBreaks.length) p.colBreaks = colBreaks;

    for (const k of Object.keys(p) as (keyof PrintSetup)[]) if (p[k] === undefined) delete p[k];
    if (Object.keys(p).length) sheet.printSetup = p;
  }
}

/** The 1-based columns whose <table:table-column> style carries a page break. Columns are stated as
    runs, so the repeat count has to be walked rather than indexed. */
function columnBreaks(table: Element, breakStyles: Set<string>): number[] {
  if (!breakStyles.size) return [];
  const out: number[] = [];
  let col = 1;
  const walk = (parent: Element): void => {
    for (const ch of Array.from(parent.children)) {
      if (ch.localName === "table-header-columns" || ch.localName === "table-columns" || ch.localName === "table-column-group") { walk(ch); continue; }
      if (ch.localName !== "table-column") continue;
      const n = Math.max(1, Number(ch.getAttribute("table:number-columns-repeated") || "1"));
      const name = ch.getAttribute("table:style-name");
      // Only the first column of a repeated run starts the page; the run shares one style.
      if (name && breakStyles.has(name)) out.push(col);
      col += n;
    }
  };
  walk(table);
  return out.sort((a, b) => a - b);
}

/** The 1-based lines whose style carries a page break. */
function linesWithBreak(styles: Map<number, string> | undefined, breakStyles: Set<string>): number[] {
  if (!styles || !breakStyles.size) return [];
  const out: number[] = [];
  for (const [line, name] of styles) if (breakStyles.has(name)) out.push(line);
  return out.sort((a, b) => a - b);
}

/** "Sheet1.A1:Sheet1.C5" or ".A1:.C5" -> a 1-based inclusive range. */
export function parseOdsRange(ref: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const cell = (part: string): { row: number; col: number } | null => {
    const body = part.slice(part.indexOf(".") + 1).replace(/\$/g, "");
    const m = /^([A-Za-z]+)(\d+)$/.exec(body);
    return m ? { col: lettersToCol(m[1]!.toUpperCase()), row: Number(m[2]) } : null;
  };
  const [a, b] = ref.split(":");
  const p1 = cell(a ?? ""), p2 = b ? cell(b) : p1;
  if (!p1 || !p2) return null;
  return { r1: Math.min(p1.row, p2.row), c1: Math.min(p1.col, p2.col), r2: Math.max(p1.row, p2.row), c2: Math.max(p1.col, p2.col) };
}
