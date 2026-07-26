import { colToLetters, parseXmlOpt, serializeXml, type Sheet, type Workbook } from "../../core/model";
import { DEFAULT_PAPER, MM_PER_INCH, PAPER_SIZES, hasPrintSetup, type HeaderFooter, type PrintSetup } from "../../core/print";
import { ODS } from "./shared";

// ODF splits page setup across two documents: styles.xml holds a <style:page-layout> (paper,
// margins, orientation, scale, what to print) and a <style:master-page> (header and footer), while
// content.xml holds the print ranges on the table plus the page breaks on row / column styles. The
// table finds its master page through its own table style.
//
// Each sheet gets its own layout + master pair, named after the sheet, because ODF has no way to
// give two tables different page setups through one master page.

const IN_TO_MM = MM_PER_INCH;
const mm = (inches: number): string => `${Math.round(inches * IN_TO_MM * 100) / 100}mm`;
/** One line of default header text, as LibreOffice sizes it. Used only for the spacing rule below. */
const NOMINAL_LINE_IN = 0.1389;

const findAll = (doc: Document, local: string): Element[] =>
  Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === local);
const childByLocal = (parent: Element, local: string): Element | undefined =>
  Array.from(parent.children).find((e) => e.localName === local);

/** An XML-name-safe suffix for a sheet, so the generated style names stay valid and unique. */
const styleSuffix = (name: string, index: number): string => {
  const safe = name.replace(/[^A-Za-z0-9_]/g, "_");
  return /^[A-Za-z_]/.test(safe) ? `${safe}_${index}` : `S_${safe}_${index}`;
};

/** Find-or-create a direct child element. */
function ensure(doc: Document, parent: Element, ns: string, qname: string, local: string): Element {
  const found = childByLocal(parent, local);
  if (found) return found;
  const el = doc.createElementNS(ns, qname);
  parent.appendChild(el);
  return el;
}

/** Build a header/footer's three regions from the model's &-coded text. */
function buildHeaderFooter(doc: Document, holder: Element, hf: HeaderFooter | undefined): void {
  while (holder.firstChild) holder.removeChild(holder.firstChild);
  if (!hf) {
    holder.setAttributeNS(ODS.style, "style:display", "false");
    return;
  }
  holder.removeAttribute("style:display");
  for (const [key, local] of [["left", "region-left"], ["center", "region-center"], ["right", "region-right"]] as const) {
    const text = hf[key];
    if (!text) continue;
    const region = doc.createElementNS(ODS.style, `style:${local}`);
    const p = doc.createElementNS(ODS.text, "text:p");
    // Excel's field codes become the ODF field elements; anything else is literal text.
    const FIELDS: Record<string, string> = { P: "text:page-number", N: "text:page-count", D: "text:date", T: "text:time", A: "text:sheet-name", F: "text:file-name" };
    let buf = "";
    const flush = () => { if (buf) { p.appendChild(doc.createTextNode(buf)); buf = ""; } };
    for (let i = 0; i < text.length; i++) {
      if (text[i] === "&" && i + 1 < text.length) {
        const c = text[i + 1]!;
        if (c === "&") { buf += "&"; i++; continue; }
        const field = FIELDS[c.toUpperCase()];
        if (field) {
          flush();
          const f = doc.createElementNS(ODS.text, field);
          // The field elements carry a placeholder value that the reader replaces.
          f.textContent = c.toUpperCase() === "P" ? "1" : c.toUpperCase() === "N" ? "1" : "";
          p.appendChild(f);
          i++;
          continue;
        }
        // A formatting code (&"font"&12&B): ODF states those as span styles, which are not
        // modelled here, so the code is dropped rather than printed as literal text.
        if (c === '"') { const end = text.indexOf('"', i + 2); i = end === -1 ? text.length : end; continue; }
        if (/[0-9]/.test(c)) { let j = i + 1; while (j < text.length && /[0-9]/.test(text[j]!)) j++; i = j - 1; continue; }
        if (/[BIULSXYEOZG]/i.test(c)) { i++; continue; }
      }
      buf += text[i];
    }
    flush();
    region.appendChild(p);
    holder.appendChild(region);
  }
}

/** Create or update this sheet's page-layout + master-page pair in styles.xml. */
function writeLayout(doc: Document, suffix: string, p: PrintSetup): string {
  const autoStyles = findAll(doc, "automatic-styles")[0] ?? (() => {
    const el = doc.createElementNS(ODS.office, "office:automatic-styles");
    doc.documentElement.insertBefore(el, doc.documentElement.firstChild);
    return el;
  })();
  const masterStyles = findAll(doc, "master-styles")[0] ?? (doc.documentElement.appendChild(doc.createElementNS(ODS.office, "office:master-styles")) as Element);

  const layoutName = `PMse_${suffix}`;
  const masterName = `PMse_master_${suffix}`;
  const layout = findAll(doc, "page-layout").find((e) => e.getAttribute("style:name") === layoutName)
    ?? (() => { const el = doc.createElementNS(ODS.style, "style:page-layout"); el.setAttributeNS(ODS.style, "style:name", layoutName); autoStyles.appendChild(el); return el; })();
  const props = ensure(doc, layout, ODS.style, "style:page-layout-properties", "page-layout-properties");

  const paper = PAPER_SIZES[p.paperSize ?? DEFAULT_PAPER] ?? PAPER_SIZES[DEFAULT_PAPER]!;
  const landscape = p.orientation === "landscape";
  props.setAttributeNS(ODS.fo, "fo:page-width", `${landscape ? paper.h : paper.w}mm`);
  props.setAttributeNS(ODS.fo, "fo:page-height", `${landscape ? paper.w : paper.h}mm`);
  props.setAttributeNS(ODS.style, "style:print-orientation", landscape ? "landscape" : "portrait");

  // ODF's page margin stops at the header block, xlsx's encloses it, so the reserved block height
  // is the difference and the page margin is the header/footer distance.
  const m = p.margins;
  if (m) {
    props.setAttributeNS(ODS.fo, "fo:margin-left", mm(m.left));
    props.setAttributeNS(ODS.fo, "fo:margin-right", mm(m.right));
    props.setAttributeNS(ODS.fo, "fo:margin-top", mm(m.header));
    props.setAttributeNS(ODS.fo, "fo:margin-bottom", mm(m.footer));
    const headBlock = ensure(doc, layout, ODS.style, "style:header-style", "header-style");
    const footBlock = ensure(doc, layout, ODS.style, "style:footer-style", "footer-style");
    const hp = ensure(doc, headBlock, ODS.style, "style:header-footer-properties", "header-footer-properties");
    const fp = ensure(doc, footBlock, ODS.style, "style:header-footer-properties", "header-footer-properties");
    const headBlockIn = Math.max(0, m.top - m.header), footBlockIn = Math.max(0, m.bottom - m.footer);
    hp.setAttributeNS(ODS.fo, "fo:min-height", mm(headBlockIn));
    fp.setAttributeNS(ODS.fo, "fo:min-height", mm(footBlockIn));
    // LibreOffice derives an OOXML top margin from page-margin + header CONTENT height + the
    // header's own margin-bottom, ignoring min-height, and defaults that margin to 20mm when it is
    // absent - which inflates the margin by ~0.5in on every ods -> xlsx conversion. Stating the
    // spacing the way LibreOffice states it in its own files (the block height less one nominal
    // text line) makes the margin survive that conversion exactly. Verified at 0.75in top.
    for (const [el, side, block] of [[hp, "fo:margin-bottom", headBlockIn], [fp, "fo:margin-top", footBlockIn]] as const) {
      el.setAttributeNS(ODS.fo, "fo:margin-left", "0mm");
      el.setAttributeNS(ODS.fo, "fo:margin-right", "0mm");
      el.setAttributeNS(ODS.fo, side, mm(Math.max(0, block - NOMINAL_LINE_IN)));
    }
  }

  if (p.pageOrder) props.setAttributeNS(ODS.style, "style:print-page-order", p.pageOrder === "overThenDown" ? "ltr" : "ttb");
  props.setAttributeNS(ODS.style, "style:first-page-number", p.firstPageNumber != null ? String(p.firstPageNumber) : "continue");

  // Fit-to wins over scale, exactly as xlsx's fitToPage decides it.
  props.removeAttribute("style:scale-to");
  props.removeAttribute("style:scale-to-X");
  props.removeAttribute("style:scale-to-Y");
  if (p.fitToPage && (p.fitToWidth || p.fitToHeight)) {
    if (p.fitToWidth) props.setAttributeNS(ODS.style, "style:scale-to-X", String(p.fitToWidth));
    if (p.fitToHeight) props.setAttributeNS(ODS.style, "style:scale-to-Y", String(p.fitToHeight));
  } else if (p.scale != null && p.scale !== 100) {
    props.setAttributeNS(ODS.style, "style:scale-to", `${p.scale}%`);
  }

  const centering = p.horizontalCentered && p.verticalCentered ? "both" : p.horizontalCentered ? "horizontal" : p.verticalCentered ? "vertical" : "none";
  props.setAttributeNS(ODS.style, "style:table-centering", centering);
  // style:print lists what goes on the paper; the two flags modelled here join the usual set.
  const parts = ["charts", "drawings", "objects", "zero-values"];
  if (p.gridLines) parts.push("grid");
  if (p.headings) parts.push("headers");
  props.setAttributeNS(ODS.style, "style:print", parts.sort().join(" "));

  const master = findAll(doc, "master-page").find((e) => e.getAttribute("style:name") === masterName)
    ?? (() => { const el = doc.createElementNS(ODS.style, "style:master-page"); el.setAttributeNS(ODS.style, "style:name", masterName); masterStyles.appendChild(el); return el; })();
  master.setAttributeNS(ODS.style, "style:page-layout-name", layoutName);
  buildHeaderFooter(doc, ensure(doc, master, ODS.style, "style:header", "header"), p.header);
  buildHeaderFooter(doc, ensure(doc, master, ODS.style, "style:footer", "footer"), p.footer);
  return masterName;
}

/** Point this sheet's table style at the master page, creating a table style when it has none. */
function bindMasterPage(wb: Workbook, sheet: Sheet, masterName: string): void {
  const doc = wb.contentDoc;
  const table = sheet.tableEl;
  if (!doc || !table) return;
  const autoStyles = findAll(doc, "automatic-styles")[0] ?? (() => {
    const el = doc.createElementNS(ODS.office, "office:automatic-styles");
    doc.documentElement.insertBefore(el, doc.documentElement.firstChild);
    return el;
  })();
  let name = table.getAttribute("table:style-name");
  let style = name ? findAll(doc, "style").find((e) => e.getAttribute("style:name") === name && e.getAttribute("style:family") === "table") : undefined;
  if (!style) {
    name = `ta_se_${masterName}`;
    style = findAll(doc, "style").find((e) => e.getAttribute("style:name") === name);
    if (!style) {
      style = doc.createElementNS(ODS.style, "style:style");
      style.setAttributeNS(ODS.style, "style:name", name);
      style.setAttributeNS(ODS.style, "style:family", "table");
      const props = doc.createElementNS(ODS.style, "style:table-properties");
      props.setAttributeNS(ODS.table, "table:display", "true");
      style.appendChild(props);
      autoStyles.appendChild(style);
    }
    table.setAttributeNS(ODS.table, "table:style-name", name);
  }
  style.setAttributeNS(ODS.style, "style:master-page-name", masterName);
}

/** Give a line's style a page break, cloning the style so other lines keep theirs. */
function applyBreaks(wb: Workbook, sheet: Sheet, p: PrintSetup): void {
  const doc = wb.contentDoc;
  const table = sheet.tableEl;
  if (!doc || !table) return;
  const autoStyles = findAll(doc, "automatic-styles")[0];
  if (!autoStyles) return;
  const styles = new Map<string, Element>();
  for (const st of findAll(doc, "style")) { const n = st.getAttribute("style:name"); if (n) styles.set(n, st); }

  /** A style like `base` but carrying (or not) fo:break-before="page". */
  const variantOf = (base: string | null, wantBreak: boolean, family: "table-row" | "table-column", prefix: string): string | null => {
    const src = base ? styles.get(base) : undefined;
    const propsLocal = family === "table-row" ? "table-row-properties" : "table-column-properties";
    const has = src ? childByLocal(src, propsLocal)?.getAttribute("fo:break-before") === "page" : false;
    if (has === wantBreak) return base;
    const name = `${prefix}${wantBreak ? "brk" : "nob"}_${base ?? "d"}`;
    if (!styles.has(name)) {
      const el = (src ? src.cloneNode(true) : doc.createElementNS(ODS.style, "style:style")) as Element;
      el.setAttributeNS(ODS.style, "style:name", name);
      el.setAttributeNS(ODS.style, "style:family", family);
      const props = ensure(doc, el, ODS.style, `style:${propsLocal}`, propsLocal);
      if (wantBreak) props.setAttributeNS(ODS.fo, "fo:break-before", "page");
      else props.removeAttribute("fo:break-before");
      autoStyles.appendChild(el);
      styles.set(name, el);
    }
    return name;
  };

  const rowBreaks = new Set(p.rowBreaks ?? []);
  // A sheet whose rows carry no style has no map yet, and a break is the reason to start one.
  const rowStyles = rowBreaks.size ? (sheet.odsRowStyles ??= new Map()) : sheet.odsRowStyles;
  if (rowStyles) {
    // Every row that has (or should lose) a break gets the matching style variant.
    const touched = new Set<number>([...rowBreaks, ...rowStyles.keys()]);
    for (const r of touched) {
      const want = rowBreaks.has(r);
      const base = rowStyles.get(r) ?? null;
      if (!want && !base) continue;
      const next = variantOf(base, want, "table-row", "rose_");
      if (next) rowStyles.set(r, next);
    }
  }
  const colBreaks = new Set(p.colBreaks ?? []);
  let col = 1;
  for (const ch of Array.from(table.children)) {
    if (ch.localName !== "table-column") continue;
    const n = Math.max(1, Number(ch.getAttribute("table:number-columns-repeated") || "1"));
    // A run whose first column takes a break has to be split, so the break lands on that column only.
    const wants = [...colBreaks].filter((c) => c >= col && c < col + n);
    const base = ch.getAttribute("table:style-name");
    if (wants.length === 1 && wants[0] === col && n === 1) {
      const next = variantOf(base, true, "table-column", "cose_");
      if (next) ch.setAttributeNS(ODS.table, "table:style-name", next);
    } else if (!wants.length) {
      const next = variantOf(base, false, "table-column", "cose_");
      if (next && next !== base) ch.setAttributeNS(ODS.table, "table:style-name", next);
    } else {
      // Split the run so each break can sit on its own column element.
      const parts: { from: number; count: number }[] = [];
      let at = col;
      for (const c of wants.sort((a, b) => a - b)) {
        if (c > at) parts.push({ from: at, count: c - at });
        parts.push({ from: c, count: 1 });
        at = c + 1;
      }
      if (at < col + n) parts.push({ from: at, count: col + n - at });
      const made: Element[] = [];
      for (const part of parts) {
        const el = ch.cloneNode(true) as Element;
        if (part.count > 1) el.setAttributeNS(ODS.table, "table:number-columns-repeated", String(part.count));
        else el.removeAttribute("table:number-columns-repeated");
        const next = variantOf(base, colBreaks.has(part.from) && part.count === 1, "table-column", "cose_");
        if (next) el.setAttributeNS(ODS.table, "table:style-name", next);
        made.push(el);
      }
      for (const el of made) ch.parentNode?.insertBefore(el, ch);
      ch.parentNode?.removeChild(ch);
    }
    col += n;
  }
  sheet.odsDirty = true;
}

/** Persist every sheet whose print setup changed. */
export function writeOdsPrintSetups(wb: Workbook): void {
  const dirty = wb.sheets.filter((s) => s.printDirty);
  if (!dirty.length) return;
  const stylesFile = wb.files["styles.xml"];
  const stylesDoc = stylesFile ? parseXmlOpt(stylesFile) : undefined;

  for (const sheet of wb.sheets) {
    if (!sheet.printDirty) continue;
    sheet.printDirty = false;
    const table = sheet.tableEl;
    const p = sheet.printSetup;
    if (!table) continue;

    if (hasPrintSetup(p) && p!.printArea?.length) {
      table.setAttributeNS(ODS.table, "table:print-ranges", p!.printArea.map((a) => `${sheet.name}.${colToLetters(a.c1)}${a.r1}:${sheet.name}.${colToLetters(a.c2)}${a.r2}`).join(" "));
    } else {
      table.removeAttribute("table:print-ranges");
    }
    if (hasPrintSetup(p)) {
      applyBreaks(wb, sheet, p!);
      if (stylesDoc) {
        const idx = wb.sheets.indexOf(sheet);
        bindMasterPage(wb, sheet, writeLayout(stylesDoc, styleSuffix(sheet.name, idx), p!));
      }
    }
    sheet.odsDirty = true;
  }
  if (stylesDoc) wb.files["styles.xml"] = serializeXml(stylesDoc);
}
