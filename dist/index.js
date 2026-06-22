import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import FormulaParser from "fast-formula-parser";
const key = (row, col) => `${row}:${col}`;
function getCell(sheet, row, col) {
    return sheet.cells.get(key(row, col));
}
function ensureCell(sheet, row, col) {
    let c = sheet.cells.get(key(row, col));
    if (!c) {
        c = { row, col, value: "", kind: "blank" };
        sheet.cells.set(key(row, col), c);
    }
    if (row > sheet.maxRow)
        sheet.maxRow = row;
    if (col > sheet.maxCol)
        sheet.maxCol = col;
    return c;
}
function noteExtent(sheet, row, col) {
    if (row > sheet.maxRow)
        sheet.maxRow = row;
    if (col > sheet.maxCol)
        sheet.maxCol = col;
}
/** Typed value of a cell for the formula engine: number, boolean, string or null. */
function typedValue(cell) {
    if (!cell)
        return null;
    if (cell.value === "")
        return cell.formula != null ? null : null;
    switch (cell.kind) {
        case "n": {
            const n = Number(cell.value);
            return Number.isFinite(n) ? n : null;
        }
        case "b":
            return cell.value === "TRUE" || cell.value === "1" || cell.value === "true";
        case "blank":
            return null;
        default:
            return cell.value;
    }
}
const isNumeric = (s) => {
    const t = s.trim();
    return t !== "" && /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t);
};
function numToStr(n) {
    if (Number.isInteger(n))
        return String(n);
    return String(parseFloat(n.toPrecision(15)));
}
/** Apply a number format (code or built-in id) to a numeric value via SSF. */
function formatNumber(fmt, value) {
    const n = Number(value);
    if (!Number.isFinite(n))
        return undefined;
    try {
        return FormulaParser.SSF.format(fmt, n);
    }
    catch {
        return undefined;
    }
}
/** Text shown in the grid for a cell: the formatted display, else the raw value. */
const cellDisplay = (cell) => (cell ? cell.display ?? cell.value : "");
// ---------------------------------------------------------------------------
// A1 helpers
// ---------------------------------------------------------------------------
export function colToLetters(col) {
    let s = "";
    while (col > 0) {
        const m = (col - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        col = Math.floor((col - 1) / 26);
    }
    return s;
}
function lettersToCol(letters) {
    let n = 0;
    for (const ch of letters)
        n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
}
function parseA1Ref(ref) {
    const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
    if (!m)
        return null;
    return { col: lettersToCol(m[1]), row: Number(m[2]) };
}
// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------
const parseXml = (file) => new DOMParser().parseFromString(strFromU8(file), "application/xml");
function serializeXml(doc) {
    let s = new XMLSerializer().serializeToString(doc);
    if (!s.startsWith("<?xml"))
        s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
    return strToU8(s);
}
const firstByLocal = (parent, local) => {
    for (const ch of Array.from(parent.children))
        if (ch.localName === local)
            return ch;
    return undefined;
};
const removeByLocal = (parent, local) => {
    for (const ch of Array.from(parent.children))
        if (ch.localName === local)
            parent.removeChild(ch);
};
// ---------------------------------------------------------------------------
// xlsx (OOXML SpreadsheetML)
// ---------------------------------------------------------------------------
const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
function readSharedStrings(file) {
    if (!file)
        return [];
    const doc = parseXml(file);
    return Array.from(doc.getElementsByTagName("si")).map((si) => Array.from(si.getElementsByTagName("t"))
        .map((t) => t.textContent ?? "")
        .join(""));
}
// ARGB ("FFRRGGBB" or "RRGGBB") -> CSS "#rrggbb".
function argbToCss(argb) {
    if (!argb)
        return undefined;
    const h = argb.length === 8 ? argb.slice(2) : argb;
    return /^[0-9a-fA-F]{6}$/.test(h) ? "#" + h.toLowerCase() : undefined;
}
// Excel tint: negative darkens toward black, positive lightens toward white.
function applyTint(hex, tint) {
    const ch = (i) => {
        const c = parseInt(hex.slice(i, i + 2), 16);
        const v = tint < 0 ? c * (1 + tint) : c * (1 - tint) + 255 * tint;
        return Math.max(0, Math.min(255, Math.round(v)))
            .toString(16)
            .padStart(2, "0");
    };
    return "#" + ch(1) + ch(3) + ch(5);
}
const findByLocal = (doc, local) => Array.from(doc.getElementsByTagName("*")).find((e) => e.localName === local);
// theme1.xml <clrScheme> -> array indexed by a <color theme="N"> index.
function readTheme(file) {
    const fallback = ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#0563c1", "#954f72"];
    if (!file)
        return fallback;
    try {
        const scheme = findByLocal(parseXml(file), "clrScheme");
        if (!scheme)
            return fallback;
        const byName = {};
        for (const el of Array.from(scheme.children)) {
            const c = el.firstElementChild;
            const css = c && (c.localName === "srgbClr" ? argbToCss(c.getAttribute("val")) : argbToCss(c.getAttribute("lastClr")));
            if (el.localName && css)
                byName[el.localName] = css;
        }
        // theme index order swaps dk/lt 1 and 2 vs the clrScheme element order.
        const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
        return order.map((nm, i) => byName[nm] ?? fallback[i]);
    }
    catch {
        return fallback;
    }
}
// Resolve a <color> element (rgb, or theme + tint) to a CSS colour.
function resolveColor(el, theme) {
    if (!el)
        return undefined;
    const rgb = el.getAttribute("rgb");
    if (rgb)
        return argbToCss(rgb);
    const t = el.getAttribute("theme");
    if (t != null) {
        const base = theme[Number(t)] ?? "#000000";
        const tint = Number(el.getAttribute("tint") || "0");
        return tint ? applyTint(base, tint) : base;
    }
    return undefined;
}
function readXlsxStyles(doc, theme) {
    const customFmt = new Map();
    const xfNumFmtIds = [];
    const xfStyles = [];
    if (!doc)
        return { customFmt, xfNumFmtIds, xfStyles };
    for (const nf of Array.from(doc.getElementsByTagName("numFmt"))) {
        const id = Number(nf.getAttribute("numFmtId"));
        const code = nf.getAttribute("formatCode");
        if (Number.isFinite(id) && code != null)
            customFmt.set(id, code);
    }
    const pool = (local) => {
        const parent = firstByLocal(doc.documentElement, local);
        return parent ? Array.from(parent.children).filter((e) => e.localName === local.replace(/s$/, "")) : [];
    };
    const fonts = pool("fonts").map((f) => ({
        bold: !!firstByLocal(f, "b"),
        italic: !!firstByLocal(f, "i"),
        color: resolveColor(firstByLocal(f, "color"), theme),
    }));
    const fills = pool("fills").map((fl) => {
        const pat = firstByLocal(fl, "patternFill");
        return pat?.getAttribute("patternType") === "solid" ? resolveColor(firstByLocal(pat, "fgColor"), theme) : undefined;
    });
    const borders = pool("borders").map((bd) => {
        const side = (name) => {
            const s = firstByLocal(bd, name);
            return s?.getAttribute("style") ? (resolveColor(firstByLocal(s, "color"), theme) ?? "#444") : undefined;
        };
        const b = { top: side("top"), right: side("right"), bottom: side("bottom"), left: side("left") };
        return b.top || b.right || b.bottom || b.left ? b : undefined;
    });
    // The cell @s indexes <cellXfs>, not <cellStyleXfs>; read that list specifically.
    const cellXfs = doc.getElementsByTagName("cellXfs")[0];
    if (cellXfs) {
        for (const xf of Array.from(cellXfs.children)) {
            if (xf.localName !== "xf")
                continue;
            xfNumFmtIds.push(Number(xf.getAttribute("numFmtId") || "0"));
            const st = {};
            const font = fonts[Number(xf.getAttribute("fontId") || "0")];
            if (font) {
                if (font.bold)
                    st.bold = true;
                if (font.italic)
                    st.italic = true;
                if (font.color)
                    st.color = font.color;
            }
            const fill = fills[Number(xf.getAttribute("fillId") || "0")];
            if (fill)
                st.bg = fill;
            const border = borders[Number(xf.getAttribute("borderId") || "0")];
            if (border)
                st.borders = border;
            const align = firstByLocal(xf, "alignment")?.getAttribute("horizontal");
            if (align === "center" || align === "right" || align === "left")
                st.align = align;
            xfStyles.push(Object.keys(st).length ? st : undefined);
        }
    }
    return { customFmt, xfNumFmtIds, xfStyles };
}
/** Resolve a cell's number format (code or built-in id), or undefined for General. */
function resolveXlsxFmt(styles, s) {
    if (s == null)
        return undefined;
    const numFmtId = styles.xfNumFmtIds[Number(s)];
    if (numFmtId == null || numFmtId === 0)
        return undefined; // 0 = General
    const custom = styles.customFmt.get(numFmtId);
    if (custom != null)
        return custom === "General" ? undefined : custom;
    return numFmtId; // built-in id; SSF resolves it
}
function readXlsx(files) {
    const wb = { kind: "xlsx", sheets: [], files };
    const wbXml = files["xl/workbook.xml"];
    if (!wbXml)
        throw new Error("not an .xlsx: xl/workbook.xml missing");
    const wbDoc = parseXml(wbXml);
    const rels = new Map();
    const relsFile = files["xl/_rels/workbook.xml.rels"];
    if (relsFile) {
        for (const r of Array.from(parseXml(relsFile).getElementsByTagName("Relationship"))) {
            const id = r.getAttribute("Id");
            const target = r.getAttribute("Target");
            if (id && target)
                rels.set(id, target);
        }
    }
    const shared = readSharedStrings(files["xl/sharedStrings.xml"]);
    const theme = readTheme(files["xl/theme/theme1.xml"]);
    wb.stylesDoc = files["xl/styles.xml"] ? parseXml(files["xl/styles.xml"]) : undefined;
    const styles = readXlsxStyles(wb.stylesDoc, theme);
    let n = 0;
    for (const sheetEl of Array.from(wbDoc.getElementsByTagName("sheet"))) {
        n++;
        const name = sheetEl.getAttribute("name") ?? `Sheet${n}`;
        const rid = sheetEl.getAttribute("r:id") ?? sheetEl.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        let target = (rid && rels.get(rid)) || `worksheets/sheet${n}.xml`;
        const path = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
        const wsFile = files[path];
        const sheet = { name, cells: new Map(), maxRow: 0, maxCol: 0, path };
        if (wsFile) {
            const doc = parseXml(wsFile);
            const sheetData = doc.getElementsByTagName("sheetData")[0];
            sheet.doc = doc;
            sheet.sheetData = sheetData;
            // Column widths: <cols><col min max width/></cols>. Width is in character units;
            // convert to px (~7px per char + padding for the default font).
            const colsEl = doc.getElementsByTagName("cols")[0];
            if (colsEl) {
                const cw = new Map();
                for (const col of Array.from(colsEl.children)) {
                    if (col.localName !== "col")
                        continue;
                    const min = Number(col.getAttribute("min") || "0");
                    const max = Number(col.getAttribute("max") || "0");
                    const width = Number(col.getAttribute("width") || "0");
                    if (!min || !width)
                        continue;
                    const px = Math.round(width * 7 + 5);
                    for (let c = min; c <= Math.min(max || min, min + 1000); c++)
                        cw.set(c, px);
                }
                if (cw.size)
                    sheet.colWidths = cw;
            }
            // Row heights: <row r ht customHeight/>. ht is in points; convert to px (~4/3 px/pt).
            if (sheetData) {
                const rh = new Map();
                for (const rowEl of Array.from(sheetData.children)) {
                    if (rowEl.localName !== "row")
                        continue;
                    const r = Number(rowEl.getAttribute("r") || "0");
                    const ht = Number(rowEl.getAttribute("ht") || "0");
                    if (r && ht)
                        rh.set(r, Math.round((ht * 4) / 3));
                }
                if (rh.size)
                    sheet.rowHeights = rh;
            }
            // Merged ranges: <mergeCells><mergeCell ref="B1:C1"/></mergeCells>.
            const mergeEls = doc.getElementsByTagName("mergeCell");
            if (mergeEls.length) {
                const merges = [];
                for (const m of Array.from(mergeEls)) {
                    const ref = m.getAttribute("ref");
                    const [a, b] = (ref ?? "").split(":");
                    const p1 = a ? parseA1Ref(a) : null;
                    const p2 = b ? parseA1Ref(b) : null;
                    if (p1 && p2)
                        merges.push({ r1: p1.row, c1: p1.col, r2: p2.row, c2: p2.col });
                }
                if (merges.length)
                    sheet.merges = merges;
            }
            if (sheetData)
                readSheetData(sheet, sheetData, shared, styles);
        }
        wb.sheets.push(sheet);
    }
    return wb;
}
function readSheetData(sheet, sheetData, shared, styles) {
    for (const rowEl of Array.from(sheetData.getElementsByTagName("row"))) {
        const rAttr = rowEl.getAttribute("r");
        let rowNum = rAttr ? Number(rAttr) : 0;
        let colCursor = 0;
        for (const c of Array.from(rowEl.children)) {
            if (c.localName !== "c")
                continue;
            const ref = c.getAttribute("r");
            let row = rowNum;
            let col;
            if (ref) {
                const p = parseA1Ref(ref);
                if (!p)
                    continue;
                row = p.row;
                col = p.col;
                colCursor = col;
            }
            else {
                col = ++colCursor;
            }
            if (!row)
                continue;
            const t = c.getAttribute("t");
            const fEl = firstByLocal(c, "f");
            const vEl = firstByLocal(c, "v");
            const isEl = firstByLocal(c, "is");
            const formulaText = fEl?.textContent?.trim();
            const formula = formulaText ? formulaText : undefined;
            let value = "";
            let kind = "blank";
            if (t === "s") {
                value = shared[Number(vEl?.textContent ?? "0")] ?? "";
                kind = "s";
            }
            else if (t === "inlineStr") {
                value = isEl ? Array.from(isEl.getElementsByTagName("t")).map((x) => x.textContent ?? "").join("") : "";
                kind = "s";
            }
            else if (t === "str") {
                value = vEl?.textContent ?? "";
                kind = "s";
            }
            else if (t === "b") {
                value = vEl?.textContent === "1" ? "TRUE" : "FALSE";
                kind = "b";
            }
            else if (t === "e") {
                value = vEl?.textContent ?? "";
                kind = "e";
            }
            else {
                value = vEl?.textContent ?? "";
                kind = value === "" ? "blank" : "n";
            }
            const cell = {
                row,
                col,
                value,
                kind,
                formula,
                el: c,
                style: c.getAttribute("s") ?? undefined,
            };
            if (kind === "n") {
                const fmt = resolveXlsxFmt(styles, cell.style);
                if (fmt != null) {
                    cell.numFmt = fmt;
                    const d = formatNumber(fmt, value);
                    if (d != null)
                        cell.display = d;
                }
            }
            if (cell.style != null)
                cell.cellStyle = styles.xfStyles[Number(cell.style)];
            sheet.cells.set(key(row, col), cell);
            noteExtent(sheet, row, col);
        }
    }
}
function ensureXlsxCellEl(sheet, cell) {
    if (cell.el)
        return cell.el;
    const doc = sheet.doc;
    const sheetData = sheet.sheetData;
    const ns = doc.documentElement.namespaceURI || SS_MAIN;
    // find or create the <row>
    let rowEl;
    let insertRowBefore = null;
    for (const r of Array.from(sheetData.children)) {
        if (r.localName !== "row")
            continue;
        const rn = Number(r.getAttribute("r") || "0");
        if (rn === cell.row) {
            rowEl = r;
            break;
        }
        if (rn > cell.row) {
            insertRowBefore = r;
            break;
        }
    }
    if (!rowEl) {
        rowEl = doc.createElementNS(ns, "row");
        rowEl.setAttribute("r", String(cell.row));
        sheetData.insertBefore(rowEl, insertRowBefore);
    }
    // find or create the <c> in column order
    const ref = colToLetters(cell.col) + cell.row;
    let insertCellBefore = null;
    for (const c of Array.from(rowEl.children)) {
        if (c.localName !== "c")
            continue;
        const cref = c.getAttribute("r");
        const p = cref ? parseA1Ref(cref) : null;
        if (p && p.col === cell.col)
            return (cell.el = c);
        if (p && p.col > cell.col) {
            insertCellBefore = c;
            break;
        }
    }
    const cEl = doc.createElementNS(ns, "c");
    cEl.setAttribute("r", ref);
    if (cell.style)
        cEl.setAttribute("s", cell.style);
    rowEl.insertBefore(cEl, insertCellBefore);
    cell.el = cEl;
    return cEl;
}
function writeXlsxCell(sheet, cell) {
    const doc = sheet.doc;
    const ns = doc.documentElement.namespaceURI || SS_MAIN;
    const c = ensureXlsxCellEl(sheet, cell);
    removeByLocal(c, "f");
    removeByLocal(c, "v");
    removeByLocal(c, "is");
    const addV = (text) => {
        const v = doc.createElementNS(ns, "v");
        v.textContent = text;
        c.appendChild(v);
    };
    if (cell.formula != null) {
        const f = doc.createElementNS(ns, "f");
        f.textContent = cell.formula;
        c.appendChild(f);
        if (cell.kind === "n") {
            c.removeAttribute("t");
            if (cell.value !== "")
                addV(cell.value);
        }
        else if (cell.kind === "b") {
            c.setAttribute("t", "b");
            addV(cell.value === "TRUE" || cell.value === "1" ? "1" : "0");
        }
        else if (cell.kind === "e") {
            c.setAttribute("t", "e");
            addV(cell.value);
        }
        else if (cell.kind === "blank" || cell.value === "") {
            c.removeAttribute("t");
        }
        else {
            c.setAttribute("t", "str");
            addV(cell.value);
        }
        return;
    }
    // literal
    if (cell.value === "" || cell.kind === "blank") {
        c.removeAttribute("t");
    }
    else if (cell.kind === "n") {
        c.removeAttribute("t");
        addV(cell.value);
    }
    else if (cell.kind === "b") {
        c.setAttribute("t", "b");
        addV(cell.value === "TRUE" || cell.value === "1" ? "1" : "0");
    }
    else if (cell.kind === "e") {
        c.setAttribute("t", "e");
        addV(cell.value);
    }
    else {
        c.setAttribute("t", "inlineStr");
        const is = doc.createElementNS(ns, "is");
        const t = doc.createElementNS(ns, "t");
        t.setAttribute("xml:space", "preserve");
        t.textContent = cell.value;
        is.appendChild(t);
        c.appendChild(is);
    }
}
// Set a single column's width (px) in the worksheet's <cols>, creating <cols>/<col>
// as needed and splitting any existing run that covers the column. Keeps colWidths in sync.
export function setXlsxColWidth(sheet, col, px) {
    if (!sheet.colWidths)
        sheet.colWidths = new Map();
    sheet.colWidths.set(col, px);
    const doc = sheet.doc;
    if (!doc)
        return;
    const ns = doc.documentElement.namespaceURI || SS_MAIN;
    const width = Math.max(0, (px - 5) / 7);
    let colsEl = doc.getElementsByTagName("cols")[0];
    if (!colsEl) {
        colsEl = doc.createElementNS(ns, "cols");
        // <cols> must precede <sheetData> per the schema.
        sheet.sheetData?.parentNode?.insertBefore(colsEl, sheet.sheetData);
    }
    // Narrow any run that spans `col` so we can give `col` its own entry.
    for (const c of Array.from(colsEl.children)) {
        if (c.localName !== "col")
            continue;
        const min = Number(c.getAttribute("min") || "0");
        const max = Number(c.getAttribute("max") || String(min));
        if (col < min || col > max)
            continue;
        if (min === max) {
            c.setAttribute("width", String(width));
            c.setAttribute("customWidth", "1");
            sheet.layoutDirty = true;
            return;
        }
        // Split: left part [min..col-1], right part [col+1..max], plus the single col.
        if (col > min) {
            const left = c.cloneNode(true);
            left.setAttribute("min", String(min));
            left.setAttribute("max", String(col - 1));
            colsEl.insertBefore(left, c);
        }
        if (col < max) {
            const right = c.cloneNode(true);
            right.setAttribute("min", String(col + 1));
            right.setAttribute("max", String(max));
            colsEl.insertBefore(right, c);
        }
        c.setAttribute("min", String(col));
        c.setAttribute("max", String(col));
        c.setAttribute("width", String(width));
        c.setAttribute("customWidth", "1");
        sheet.layoutDirty = true;
        return;
    }
    const colEl = doc.createElementNS(ns, "col");
    colEl.setAttribute("min", String(col));
    colEl.setAttribute("max", String(col));
    colEl.setAttribute("width", String(width));
    colEl.setAttribute("customWidth", "1");
    colsEl.appendChild(colEl);
    sheet.layoutDirty = true;
}
// Set a single row's height (px) on its <row>, creating the row element if absent.
export function setXlsxRowHeight(sheet, row, px) {
    if (!sheet.rowHeights)
        sheet.rowHeights = new Map();
    sheet.rowHeights.set(row, px);
    const doc = sheet.doc;
    const sd = sheet.sheetData;
    if (!doc || !sd)
        return;
    const ns = doc.documentElement.namespaceURI || SS_MAIN;
    const pt = (px * 3) / 4;
    let rowEl;
    for (const re of Array.from(sd.children)) {
        if (re.localName === "row" && Number(re.getAttribute("r") || "0") === row) {
            rowEl = re;
            break;
        }
    }
    if (!rowEl) {
        rowEl = doc.createElementNS(ns, "row");
        rowEl.setAttribute("r", String(row));
        // Insert keeping rows in ascending order.
        let next = null;
        for (const re of Array.from(sd.children)) {
            if (re.localName === "row" && Number(re.getAttribute("r") || "0") > row) {
                next = re;
                break;
            }
        }
        sd.insertBefore(rowEl, next);
    }
    rowEl.setAttribute("ht", String(pt));
    rowEl.setAttribute("customHeight", "1");
    sheet.layoutDirty = true;
}
// Add or remove a merged range (1-based, inclusive). The top-left cell shows through;
// any cells the merge hides keep their data (so unmerging restores it). Updates the
// worksheet's <mergeCells> element and the in-memory merge list.
export function setXlsxMerge(sheet, r1, c1, r2, c2, merge) {
    const top = Math.min(r1, r2), left = Math.min(c1, c2), bottom = Math.max(r1, r2), right = Math.max(c1, c2);
    const ref = `${colToLetters(left)}${top}:${colToLetters(right)}${bottom}`;
    const merges = (sheet.merges ??= []);
    const idx = merges.findIndex((m) => m.r1 === top && m.c1 === left && m.r2 === bottom && m.c2 === right);
    if (merge) {
        if (idx === -1)
            merges.push({ r1: top, c1: left, r2: bottom, c2: right });
    }
    else if (idx !== -1) {
        merges.splice(idx, 1);
    }
    const doc = sheet.doc;
    if (!doc)
        return;
    const ns = doc.documentElement.namespaceURI || SS_MAIN;
    let mcEl = doc.getElementsByTagName("mergeCells")[0];
    if (merge) {
        if (!mcEl) {
            mcEl = doc.createElementNS(ns, "mergeCells");
            // <mergeCells> follows <sheetData> in the schema.
            sheet.sheetData?.parentNode?.insertBefore(mcEl, sheet.sheetData.nextSibling);
        }
        const exists = Array.from(mcEl.children).some((m) => m.getAttribute("ref") === ref);
        if (!exists) {
            const mc = doc.createElementNS(ns, "mergeCell");
            mc.setAttribute("ref", ref);
            mcEl.appendChild(mc);
        }
    }
    else if (mcEl) {
        for (const m of Array.from(mcEl.children))
            if (m.getAttribute("ref") === ref)
                mcEl.removeChild(m);
    }
    if (mcEl) {
        if (mcEl.children.length === 0)
            mcEl.parentNode?.removeChild(mcEl);
        else
            mcEl.setAttribute("count", String(mcEl.children.length));
    }
    sheet.layoutDirty = true;
}
const xmlOf = (el) => new XMLSerializer().serializeToString(el);
// Find a matching child in a style pool (deduped by serialized form) or append it;
// returns its index and keeps the pool's count attribute in sync.
function poolIndex(parent, candidate) {
    const want = xmlOf(candidate);
    const kids = Array.from(parent.children);
    for (let i = 0; i < kids.length; i++)
        if (xmlOf(kids[i]) === want)
            return i;
    parent.appendChild(candidate);
    parent.setAttribute("count", String(parent.children.length));
    return parent.children.length - 1;
}
const argbOf = (css) => "FF" + css.replace("#", "").toUpperCase();
/**
 * Apply a style change to a cell, managing the xlsx style pools: derive a new font /
 * fill / border from the cell's current format plus the change, find-or-create each in
 * styles.xml, find-or-create the combined <xf>, and point the cell at it.
 */
export function setXlsxCellStyle(wb, sheet, cell, change) {
    const doc = wb.stylesDoc;
    if (!doc)
        return;
    const ns = doc.documentElement.namespaceURI || SS_MAIN;
    const ce = (name) => doc.createElementNS(ns, name);
    const root = doc.documentElement;
    const pool = (name) => firstByLocal(root, name) ?? root.appendChild(ce(name));
    const fontsEl = pool("fonts");
    const fillsEl = pool("fills");
    const bordersEl = pool("borders");
    const cellXfsEl = pool("cellXfs");
    const curXf = cellXfsEl.children[cell.style ? Number(cell.style) : 0];
    const numFmtId = curXf?.getAttribute("numFmtId") || "0";
    const curFontId = Number(curXf?.getAttribute("fontId") || "0");
    const curFillId = Number(curXf?.getAttribute("fillId") || "0");
    const curBorderId = Number(curXf?.getAttribute("borderId") || "0");
    const cur = cell.cellStyle ?? {};
    const bold = change.bold ?? cur.bold;
    const italic = change.italic ?? cur.italic;
    const color = change.color ?? cur.color;
    const bg = change.bg ?? cur.bg;
    const align = change.align ?? cur.align;
    // Border sides: start from the current borders, apply the all-sides toggle and/or per-side change.
    const curSides = {
        top: !!cur.borders?.top,
        right: !!cur.borders?.right,
        bottom: !!cur.borders?.bottom,
        left: !!cur.borders?.left,
    };
    let sides = curSides;
    let borderChanged = false;
    if (change.border !== undefined) {
        sides = { top: change.border, right: change.border, bottom: change.border, left: change.border };
        borderChanged = true;
    }
    if (change.borderSides) {
        sides = { ...sides, ...change.borderSides };
        borderChanged = true;
    }
    // Font: clone the current one and toggle bold/italic/colour.
    const baseFont = fontsEl.children[curFontId];
    const font = baseFont ? baseFont.cloneNode(true) : ce("font");
    const flag = (tag, on) => {
        const ex = firstByLocal(font, tag);
        if (on && !ex)
            font.appendChild(ce(tag));
        else if (!on && ex)
            font.removeChild(ex);
    };
    flag("b", bold);
    flag("i", italic);
    if (color) {
        const col = firstByLocal(font, "color") ?? font.appendChild(ce("color"));
        col.removeAttribute("theme");
        col.removeAttribute("tint");
        col.setAttribute("rgb", argbOf(color));
    }
    const fontId = poolIndex(fontsEl, font);
    // Fill (solid) when set; else keep the current fill.
    let fillId = curFillId;
    if (bg) {
        const fill = ce("fill");
        const pat = ce("patternFill");
        pat.setAttribute("patternType", "solid");
        const fg = ce("fgColor");
        fg.setAttribute("rgb", argbOf(bg));
        pat.appendChild(fg);
        fill.appendChild(pat);
        fillId = poolIndex(fillsEl, fill);
    }
    // Border: rebuild the per-side border element only when the change touches borders.
    let borderId = curBorderId;
    if (borderChanged) {
        const bd = ce("border");
        for (const side of ["left", "right", "top", "bottom"]) {
            const s = ce(side);
            if (sides[side]) {
                s.setAttribute("style", "thin");
                const cc = ce("color");
                cc.setAttribute("rgb", "FF000000");
                s.appendChild(cc);
            }
            bd.appendChild(s);
        }
        borderId = poolIndex(bordersEl, bd);
    }
    const xf = ce("xf");
    xf.setAttribute("numFmtId", numFmtId);
    xf.setAttribute("fontId", String(fontId));
    xf.setAttribute("fillId", String(fillId));
    xf.setAttribute("borderId", String(borderId));
    xf.setAttribute("xfId", "0");
    xf.setAttribute("applyFont", "1");
    if (bg)
        xf.setAttribute("applyFill", "1");
    if (borderId)
        xf.setAttribute("applyBorder", "1");
    if (align) {
        xf.setAttribute("applyAlignment", "1");
        const a = ce("alignment");
        a.setAttribute("horizontal", align);
        xf.appendChild(a);
    }
    const sIdx = poolIndex(cellXfsEl, xf);
    cell.style = String(sIdx);
    ensureXlsxCellEl(sheet, cell).setAttribute("s", String(sIdx));
    const anySide = sides.top || sides.right || sides.bottom || sides.left;
    cell.cellStyle = {
        bold,
        italic,
        color,
        bg,
        align,
        borders: anySide
            ? {
                top: sides.top ? "#000" : undefined,
                right: sides.right ? "#000" : undefined,
                bottom: sides.bottom ? "#000" : undefined,
                left: sides.left ? "#000" : undefined,
            }
            : undefined,
    };
    cell.edited = true;
    wb.stylesDirty = true;
}
function writeXlsx(wb) {
    for (const sheet of wb.sheets) {
        if (!sheet.doc || !sheet.sheetData)
            continue;
        let touched = false;
        for (const cell of sheet.cells.values()) {
            if (cell.edited || cell.recomputed) {
                writeXlsxCell(sheet, cell);
                touched = true;
            }
        }
        if ((touched || sheet.layoutDirty) && sheet.path)
            wb.files[sheet.path] = serializeXml(sheet.doc);
    }
    if (wb.stylesDirty && wb.stylesDoc)
        wb.files["xl/styles.xml"] = serializeXml(wb.stylesDoc);
}
// ---------------------------------------------------------------------------
// ods (ODF spreadsheet)
// ---------------------------------------------------------------------------
const ODS = {
    office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
};
const REPEAT_CAP = 1024;
/** Replace text outside single-quoted string literals. */
function replaceOutsideStrings(s, fn) {
    let out = "";
    let i = 0;
    while (i < s.length) {
        const q = s.indexOf('"', i);
        const qq = s.indexOf("'", i);
        let next = -1;
        let quote = '"';
        if (q === -1 && qq === -1)
            next = -1;
        else if (q === -1) {
            next = qq;
            quote = "'";
        }
        else if (qq === -1) {
            next = q;
            quote = '"';
        }
        else if (q < qq) {
            next = q;
            quote = '"';
        }
        else {
            next = qq;
            quote = "'";
        }
        if (next === -1) {
            out += fn(s.slice(i));
            break;
        }
        out += fn(s.slice(i, next));
        const end = s.indexOf(quote, next + 1);
        if (end === -1) {
            out += s.slice(next);
            break;
        }
        out += s.slice(next, end + 1);
        i = end + 1;
    }
    return out;
}
/** ODF formula (`of:=[.A1]+[.B1]`) -> A1 (`A1+B1`). */
export function odfToA1(odf) {
    let core = odf.replace(/^of:=/, "").replace(/^=/, "");
    core = core.replace(/\[([^\]]*)\]/g, (_, inner) => {
        if (!inner || inner.includes("#"))
            return inner.replace(/\./g, ""); // #REF! etc.
        const parts = inner.split(":");
        const mapped = parts.map((part) => {
            const dot = part.lastIndexOf(".");
            const sheet = dot >= 0 ? part.slice(0, dot) : "";
            const ref = dot >= 0 ? part.slice(dot + 1) : part;
            return { sheet, ref };
        });
        const sheet = mapped[0].sheet;
        const cells = mapped.map((m) => m.ref).join(":");
        return (sheet ? sheet + "!" : "") + cells;
    });
    return replaceOutsideStrings(core, (chunk) => chunk.replace(/;/g, ","));
}
/** A1 (`A1+B1`) -> ODF formula (`of:=[.A1]+[.B1]`). Used only for user-typed formulas. */
export function a1ToOdf(a1) {
    const refRe = /(?<![A-Za-z0-9_.$])(?:('[^']*'|[A-Za-z_][A-Za-z0-9_]*)!)?(\$?[A-Z]{1,3}\$?[0-9]+)(?::(\$?[A-Z]{1,3}\$?[0-9]+))?(?![A-Za-z0-9_(])/g;
    const converted = replaceOutsideStrings(a1, (chunk) => {
        const semi = chunk.replace(/,/g, ";");
        return semi.replace(refRe, (_m, sheet, c1, c2) => {
            const sh = sheet ?? "";
            const range = c2 ? `${c1}:.${c2}` : c1;
            return `[${sh}.${range}]`;
        });
    });
    return "of:=" + converted;
}
function odsCellText(cell) {
    return Array.from(cell.getElementsByTagName("text:p"))
        .map((p) => p.textContent ?? "")
        .join("\n");
}
function readOds(files) {
    const contentFile = files["content.xml"];
    if (!contentFile)
        throw new Error("not an .ods: content.xml missing");
    const contentDoc = parseXml(contentFile);
    const wb = { kind: "ods", sheets: [], files, contentDoc, contentPath: "content.xml" };
    for (const table of Array.from(contentDoc.getElementsByTagName("table:table"))) {
        const name = table.getAttribute("table:name") ?? `Sheet${wb.sheets.length + 1}`;
        const sheet = { name, cells: new Map(), maxRow: 0, maxCol: 0, tableEl: table };
        readOdsTable(sheet, table);
        wb.sheets.push(sheet);
    }
    return wb;
}
function readOdsTable(sheet, table) {
    let rowNum = 0;
    const rows = [];
    const collect = (parent) => {
        for (const ch of Array.from(parent.children)) {
            if (ch.localName === "table-row")
                rows.push(ch);
            else if (ch.localName === "table-header-rows" || ch.localName === "table-rows")
                collect(ch);
        }
    };
    collect(table);
    for (const rowEl of rows) {
        const rrep = Math.max(1, Number(rowEl.getAttribute("table:number-rows-repeated") || "1"));
        const parsedCells = parseOdsRow(rowEl);
        const rowHasContent = parsedCells.some((c) => c.has);
        const copies = rowHasContent ? Math.min(rrep, REPEAT_CAP) : 0;
        for (let k = 0; k < copies; k++) {
            const r = rowNum + 1 + k;
            for (const pc of parsedCells) {
                if (!pc.has)
                    continue;
                for (let j = 0; j < pc.span; j++) {
                    const c = pc.cell;
                    sheet.cells.set(key(r, c.col + j), { ...c, row: r, col: c.col + j });
                    noteExtent(sheet, r, c.col + j);
                }
            }
        }
        rowNum += rrep;
    }
}
function parseOdsRow(rowEl) {
    const out = [];
    let col = 0;
    for (const cellEl of Array.from(rowEl.children)) {
        const local = cellEl.localName;
        if (local !== "table-cell" && local !== "covered-table-cell")
            continue;
        const crep = Math.max(1, Number(cellEl.getAttribute("table:number-columns-repeated") || "1"));
        const startCol = col + 1;
        col += crep;
        if (local === "covered-table-cell")
            continue; // merged-away cell
        const valueType = cellEl.getAttribute("office:value-type");
        const formulaRaw = cellEl.getAttribute("table:formula") ?? undefined;
        const style = cellEl.getAttribute("table:style-name") ?? undefined;
        const text = odsCellText(cellEl);
        let value = "";
        let kind = "blank";
        let display;
        if (valueType === "float" || valueType === "percentage" || valueType === "currency") {
            value = cellEl.getAttribute("office:value") ?? text;
            // ODF stores the producer's formatted text in <text:p>; use it as the display.
            if (text !== "" && text !== value)
                display = text;
            kind = "n";
        }
        else if (valueType === "boolean") {
            value = cellEl.getAttribute("office:boolean-value") === "true" ? "TRUE" : "FALSE";
            kind = "b";
        }
        else if (valueType === "string") {
            value = cellEl.getAttribute("office:string-value") ?? odsCellText(cellEl);
            kind = "s";
        }
        else if (valueType === "date") {
            value = cellEl.getAttribute("office:date-value") ?? text;
            if (text !== "" && text !== value)
                display = text;
            kind = "s";
        }
        else if (valueType === "time") {
            value = cellEl.getAttribute("office:time-value") ?? text;
            if (text !== "" && text !== value)
                display = text;
            kind = "s";
        }
        else {
            value = odsCellText(cellEl);
            kind = value === "" ? "blank" : "s";
        }
        const has = value !== "" || formulaRaw != null || style != null;
        if (!has) {
            out.push({ has: false, span: crep });
            continue;
        }
        const cell = {
            row: 0,
            col: startCol,
            value,
            kind,
            display,
            formula: formulaRaw ? odfToA1(formulaRaw) : undefined,
            odfFormula: formulaRaw,
            style,
            el: cellEl,
        };
        out.push({ has: true, span: Math.min(crep, REPEAT_CAP), cell });
    }
    return out;
}
function makeOdsCell(doc, cell, edited) {
    // Untouched cell: clone the original verbatim (preserves dates, formats, rich text).
    if (cell.el && !cell.edited && !cell.recomputed) {
        const clone = cell.el.cloneNode(true);
        clone.removeAttribute("table:number-columns-repeated");
        clone.removeAttribute("table:number-rows-repeated");
        return clone;
    }
    const c = doc.createElementNS(ODS.table, "table:table-cell");
    if (cell.style)
        c.setAttributeNS(ODS.table, "table:style-name", cell.style);
    const formulaToWrite = edited && cell.formula != null ? a1ToOdf(cell.formula) : cell.odfFormula;
    if (formulaToWrite)
        c.setAttributeNS(ODS.table, "table:formula", formulaToWrite);
    const addText = (text) => {
        if (text === "")
            return;
        const p = doc.createElementNS(ODS.text, "text:p");
        p.textContent = text;
        c.appendChild(p);
    };
    if (cell.kind === "n") {
        c.setAttributeNS(ODS.office, "office:value-type", "float");
        c.setAttributeNS(ODS.office, "office:value", cell.value);
        addText(cell.value);
    }
    else if (cell.kind === "b") {
        c.setAttributeNS(ODS.office, "office:value-type", "boolean");
        c.setAttributeNS(ODS.office, "office:boolean-value", cell.value === "TRUE" ? "true" : "false");
        addText(cell.value);
    }
    else if (cell.kind === "s" || cell.kind === "e") {
        c.setAttributeNS(ODS.office, "office:value-type", "string");
        c.setAttributeNS(ODS.office, "office:string-value", cell.value);
        addText(cell.value);
    }
    return c;
}
function writeOds(wb) {
    const doc = wb.contentDoc;
    for (const sheet of wb.sheets) {
        const table = sheet.tableEl;
        if (!table)
            continue;
        // preserve structural children (column definitions etc.), drop existing rows
        const keep = [];
        for (const ch of Array.from(table.children)) {
            if (ch.localName !== "table-row" && ch.localName !== "table-header-rows" && ch.localName !== "table-rows") {
                keep.push(ch);
            }
        }
        while (table.firstChild)
            table.removeChild(table.firstChild);
        for (const k of keep)
            table.appendChild(k);
        const maxRow = Math.max(1, sheet.maxRow);
        const maxCol = Math.max(1, sheet.maxCol);
        for (let r = 1; r <= maxRow; r++) {
            const rowEl = doc.createElementNS(ODS.table, "table:table-row");
            let lastContent = 0;
            for (let c = maxCol; c >= 1; c--) {
                if (getCell(sheet, r, c)) {
                    lastContent = c;
                    break;
                }
            }
            for (let c = 1; c <= lastContent; c++) {
                const cell = getCell(sheet, r, c);
                rowEl.appendChild(cell ? makeOdsCell(doc, cell, !!cell.edited) : doc.createElementNS(ODS.table, "table:table-cell"));
            }
            if (lastContent < maxCol) {
                const filler = doc.createElementNS(ODS.table, "table:table-cell");
                filler.setAttributeNS(ODS.table, "table:number-columns-repeated", String(maxCol - lastContent));
                rowEl.appendChild(filler);
            }
            table.appendChild(rowEl);
        }
    }
    wb.files["content.xml"] = serializeXml(doc);
}
function applyResult(cell, res) {
    let value;
    let kind;
    if (res == null) {
        value = "";
        kind = "blank";
    }
    else if (typeof res === "number") {
        value = Number.isFinite(res) ? numToStr(res) : "#NUM!";
        kind = Number.isFinite(res) ? "n" : "e";
    }
    else if (typeof res === "boolean") {
        value = res ? "TRUE" : "FALSE";
        kind = "b";
    }
    else if (Array.isArray(res)) {
        applyResult(cell, (res[0] && res[0][0]) ?? "");
        return;
    }
    else if (typeof res === "object") {
        value = String(res); // FormulaError -> "#DIV/0!" etc.
        kind = "e";
    }
    else {
        value = String(res);
        kind = "s";
    }
    if (value !== cell.value || kind !== cell.kind) {
        cell.value = value;
        cell.kind = kind;
        cell.recomputed = true;
    }
    // Refresh the formatted display from the (possibly new) value.
    cell.display = kind === "n" && cell.numFmt != null ? formatNumber(cell.numFmt, value) ?? undefined : undefined;
}
/** Recompute every formula cell's cached value, in dependency order. */
export function recalc(wb) {
    const FP = FormulaParser;
    const byName = new Map();
    for (const s of wb.sheets)
        byName.set(s.name, s);
    const defaultSheet = wb.sheets[0]?.name;
    const lookup = (sheetName, r, c) => {
        const sheet = (sheetName && byName.get(sheetName)) || (defaultSheet ? byName.get(defaultSheet) : undefined);
        return sheet ? getCell(sheet, r, c) : undefined;
    };
    const nodes = [];
    const index = new Map();
    const idOf = (sheetName, r, c) => `${sheetName} ${r}:${c}`;
    for (const sheet of wb.sheets) {
        for (const cell of sheet.cells.values()) {
            if (cell.formula == null)
                continue;
            const node = { sheet, cell, id: idOf(sheet.name, cell.row, cell.col), deps: new Set() };
            nodes.push(node);
            index.set(node.id, node);
        }
    }
    if (!nodes.length)
        return;
    const depParser = new FP.DepParser({ onVariable: () => null });
    for (const node of nodes) {
        let refs = [];
        try {
            refs = depParser.parse(node.cell.formula, { row: node.cell.row, col: node.cell.col, sheet: node.sheet.name });
        }
        catch {
            refs = [];
        }
        for (const ref of refs) {
            const sName = ref.sheet ?? node.sheet.name;
            if (ref.from) {
                const from = ref.from;
                const to = ref.to;
                for (const other of nodes) {
                    if (other.sheet.name !== sName)
                        continue;
                    const { row, col } = other.cell;
                    if (row >= from.row && row <= to.row && col >= from.col && col <= to.col)
                        node.deps.add(other.id);
                }
            }
            else {
                const depId = idOf(sName, ref.row, ref.col);
                if (index.has(depId))
                    node.deps.add(depId);
            }
        }
    }
    // Kahn topological sort: dependencies evaluated before dependents.
    const indeg = new Map();
    const dependents = new Map();
    for (const node of nodes)
        indeg.set(node.id, node.deps.size);
    for (const node of nodes)
        for (const d of node.deps) {
            if (!dependents.has(d))
                dependents.set(d, []);
            dependents.get(d).push(node.id);
        }
    const queue = [];
    for (const node of nodes)
        if ((indeg.get(node.id) ?? 0) === 0)
            queue.push(node.id);
    const order = [];
    while (queue.length) {
        const id = queue.shift();
        order.push(index.get(id));
        for (const dep of dependents.get(id) ?? []) {
            const d = (indeg.get(dep) ?? 1) - 1;
            indeg.set(dep, d);
            if (d === 0)
                queue.push(dep);
        }
    }
    if (order.length < nodes.length) {
        const seen = new Set(order.map((n) => n.id)); // cycles: best-effort single pass
        for (const node of nodes)
            if (!seen.has(node.id))
                order.push(node);
    }
    const parser = new FP({
        onCell: (ref) => typedValue(lookup(ref.sheet, ref.row, ref.col)),
        onRange: (ref) => {
            const out = [];
            for (let r = ref.from.row; r <= ref.to.row; r++) {
                const rowArr = [];
                for (let c = ref.from.col; c <= ref.to.col; c++)
                    rowArr.push(typedValue(lookup(ref.sheet, r, c)));
                out.push(rowArr);
            }
            return out;
        },
    });
    for (const node of order) {
        let res;
        try {
            res = parser.parse(node.cell.formula, { row: node.cell.row, col: node.cell.col, sheet: node.sheet.name });
        }
        catch {
            continue; // unsupported function / parse error: keep the file's cached value
        }
        // A fresh recompute can error on blank inputs (e.g. DATEDIF on an empty date) even
        // though the file holds a valid cached result; keep that result rather than show an error.
        const isErr = res != null && typeof res === "object" && !Array.isArray(res);
        if (isErr && node.cell.value !== "" && node.cell.kind !== "e")
            continue;
        applyResult(node.cell, res);
    }
}
// ---------------------------------------------------------------------------
// Public read / write
// ---------------------------------------------------------------------------
export function readWorkbook(bytes) {
    const files = unzipSync(bytes);
    if (files["xl/workbook.xml"])
        return readXlsx(files);
    if (files["content.xml"]) {
        const mt = files["mimetype"] ? strFromU8(files["mimetype"]) : "";
        if (!mt || mt.includes("spreadsheet"))
            return readOds(files);
    }
    throw new Error("unrecognized workbook: expected .xlsx or .ods");
}
export function writeWorkbook(wb) {
    recalc(wb);
    if (wb.kind === "xlsx") {
        writeXlsx(wb);
        return zipSync(wb.files);
    }
    writeOds(wb);
    // ODF requires the "mimetype" entry first and stored (uncompressed).
    const repacked = {};
    if (wb.files["mimetype"])
        repacked["mimetype"] = [wb.files["mimetype"], { level: 0 }];
    for (const [name, data] of Object.entries(wb.files)) {
        if (name === "mimetype")
            continue;
        repacked[name] = data;
    }
    return zipSync(repacked);
}
/** Commit a raw grid edit (the text a user typed) into the model. */
export function setCellInput(sheet, row, col, raw) {
    const existing = getCell(sheet, row, col);
    if (raw.startsWith("=")) {
        const cell = ensureCell(sheet, row, col);
        cell.formula = raw.slice(1).trim();
        cell.odfFormula = undefined;
        cell.edited = true;
        return;
    }
    if (existing == null && raw === "")
        return;
    const cell = ensureCell(sheet, row, col);
    cell.formula = undefined;
    cell.odfFormula = undefined;
    cell.edited = true;
    if (raw === "") {
        cell.value = "";
        cell.kind = "blank";
    }
    else if (isNumeric(raw)) {
        cell.value = raw.trim();
        cell.kind = "n";
    }
    else if (raw.toUpperCase() === "TRUE" || raw.toUpperCase() === "FALSE") {
        cell.value = raw.toUpperCase();
        cell.kind = "b";
    }
    else {
        cell.value = raw;
        cell.kind = "s";
    }
    // Re-apply the cell's number format (a typed value keeps the cell's format, like Excel).
    cell.display =
        cell.kind === "n" && cell.numFmt != null ? formatNumber(cell.numFmt, cell.value) ?? undefined : undefined;
}
const ROWS_MIN = 24;
const COLS_MIN = 12;
const ROWS_CAP = 5000;
const COLS_CAP = 256;
const ROW_CHUNK = 20; // rows added per "+ Row" click
const COL_CHUNK = 6; // columns added per "+ Col" click
const STYLE_ID = "sheetedit-style";
function injectStyles() {
    if (document.getElementById(STYLE_ID))
        return;
    const s = document.createElement("style");
    s.id = STYLE_ID;
    s.textContent = `
    .sheetedit-wrap { display:flex; flex-direction:column; height:100%; background:#1f2227; color:#e6e6e6; font:13px system-ui, sans-serif; }
    .sheetedit-toolbar { display:flex; flex-wrap:wrap; align-items:center; gap:5px; padding:5px 8px; background:#2b2f36; border-bottom:1px solid #1c1f24; }
    .sheetedit-btn {
      font:inherit; font-size:13px; background:#3a3f47; color:#e6e6e6; border:1px solid #4a4f57;
      border-radius:6px; padding:4px 9px; cursor:pointer; min-width:32px; line-height:1.1;
    }
    .sheetedit-btn:hover { background:#454b54; }
    .sheetedit-btn:focus-visible { outline:2px solid #6e7bff; outline-offset:1px; }
    .sheetedit-tb-sep { width:1px; align-self:stretch; background:#4a4f57; margin:1px 3px; }
    .sheetedit-color { width:30px; height:28px; padding:0; border:1px solid #4a4f57; border-radius:6px; background:#3a3f47; cursor:pointer; }
    .sheetedit-btn svg { display:block; width:16px; height:16px; }
    .sheetedit-table th.colhead, .sheetedit-table th.rownum, .sheetedit-table th.corner { cursor:pointer; }
    .sheetedit-table th.colhead:hover, .sheetedit-table th.rownum:hover, .sheetedit-table th.corner:hover { background:#e3e3e8; }
    .sheetedit-table td.sheetedit-sel input { background:rgba(110,123,255,0.18); }
    .sheetedit-pop { position:fixed; z-index:30; background:#2b2f36; border:1px solid #4a4f57; border-radius:8px; padding:4px; box-shadow:0 6px 18px rgba(0,0,0,0.45); display:flex; flex-direction:column; min-width:130px; }
    .sheetedit-pop-item { font:inherit; font-size:13px; text-align:left; background:transparent; color:#e6e6e6; border:0; border-radius:5px; padding:7px 11px; cursor:pointer; }
    .sheetedit-pop-item:hover { background:#3a3f47; }
    /* The grid is a light canvas (like a real spreadsheet) so the file's fills and
       font colours render faithfully and stay readable; the chrome stays dark. */
    .sheetedit-grid { flex:1; min-height:0; overflow:auto; background:#e9e9ec; }
    table.sheetedit-table { border-collapse:collapse; table-layout:fixed; font:13px/1.3 ui-sans-serif, system-ui, sans-serif; }
    .sheetedit-table th, .sheetedit-table td { padding:0; margin:0; }
    .sheetedit-table th { border:1px solid #d4d4d8; }
    /* Cell gridlines as box-shadows (not borders) so the file's own borders sit flush
       against their neighbours and touch, like a real spreadsheet. */
    .sheetedit-table td { background:#fff; box-shadow: inset -1px -1px 0 0 #e3e3e6; }
    .sheetedit-table th {
      position:sticky; top:0; z-index:2; background:#f1f1f4; color:#555; font-weight:600;
      padding:3px 8px; text-align:center; user-select:none;
    }
    .sheetedit-table th.corner { left:0; z-index:3; }
    .sheetedit-table th.rownum { position:sticky; left:0; z-index:1; top:auto; text-align:right; background:#f1f1f4; }
    /* Resize grips: a thin strip on the header border, wide enough to grab on touch. */
    .sheetedit-colgrip { position:absolute; top:0; right:-4px; width:9px; height:100%; cursor:col-resize; z-index:4; touch-action:none; }
    .sheetedit-rowgrip { position:absolute; left:0; bottom:-4px; width:100%; height:9px; cursor:row-resize; z-index:4; touch-action:none; }
    .sheetedit-colgrip:hover { box-shadow:inset -2px 0 0 0 #6e7bff; }
    .sheetedit-rowgrip:hover { box-shadow:inset 0 -2px 0 0 #6e7bff; }
    .sheetedit-table input {
      border:0; background:transparent; color:#1a1a1a; font:inherit; padding:3px 8px;
      width:100%; box-sizing:border-box; outline:none;
    }
    .sheetedit-table td.num input { text-align:right; font-variant-numeric:tabular-nums; }
    .sheetedit-table input:focus { box-shadow:inset 0 0 0 2px #6e7bff; background:#eef0ff; }
    .sheetedit-tabs { display:flex; align-items:center; gap:2px; padding:5px 8px; background:#2b2f36; border-top:1px solid #1c1f24; overflow-x:auto; }
    .sheetedit-tab {
      font:inherit; background:#3a3f47; color:#cfd3da; border:1px solid #4a4f57; border-bottom:none;
      border-radius:5px 5px 0 0; padding:4px 12px; cursor:pointer; white-space:nowrap;
    }
    .sheetedit-tab[aria-selected="true"] { background:#6e7bff; color:#fff; border-color:#6e7bff; }
    .sheetedit-tab:focus-visible { outline:2px solid #fff; outline-offset:1px; }
  `;
    document.head.appendChild(s);
}
export function createSheetEditor(container, bytes, options = {}) {
    const original = bytes.slice();
    let dirty = false;
    injectStyles();
    const wb = readWorkbook(bytes);
    // Trust the file's cached results on open (like Excel/LibreOffice); recalc only runs
    // after an edit. Recomputing on load would overwrite valid cached values whose inputs
    // are blank in this session (e.g. a DATEDIF age before a birthdate is entered).
    const wrap = document.createElement("div");
    wrap.className = "sheetedit-wrap";
    const toolbar = document.createElement("div");
    toolbar.className = "sheetedit-toolbar";
    const gridScroll = document.createElement("div");
    gridScroll.className = "sheetedit-grid";
    const tabs = document.createElement("div");
    tabs.className = "sheetedit-tabs";
    tabs.setAttribute("role", "tablist");
    tabs.setAttribute("aria-label", "Sheets");
    wrap.append(toolbar, gridScroll, tabs);
    container.appendChild(wrap);
    let active = 0;
    let inputs = new Map();
    let tds = new Map();
    // Extra rows/columns the user added beyond the sheet's used extent (per active sheet).
    let extraRows = 0;
    let extraCols = 0;
    // Selection rectangle (1-based, inclusive) and the anchor for shift-extend.
    let sel = null;
    let anchor = null;
    const paintSel = () => {
        for (const td of tds.values())
            td.classList.remove("sheetedit-sel");
        if (!sel)
            return;
        for (let r = sel.r1; r <= sel.r2; r++)
            for (let c = sel.c1; c <= sel.c2; c++)
                tds.get(key(r, c))?.classList.add("sheetedit-sel");
    };
    const setSel = (r1, c1, r2, c2) => {
        sel = { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
        paintSel();
    };
    const selectCell = (r, c, extend) => {
        if (extend && anchor)
            setSel(anchor.r, anchor.c, r, c);
        else {
            anchor = { r, c };
            setSel(r, c, r, c);
        }
    };
    // Rectangular range selection by dragging: mouse drag, or touch long-press then drag.
    // A plain tap still focuses a cell for editing; header/corner taps select a whole line.
    const cellAtPoint = (x, y) => {
        const el = document.elementFromPoint(x, y);
        const td = el?.closest("td");
        const rc = td?.dataset.rc;
        if (!rc)
            return null;
        const [r, c] = rc.split(":").map(Number);
        return { r, c };
    };
    let dragAnchor = null;
    let dragActive = false;
    let justDragged = false;
    let lpTimer = null;
    let lpStart = null;
    gridScroll.addEventListener("pointerdown", (e) => {
        if (resizing)
            return;
        const cell = cellAtPoint(e.clientX, e.clientY);
        if (!cell)
            return;
        if (e.pointerType === "touch") {
            lpStart = { x: e.clientX, y: e.clientY };
            lpTimer = window.setTimeout(() => {
                lpTimer = null;
                dragActive = true;
                dragAnchor = cell;
                anchor = cell;
                document.activeElement?.blur?.();
                setSel(cell.r, cell.c, cell.r, cell.c);
            }, 250);
        }
        else {
            dragAnchor = cell; // mouse: a click still edits; a drag selects
        }
    });
    gridScroll.addEventListener("pointermove", (e) => {
        if (e.pointerType === "touch") {
            if (lpTimer != null && lpStart) {
                if (Math.abs(e.clientX - lpStart.x) > 8 || Math.abs(e.clientY - lpStart.y) > 8) {
                    clearTimeout(lpTimer); // moved before the long-press fired: it is a scroll
                    lpTimer = null;
                }
                return;
            }
            if (!dragActive || !dragAnchor)
                return;
            const cell = cellAtPoint(e.clientX, e.clientY);
            if (cell) {
                e.preventDefault();
                setSel(dragAnchor.r, dragAnchor.c, cell.r, cell.c);
            }
        }
        else {
            if (!dragAnchor || e.buttons === 0)
                return;
            const cell = cellAtPoint(e.clientX, e.clientY);
            if (!cell || (cell.r === dragAnchor.r && cell.c === dragAnchor.c && !dragActive))
                return;
            if (!dragActive) {
                dragActive = true;
                anchor = dragAnchor;
                document.activeElement?.blur?.();
            }
            e.preventDefault();
            setSel(dragAnchor.r, dragAnchor.c, cell.r, cell.c);
        }
    }, { passive: false });
    const endDrag = () => {
        if (lpTimer != null) {
            clearTimeout(lpTimer);
            lpTimer = null;
        }
        dragAnchor = null;
        if (dragActive) {
            justDragged = true; // swallow the trailing tap so it does not enter edit mode
            setTimeout(() => (justDragged = false), 350);
        }
        dragActive = false;
    };
    window.addEventListener("pointerup", endDrag);
    window.addEventListener("pointercancel", endDrag);
    // Column-width / row-height drag from the header borders. `resizing` suppresses the
    // cell drag-select while a resize is in progress.
    let resizing = false;
    const startColResize = (e, col, colEl, startW) => {
        e.preventDefault();
        e.stopPropagation();
        resizing = true;
        const x0 = e.clientX;
        const onMove = (ev) => {
            const w = Math.max(24, Math.round(startW + (ev.clientX - x0)));
            colEl.style.width = `${w}px`;
        };
        const onUp = (ev) => {
            window.removeEventListener("pointermove", onMove);
            window.removeEventListener("pointerup", onUp);
            const w = Math.max(24, Math.round(startW + (ev.clientX - x0)));
            const sheet = wb.sheets[active];
            if (sheet && wb.kind === "xlsx") {
                setXlsxColWidth(sheet, col, w);
                mark();
            }
            else if (sheet) {
                (sheet.colWidths ??= new Map()).set(col, w);
            }
            renderGrid();
            setTimeout(() => (resizing = false), 0);
        };
        window.addEventListener("pointermove", onMove, { passive: false });
        window.addEventListener("pointerup", onUp);
    };
    const startRowResize = (e, row, rowEl, startH) => {
        e.preventDefault();
        e.stopPropagation();
        resizing = true;
        const y0 = e.clientY;
        const onMove = (ev) => {
            const h = Math.max(16, Math.round(startH + (ev.clientY - y0)));
            rowEl.style.height = `${h}px`;
        };
        const onUp = (ev) => {
            window.removeEventListener("pointermove", onMove);
            window.removeEventListener("pointerup", onUp);
            const h = Math.max(16, Math.round(startH + (ev.clientY - y0)));
            const sheet = wb.sheets[active];
            if (sheet && wb.kind === "xlsx") {
                setXlsxRowHeight(sheet, row, h);
                mark();
            }
            else if (sheet) {
                (sheet.rowHeights ??= new Map()).set(row, h);
            }
            renderGrid();
            setTimeout(() => (resizing = false), 0);
        };
        window.addEventListener("pointermove", onMove, { passive: false });
        window.addEventListener("pointerup", onUp);
    };
    const tbBtn = (label, title, onClick) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-btn";
        b.textContent = label;
        b.title = title;
        b.setAttribute("aria-label", title);
        b.addEventListener("click", onClick);
        return b;
    };
    const tbIcon = (svg, title, onClick) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "sheetedit-btn";
        b.innerHTML = svg;
        b.title = title;
        b.setAttribute("aria-label", title);
        b.addEventListener("click", onClick);
        return b;
    };
    // Style of the selection's top-left cell (used to toggle bold/italic/borders).
    const curStyle = () => (sel ? getCell(wb.sheets[active], sel.r1, sel.c1)?.cellStyle : undefined);
    // Apply a style change to every cell in the selection (xlsx only), then re-render.
    const applyStyle = (change) => {
        if (wb.kind !== "xlsx" || !sel)
            return;
        const sheet = wb.sheets[active];
        if (!sheet)
            return;
        let n = 0;
        for (let r = sel.r1; r <= sel.r2 && n < 4000; r++)
            for (let c = sel.c1; c <= sel.c2 && n < 4000; c++, n++)
                setXlsxCellStyle(wb, sheet, ensureCell(sheet, r, c), change);
        mark();
        renderGrid();
    };
    const applyBorder = (mode) => {
        if (wb.kind !== "xlsx" || !sel)
            return;
        const sheet = wb.sheets[active];
        if (!sheet)
            return;
        const { r1, c1, r2, c2 } = sel;
        let n = 0;
        for (let r = r1; r <= r2 && n < 4000; r++)
            for (let c = c1; c <= c2 && n < 4000; c++, n++) {
                let sides = {};
                if (mode === "all")
                    sides = { top: true, right: true, bottom: true, left: true };
                else if (mode === "none")
                    sides = { top: false, right: false, bottom: false, left: false };
                else {
                    if ((mode === "outer" || mode === "top") && r === r1)
                        sides.top = true;
                    if ((mode === "outer" || mode === "bottom") && r === r2)
                        sides.bottom = true;
                    if ((mode === "outer" || mode === "left") && c === c1)
                        sides.left = true;
                    if ((mode === "outer" || mode === "right") && c === c2)
                        sides.right = true;
                }
                if (Object.keys(sides).length)
                    setXlsxCellStyle(wb, sheet, ensureCell(sheet, r, c), { borderSides: sides });
            }
        mark();
        renderGrid();
    };
    let borderPop = null;
    const openBorderPopover = (btn) => {
        if (borderPop) {
            borderPop.remove();
            borderPop = null;
            return;
        }
        const pop = document.createElement("div");
        pop.className = "sheetedit-pop";
        const close = () => {
            pop.remove();
            borderPop = null;
            document.removeEventListener("pointerdown", onOutside, true);
        };
        const onOutside = (e) => {
            const t = e.target;
            if (!pop.contains(t) && !btn.contains(t))
                close();
        };
        const opts = [
            ["All borders", "all"],
            ["Outer border", "outer"],
            ["Top", "top"],
            ["Bottom", "bottom"],
            ["Left", "left"],
            ["Right", "right"],
            ["No border", "none"],
        ];
        for (const [label, mode] of opts) {
            const b = document.createElement("button");
            b.type = "button";
            b.className = "sheetedit-pop-item";
            b.textContent = label;
            b.addEventListener("click", () => {
                applyBorder(mode);
                close();
            });
            pop.appendChild(b);
        }
        document.body.appendChild(pop);
        borderPop = pop;
        const r = btn.getBoundingClientRect();
        pop.style.left = `${Math.round(r.left)}px`;
        pop.style.top = `${Math.round(r.bottom + 4)}px`;
        setTimeout(() => document.addEventListener("pointerdown", onOutside, true), 0);
    };
    // Merge the selection into one cell, or unmerge when the selection sits on an
    // existing merge. Merging a region first clears any merges it overlaps.
    const toggleMerge = () => {
        if (wb.kind !== "xlsx" || !sel)
            return;
        const sheet = wb.sheets[active];
        if (!sheet)
            return;
        const { r1, c1, r2, c2 } = sel;
        const merges = sheet.merges ?? [];
        const within = (m) => r1 >= m.r1 && c1 >= m.c1 && r2 <= m.r2 && c2 <= m.c2;
        const intersects = (m) => !(r2 < m.r1 || r1 > m.r2 || c2 < m.c1 || c1 > m.c2);
        const containing = merges.find(within); // selection inside (or equal to) a merge
        if (containing) {
            setXlsxMerge(sheet, containing.r1, containing.c1, containing.r2, containing.c2, false);
        }
        else if (r1 !== r2 || c1 !== c2) {
            for (const m of merges.filter(intersects))
                setXlsxMerge(sheet, m.r1, m.c1, m.r2, m.c2, false);
            setXlsxMerge(sheet, r1, c1, r2, c2, true);
        }
        else {
            return; // a single, unmerged cell: nothing to do
        }
        mark();
        renderGrid();
    };
    const ICON = {
        left: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M2 8h7M2 12h10"/></svg>`,
        center: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M4.5 8h7M3 12h10"/></svg>`,
        right: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M2 4h12M7 8h7M4 12h10"/></svg>`,
        borders: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4"><rect x="2" y="2" width="12" height="12"/><path d="M8 2v12M2 8h12"/></svg>`,
        merge: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"><rect x="1.5" y="4" width="13" height="8"/><path d="M5 6 7.5 8 5 10M11 6 8.5 8 11 10"/></svg>`,
    };
    const buildToolbar = () => {
        toolbar.innerHTML = "";
        const sep = () => {
            const d = document.createElement("div");
            d.className = "sheetedit-tb-sep";
            return d;
        };
        const colorInput = (title, def, apply) => {
            const i = document.createElement("input");
            i.type = "color";
            i.title = title;
            i.setAttribute("aria-label", title);
            i.className = "sheetedit-color";
            i.value = def;
            i.addEventListener("change", () => apply(i.value));
            return i;
        };
        toolbar.append(tbBtn("+ Row", "Add rows", () => {
            extraRows += ROW_CHUNK;
            renderGrid();
        }), tbBtn("+ Col", "Add columns", () => {
            extraCols += COL_CHUNK;
            renderGrid();
        }));
        if (wb.kind !== "xlsx")
            return; // setting styles is xlsx-only for now
        const bold = tbBtn("B", "Bold", () => applyStyle({ bold: !curStyle()?.bold }));
        bold.style.fontWeight = "700";
        const italic = tbBtn("I", "Italic", () => applyStyle({ italic: !curStyle()?.italic }));
        italic.style.fontStyle = "italic";
        toolbar.append(sep(), bold, italic, colorInput("Text colour", "#000000", (v) => applyStyle({ color: v })), colorInput("Fill colour", "#ffff00", (v) => applyStyle({ bg: v })), sep(), tbIcon(ICON.left, "Align left", () => applyStyle({ align: "left" })), tbIcon(ICON.center, "Align centre", () => applyStyle({ align: "center" })), tbIcon(ICON.right, "Align right", () => applyStyle({ align: "right" })), sep());
        const borderBtn = tbIcon(ICON.borders, "Borders", () => openBorderPopover(borderBtn));
        toolbar.append(borderBtn, tbIcon(ICON.merge, "Merge or unmerge cells", toggleMerge));
    };
    const mark = () => {
        if (!dirty) {
            dirty = true;
        }
        options.onChange?.();
    };
    const displayValue = (sheet, r, c) => cellDisplay(getCell(sheet, r, c));
    const refreshDisplays = (sheet, except) => {
        for (const [k, input] of inputs) {
            if (input === except)
                continue;
            const [r, c] = k.split(":").map(Number);
            input.value = displayValue(sheet, r, c);
            const cell = getCell(sheet, r, c);
            input.parentElement?.classList.toggle("num", cell?.kind === "n");
        }
    };
    const renderGrid = () => {
        const sheet = wb.sheets[active];
        if (!sheet)
            return;
        inputs = new Map();
        tds = new Map();
        gridScroll.innerHTML = "";
        const rows = Math.min(ROWS_CAP, Math.max(ROWS_MIN, sheet.maxRow + 6) + extraRows);
        const cols = Math.min(COLS_CAP, Math.max(COLS_MIN, sheet.maxCol + 2) + extraCols);
        const table = document.createElement("table");
        table.className = "sheetedit-table";
        // Column widths (table-layout is fixed, so these are authoritative). The table is
        // sized to the sum so columns keep their width and the grid scrolls horizontally,
        // rather than the table shrinking to the viewport and squashing every column.
        const colgroup = document.createElement("colgroup");
        const rnCol = document.createElement("col");
        rnCol.style.width = "44px";
        colgroup.appendChild(rnCol);
        let totalW = 44;
        const colEls = [];
        for (let c = 1; c <= cols; c++) {
            const w = sheet.colWidths?.get(c) ?? 96;
            const col = document.createElement("col");
            col.style.width = `${w}px`;
            colgroup.appendChild(col);
            colEls.push(col);
            totalW += w;
        }
        table.appendChild(colgroup);
        table.style.width = `${totalW}px`;
        const head = document.createElement("tr");
        const corner = document.createElement("th");
        corner.className = "corner";
        corner.title = "Select all";
        corner.addEventListener("click", () => setSel(1, 1, rows, cols));
        head.appendChild(corner);
        for (let c = 1; c <= cols; c++) {
            const th = document.createElement("th");
            th.className = "colhead";
            th.textContent = colToLetters(c);
            th.title = `Select column ${colToLetters(c)} (drag the right edge to resize)`;
            th.addEventListener("click", () => {
                if (resizing)
                    return;
                anchor = { r: 1, c };
                setSel(1, c, rows, c);
            });
            const grip = document.createElement("div");
            grip.className = "sheetedit-colgrip";
            const colEl = colEls[c - 1];
            grip.addEventListener("pointerdown", (e) => startColResize(e, c, colEl, sheet.colWidths?.get(c) ?? 96));
            th.appendChild(grip);
            head.appendChild(th);
        }
        table.appendChild(head);
        // Merged ranges: the top-left cell spans; covered cells are not rendered.
        const covered = new Set();
        const spanAt = new Map();
        for (const m of sheet.merges ?? []) {
            spanAt.set(key(m.r1, m.c1), { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
            for (let r = m.r1; r <= m.r2; r++)
                for (let c = m.c1; c <= m.c2; c++)
                    if (r !== m.r1 || c !== m.c1)
                        covered.add(key(r, c));
        }
        for (let r = 1; r <= rows; r++) {
            const tr = document.createElement("tr");
            const rh = sheet.rowHeights?.get(r);
            if (rh)
                tr.style.height = `${rh}px`;
            const rn = document.createElement("th");
            rn.className = "rownum";
            rn.textContent = String(r);
            rn.title = `Select row ${r} (drag the bottom edge to resize)`;
            rn.addEventListener("click", () => {
                if (resizing)
                    return;
                anchor = { r, c: 1 };
                setSel(r, 1, r, cols);
            });
            const rgrip = document.createElement("div");
            rgrip.className = "sheetedit-rowgrip";
            rgrip.addEventListener("pointerdown", (e) => startRowResize(e, r, tr, rh ?? 22));
            rn.appendChild(rgrip);
            tr.appendChild(rn);
            for (let c = 1; c <= cols; c++) {
                if (covered.has(key(r, c)))
                    continue; // part of a merge; the top-left cell spans it
                const td = document.createElement("td");
                td.dataset.rc = key(r, c);
                tds.set(key(r, c), td);
                const sp = spanAt.get(key(r, c));
                if (sp) {
                    if (sp.rs > 1)
                        td.rowSpan = sp.rs;
                    if (sp.cs > 1)
                        td.colSpan = sp.cs;
                }
                const cell = getCell(sheet, r, c);
                if (cell?.kind === "n")
                    td.classList.add("num");
                const input = document.createElement("input");
                input.type = "text";
                input.value = cellDisplay(cell);
                input.setAttribute("aria-label", `${colToLetters(c)}${r}`);
                // Apply the file's visual style (fill/borders on the cell, font/colour/align on the text).
                const cs = cell?.cellStyle;
                if (cs) {
                    if (cs.bg)
                        td.style.background = cs.bg;
                    if (cs.borders) {
                        // Override the default gridline box-shadow: keep light right/bottom unless the
                        // file specifies a border there, and add the file's top/left where present.
                        const bd = cs.borders;
                        const g = "#e3e3e6";
                        const sh = [`inset -1px 0 0 0 ${bd.right ?? g}`, `inset 0 -1px 0 0 ${bd.bottom ?? g}`];
                        if (bd.top)
                            sh.push(`inset 0 1px 0 0 ${bd.top}`);
                        if (bd.left)
                            sh.push(`inset 1px 0 0 0 ${bd.left}`);
                        td.style.boxShadow = sh.join(", ");
                    }
                    if (cs.bold)
                        input.style.fontWeight = "700";
                    if (cs.italic)
                        input.style.fontStyle = "italic";
                    if (cs.color)
                        input.style.color = cs.color;
                    if (cs.align)
                        input.style.textAlign = cs.align;
                }
                const ki = key(r, c);
                // Shift-click extends the selection from the anchor (no caret/edit).
                input.addEventListener("mousedown", (e) => {
                    if (e.shiftKey) {
                        e.preventDefault();
                        selectCell(r, c, true);
                    }
                });
                input.addEventListener("focus", () => {
                    if (justDragged) {
                        input.blur(); // a range was just drag-selected; do not enter edit on the trailing tap
                        return;
                    }
                    selectCell(r, c, false); // tapping a cell selects it; toolbar styles target the selection
                    const live = getCell(sheet, r, c);
                    if (!live)
                        return;
                    // Show the editable underlying value (formula or raw), not the formatted display.
                    input.value = live.formula != null ? "=" + live.formula : live.value;
                });
                const commit = () => {
                    const raw = input.value;
                    const live = getCell(sheet, r, c);
                    const before = live ? (live.formula != null ? "=" + live.formula : live.value) : "";
                    if (raw === before) {
                        input.value = displayValue(sheet, r, c);
                        return;
                    }
                    setCellInput(sheet, r, c, raw);
                    recalc(wb);
                    mark();
                    refreshDisplays(sheet);
                    input.value = displayValue(sheet, r, c);
                    input.parentElement?.classList.toggle("num", getCell(sheet, r, c)?.kind === "n");
                };
                input.addEventListener("blur", commit);
                input.addEventListener("keydown", (e) => {
                    if (e.key === "Enter") {
                        e.preventDefault();
                        input.blur();
                        const below = inputs.get(key(r + 1, c));
                        below?.focus();
                    }
                    else if (e.key === "Escape") {
                        input.value = displayValue(sheet, r, c);
                        input.blur();
                    }
                });
                td.appendChild(input);
                tr.appendChild(td);
                inputs.set(ki, input);
            }
            table.appendChild(tr);
        }
        gridScroll.appendChild(table);
        paintSel(); // restore the selection highlight after a re-render
    };
    const renderTabs = () => {
        tabs.innerHTML = "";
        wb.sheets.forEach((sheet, i) => {
            const b = document.createElement("button");
            b.type = "button";
            b.className = "sheetedit-tab";
            b.textContent = sheet.name;
            b.setAttribute("role", "tab");
            b.setAttribute("aria-selected", String(i === active));
            b.addEventListener("click", () => {
                if (i === active)
                    return;
                active = i;
                extraRows = 0; // each sheet starts at its own extent
                extraCols = 0;
                sel = null;
                anchor = null;
                renderTabs();
                renderGrid();
            });
            tabs.appendChild(b);
        });
    };
    buildToolbar();
    renderTabs();
    renderGrid();
    return {
        isDirty() {
            return dirty;
        },
        async getBytes() {
            return dirty ? writeWorkbook(wb) : original.slice();
        },
        destroy() {
            wrap.remove();
        },
    };
}
