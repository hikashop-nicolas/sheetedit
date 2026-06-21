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
function readXlsxStyles(file) {
    const customFmt = new Map();
    const xfNumFmtIds = [];
    if (!file)
        return { customFmt, xfNumFmtIds };
    const doc = parseXml(file);
    for (const nf of Array.from(doc.getElementsByTagName("numFmt"))) {
        const id = Number(nf.getAttribute("numFmtId"));
        const code = nf.getAttribute("formatCode");
        if (Number.isFinite(id) && code != null)
            customFmt.set(id, code);
    }
    // The cell @s indexes <cellXfs>, not <cellStyleXfs>; read that list specifically.
    const cellXfs = doc.getElementsByTagName("cellXfs")[0];
    if (cellXfs) {
        for (const xf of Array.from(cellXfs.children)) {
            if (xf.localName === "xf")
                xfNumFmtIds.push(Number(xf.getAttribute("numFmtId") || "0"));
        }
    }
    return { customFmt, xfNumFmtIds };
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
    const styles = readXlsxStyles(files["xl/styles.xml"]);
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
        if (touched && sheet.path)
            wb.files[sheet.path] = serializeXml(sheet.doc);
    }
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
            res = "#ERROR!";
        }
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
const ROWS_CAP = 400;
const COLS_CAP = 60;
const STYLE_ID = "sheetedit-style";
function injectStyles() {
    if (document.getElementById(STYLE_ID))
        return;
    const s = document.createElement("style");
    s.id = STYLE_ID;
    s.textContent = `
    .sheetedit-wrap { display:flex; flex-direction:column; height:100%; background:#1f2227; color:#e6e6e6; font:13px system-ui, sans-serif; }
    .sheetedit-grid { flex:1; min-height:0; overflow:auto; background:#2a2d33; }
    table.sheetedit-table { border-collapse:collapse; font:13px/1.3 ui-sans-serif, system-ui, sans-serif; }
    .sheetedit-table th, .sheetedit-table td { border:1px solid #3a3f47; padding:0; margin:0; }
    .sheetedit-table th {
      position:sticky; top:0; z-index:2; background:#33373e; color:#aeb4bf; font-weight:600;
      padding:3px 8px; text-align:center; user-select:none; min-width:64px;
    }
    .sheetedit-table th.corner { left:0; z-index:3; }
    .sheetedit-table th.rownum { position:sticky; left:0; z-index:1; top:auto; min-width:38px; text-align:right; }
    .sheetedit-table input {
      border:0; background:transparent; color:#e6e6e6; font:inherit; padding:3px 8px;
      width:96px; box-sizing:border-box; outline:none;
    }
    .sheetedit-table td.num input { text-align:right; font-variant-numeric:tabular-nums; }
    .sheetedit-table input:focus { box-shadow:inset 0 0 0 2px #6e7bff; background:#23262b; }
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
    recalc(wb);
    const wrap = document.createElement("div");
    wrap.className = "sheetedit-wrap";
    const gridScroll = document.createElement("div");
    gridScroll.className = "sheetedit-grid";
    const tabs = document.createElement("div");
    tabs.className = "sheetedit-tabs";
    tabs.setAttribute("role", "tablist");
    tabs.setAttribute("aria-label", "Sheets");
    wrap.append(gridScroll, tabs);
    container.appendChild(wrap);
    let active = 0;
    let inputs = new Map();
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
        gridScroll.innerHTML = "";
        const rows = Math.min(ROWS_CAP, Math.max(ROWS_MIN, sheet.maxRow + 6));
        const cols = Math.min(COLS_CAP, Math.max(COLS_MIN, sheet.maxCol + 2));
        const table = document.createElement("table");
        table.className = "sheetedit-table";
        const head = document.createElement("tr");
        const corner = document.createElement("th");
        corner.className = "corner";
        head.appendChild(corner);
        for (let c = 1; c <= cols; c++) {
            const th = document.createElement("th");
            th.textContent = colToLetters(c);
            head.appendChild(th);
        }
        table.appendChild(head);
        for (let r = 1; r <= rows; r++) {
            const tr = document.createElement("tr");
            const rn = document.createElement("th");
            rn.className = "rownum";
            rn.textContent = String(r);
            tr.appendChild(rn);
            for (let c = 1; c <= cols; c++) {
                const td = document.createElement("td");
                const cell = getCell(sheet, r, c);
                if (cell?.kind === "n")
                    td.classList.add("num");
                const input = document.createElement("input");
                input.type = "text";
                input.value = cellDisplay(cell);
                input.setAttribute("aria-label", `${colToLetters(c)}${r}`);
                const ki = key(r, c);
                input.addEventListener("focus", () => {
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
                renderTabs();
                renderGrid();
            });
            tabs.appendChild(b);
        });
    };
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
