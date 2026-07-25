import { parseXmlOpt, serializeXml, type Workbook } from "../../core/model";

// Freeze panes live in ODF's view settings (settings.xml), not in content.xml: per sheet, a
// config-item-map-entry carrying HorizontalSplitMode / VerticalSplitMode (2 = frozen, 1 = split,
// 0 = none) with the counts in *SplitPosition, plus PositionRight / PositionBottom for the first
// cell of the scrolling pane. A workbook with no settings.xml gets a minimal one.

const CFG = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
const OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

const EMPTY_SETTINGS =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<office:document-settings xmlns:office="${OFFICE}" xmlns:config="${CFG}" office:version="1.3">` +
  `<office:settings><config:config-item-set config:name="ooo:view-settings">` +
  `<config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>` +
  `<config:config-item-map-named config:name="Tables"/>` +
  `</config:config-item-map-entry></config:config-item-map-indexed>` +
  `</config:config-item-set></office:settings></office:document-settings>`;

const childByName = (parent: Element, local: string, name: string): Element | undefined =>
  Array.from(parent.children).find((e) => e.localName === local && e.getAttribute("config:name") === name);

/** Set one config-item (int) on a table's settings entry. */
function setItem(doc: Document, entry: Element, name: string, value: number): void {
  let item = childByName(entry, "config-item", name);
  if (!item) {
    item = doc.createElementNS(CFG, "config:config-item");
    item.setAttribute("config:name", name);
    item.setAttribute("config:type", "int");
    entry.appendChild(item);
  }
  item.textContent = String(value);
}

/** The <config:config-item-map-named config:name="Tables"> node, created along the way if needed. */
function tablesNode(doc: Document): Element | undefined {
  const named = Array.from(doc.getElementsByTagName("*")).find(
    (e) => e.localName === "config-item-map-named" && e.getAttribute("config:name") === "Tables",
  );
  if (named) return named;
  // A settings.xml without a Views entry: hang one off the first config-item-set.
  const set = Array.from(doc.getElementsByTagName("*")).find((e) => e.localName === "config-item-set");
  if (!set) return undefined;
  const indexed = doc.createElementNS(CFG, "config:config-item-map-indexed");
  indexed.setAttribute("config:name", "Views");
  const entry = doc.createElementNS(CFG, "config:config-item-map-entry");
  const tables = doc.createElementNS(CFG, "config:config-item-map-named");
  tables.setAttribute("config:name", "Tables");
  entry.appendChild(tables);
  indexed.appendChild(entry);
  set.appendChild(indexed);
  return tables;
}

/** Register settings.xml in the manifest when it was not there before. */
function addManifestEntry(wb: Workbook): void {
  const m = wb.files["META-INF/manifest.xml"];
  const doc = m ? parseXmlOpt(m) : undefined;
  if (!doc) return;
  const has = Array.from(doc.getElementsByTagName("*")).some(
    (e) => e.localName === "file-entry" && e.getAttribute("manifest:full-path") === "settings.xml",
  );
  if (has) return;
  const fe = doc.createElementNS(MANIFEST, "manifest:file-entry");
  fe.setAttribute("manifest:full-path", "settings.xml");
  fe.setAttribute("manifest:media-type", "text/xml");
  doc.documentElement.appendChild(fe);
  wb.files["META-INF/manifest.xml"] = serializeXml(doc);
}

/** Persist every sheet whose freeze changed into settings.xml. */
export function writeOdsFreezes(wb: Workbook): void {
  if (!wb.sheets.some((s) => s.freezeDirty)) return;
  const had = !!wb.files["settings.xml"];
  const doc = parseXmlOpt(wb.files["settings.xml"] ?? new TextEncoder().encode(EMPTY_SETTINGS));
  const tables = doc ? tablesNode(doc) : undefined;
  if (!doc || !tables) return;
  for (const sheet of wb.sheets) {
    if (!sheet.freezeDirty) continue;
    let entry = childByName(tables, "config-item-map-entry", sheet.name);
    if (!entry) {
      entry = doc.createElementNS(CFG, "config:config-item-map-entry");
      entry.setAttribute("config:name", sheet.name);
      tables.appendChild(entry);
    }
    const rows = sheet.freeze?.rows ?? 0, cols = sheet.freeze?.cols ?? 0;
    setItem(doc, entry, "HorizontalSplitMode", cols > 0 ? 2 : 0);
    setItem(doc, entry, "VerticalSplitMode", rows > 0 ? 2 : 0);
    setItem(doc, entry, "HorizontalSplitPosition", cols);
    setItem(doc, entry, "VerticalSplitPosition", rows);
    // Which pane holds the cursor: the one past the frozen lines.
    setItem(doc, entry, "PositionRight", cols);
    setItem(doc, entry, "PositionBottom", rows);
    setItem(doc, entry, "ActiveSplitRange", rows > 0 && cols > 0 ? 3 : rows > 0 ? 2 : cols > 0 ? 3 : 2);
    sheet.freezeDirty = false;
  }
  wb.files["settings.xml"] = serializeXml(doc);
  if (!had) addManifestEntry(wb);
}
