// Line icons for the Power Query editor, matching sheetedit's toolbar style (16x16 viewBox,
// stroke="currentColor"). Inner SVG markup only; the button helper wraps it in an <svg>.

/** Ribbon transform icons, keyed by TransformSpec.id. */
export const TRANSFORM_ICONS: Record<string, string> = {
  removeColumns: `<rect x="2.5" y="2.5" width="11" height="11" rx="1.5"/><path d="M5.5 8h5"/>`,
  chooseColumns: `<rect x="2.5" y="2.5" width="11" height="11" rx="1.5"/><path d="M5.3 8.2 7.1 10 10.7 6"/>`,
  renameColumn: `<path d="M11.2 3 13 4.8l-7 7-2.3.5.5-2.3z"/>`,
  filterRows: `<path d="M3 4h10l-3.8 4.5v3.7L6.8 13V8.5z"/>`,
  sort: `<path d="M5 4v8M3 10l2 2 2-2M11 12V4M9 6l2-2 2 2"/>`,
  keepTop: `<rect x="2.5" y="2.8" width="11" height="3" rx="1"/><path d="M4 8.8h8M4 11.8h8"/>`,
  keepBottom: `<rect x="2.5" y="10.2" width="11" height="3" rx="1"/><path d="M4 4.2h8M4 7.2h8"/>`,
  removeTop: `<path d="M4 3.8h8"/><rect x="2.5" y="7" width="11" height="6.5" rx="1"/>`,
  removeDuplicates: `<rect x="2.5" y="2.5" width="8" height="8" rx="1.2"/><rect x="5.5" y="5.5" width="8" height="8" rx="1.2"/>`,
  reverse: `<path d="M4 5h6L8 3M12 11H6l2 2"/>`,
  changeType: `<path d="M3.5 4.4V3.4h6v1M6.5 3.4v8.2M5 11.6h3"/><path d="M11 8.8l2 2 2-2M13 4v6.8"/>`,
  replaceValues: `<path d="M3.5 5h6L7.5 3M12.5 11h-6l2 2"/>`,
  splitColumn: `<rect x="5.5" y="2.5" width="5" height="3.6" rx="1"/><path d="M8 6.1v3.2M8 9.3H4.5v3.7M8 9.3h3.5v3.7"/>`,
  transpose: `<path d="M3 3h5M3 3v5M3 3l5 5"/><path d="M13 13H8M13 13V8M13 13 8 8"/>`,
  promoteHeaders: `<rect x="2.5" y="2.5" width="11" height="3.4" rx="1"/><path d="M8 13.5V8M5.6 10.4 8 8l2.4 2.4"/>`,
  unpivotOthers: `<rect x="2.5" y="2.5" width="11" height="11" rx="1"/><path d="M6 2.5v11M6 6.5h7M6 10h7"/>`,
  groupBy: `<path d="M6 3H4.6C3.7 3 3.1 3.6 3.1 4.5v2C3.1 7.3 2.7 8 2.1 8c.6 0 1 .7 1 1.5v2c0 .9.6 1.5 1.5 1.5H6"/><path d="M9 6.5h5M9 9.5h5"/>`,
  customColumn: `<rect x="2.5" y="2.5" width="6" height="11" rx="1"/><path d="M12 5v6M9 8h6"/>`,
  indexColumn: `<path d="M3 3.6h1.5v8.8M3 12.4h3"/><path d="M8 5.6h5M8 8h5M8 10.4h5"/>`,
};

/** Cross-query (Combine) icons. */
export const APPEND_ICON = `<rect x="2.5" y="2.5" width="11" height="4" rx="1"/><rect x="2.5" y="9.5" width="11" height="4" rx="1"/><path d="M8 6.6v2.8M6.8 8.2 8 9.4 9.2 8.2"/>`;
export const MERGE_ICON = `<circle cx="6" cy="8" r="3.3"/><circle cx="10" cy="8" r="3.3"/>`;

/** Title-bar and pane icons. */
export const LOAD_ICON = `<path d="M8 2.5v7.5M5 7l3 3 3-3M3 13.5h10"/>`;
export const CANCEL_ICON = `<path d="M4 4l8 8M12 4l-8 8"/>`;
export const SAVE_ICON = `<path d="M3.5 3h7l3 3v6.5a.9.9 0 0 1-.9.9H3.4a.9.9 0 0 1-.9-.9V3.9A.9.9 0 0 1 3.5 3z"/><path d="M5.5 3v3h4.5M5.5 13v-4h5v4"/>`;
export const NEWQUERY_ICON = `<path d="M8 3.5v9M3.5 8h9"/>`;

const FALLBACK = `<rect x="2.5" y="2.5" width="11" height="11" rx="1.5"/>`;

/** Build an <svg> element from inner markup, sized for a button. */
export function svgIcon(inner: string, size = 15): SVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("width", String(size));
  svg.setAttribute("height", String(size));
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.4");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  svg.setAttribute("aria-hidden", "true");
  svg.innerHTML = inner || FALLBACK;
  return svg;
}
