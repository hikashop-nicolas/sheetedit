// Self-contained i18n for sheetedit so the library is a complete multilingual product on
// its own. Detects the locale from the browser / device preferred-languages list (base
// language, first match), English fallback. Adding a language = add a dict to LOCALES;
// hosts may force one via setLocale().
const en = {
    sheets: "Sheets",
    addRow: "+ Row",
    addCol: "+ Col",
    addRows: "Add rows",
    addCols: "Add columns",
    bold: "Bold",
    italic: "Italic",
    textColour: "Text colour",
    fillColour: "Fill colour",
    alignLeft: "Align left",
    alignCentre: "Align centre",
    alignRight: "Align right",
    borders: "Borders",
    merge: "Merge or unmerge cells",
    borderAll: "All borders",
    borderOuter: "Outer border",
    borderTop: "Top",
    borderBottom: "Bottom",
    borderLeft: "Left",
    borderRight: "Right",
    borderNone: "No border",
    selectAll: "Select all",
    selectColumn: "Select column {col} (drag the right edge to resize)",
    selectRow: "Select row {row} (drag the bottom edge to resize)",
};
const fr = {
    sheets: "Feuilles",
    addRow: "+ Ligne",
    addCol: "+ Col.",
    addRows: "Ajouter des lignes",
    addCols: "Ajouter des colonnes",
    bold: "Gras",
    italic: "Italique",
    textColour: "Couleur du texte",
    fillColour: "Couleur de remplissage",
    alignLeft: "Aligner à gauche",
    alignCentre: "Centrer",
    alignRight: "Aligner à droite",
    borders: "Bordures",
    merge: "Fusionner ou défusionner les cellules",
    borderAll: "Toutes les bordures",
    borderOuter: "Bordure extérieure",
    borderTop: "Haut",
    borderBottom: "Bas",
    borderLeft: "Gauche",
    borderRight: "Droite",
    borderNone: "Aucune bordure",
    selectAll: "Tout sélectionner",
    selectColumn: "Sélectionner la colonne {col} (glisser le bord droit pour redimensionner)",
    selectRow: "Sélectionner la ligne {row} (glisser le bord inférieur pour redimensionner)",
};
const LOCALES = { en, fr };
let active = null;
function detect() {
    const prefs = (typeof navigator !== "undefined" && navigator.languages) || ["en"];
    for (const tag of prefs) {
        const base = tag.toLowerCase().split("-")[0];
        if (LOCALES[base])
            return LOCALES[base];
    }
    return en;
}
/** Force a locale (host escape hatch). Unknown codes fall back to English. */
export function setLocale(code) {
    active = LOCALES[code.toLowerCase().split("-")[0]] ?? en;
}
export function t(key, params) {
    if (!active)
        active = detect();
    let s = active[key] ?? en[key] ?? key;
    if (params)
        s = s.replace(/\{(\w+)\}/g, (_, k) => (k in params ? String(params[k]) : `{${k}}`));
    return s;
}
