// Self-contained i18n for sheetedit so the library is a complete multilingual product on
// its own. Detects the locale from the browser / device preferred-languages list (base
// language, first match), English fallback. Adding a language = add a dict to LOCALES;
// hosts may force one via setLocale().

type Dict = Record<string, string>;

const en: Dict = {
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
  openFailed: "This file could not be opened as a spreadsheet. It may be corrupted or password-protected. Saving returns the file unchanged.",
  formulaBar: "Formula",
  activeCell: "Active cell",
  insertSum: "Insert SUM of a range",
  functions: "Functions",
  pickRange: "Select a range of cells…",
  formatting: "Formatting",
  undo: "Undo (Ctrl+Z)",
  redo: "Redo (Ctrl+Shift+Z)",
};

const fr: Dict = {
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
  openFailed: "Ce fichier n'a pas pu être ouvert comme classeur. Il est peut-être corrompu ou protégé par mot de passe. L'enregistrement renvoie le fichier inchangé.",
  formulaBar: "Formule",
  activeCell: "Cellule active",
  insertSum: "Insérer la somme d'une plage",
  functions: "Fonctions",
  pickRange: "Sélectionnez une plage de cellules…",
  formatting: "Mise en forme",
  undo: "Annuler (Ctrl+Z)",
  redo: "Rétablir (Ctrl+Maj+Z)",
};

const ja: Dict = {
  sheets: "シート",
  addRow: "+ 行",
  addCol: "+ 列",
  addRows: "行を追加",
  addCols: "列を追加",
  bold: "太字",
  italic: "斜体",
  textColour: "文字の色",
  fillColour: "塗りつぶしの色",
  alignLeft: "左揃え",
  alignCentre: "中央揃え",
  alignRight: "右揃え",
  borders: "罫線",
  merge: "セルの結合/結合解除",
  borderAll: "格子",
  borderOuter: "外枠",
  borderTop: "上",
  borderBottom: "下",
  borderLeft: "左",
  borderRight: "右",
  borderNone: "罫線なし",
  selectAll: "すべて選択",
  selectColumn: "列 {col} を選択（右端をドラッグでサイズ変更）",
  selectRow: "行 {row} を選択（下端をドラッグでサイズ変更）",
  openFailed: "このファイルはスプレッドシートとして開けませんでした。破損しているか、パスワードで保護されている可能性があります。保存してもファイルは変更されません。",
  formulaBar: "数式",
  activeCell: "アクティブセル",
  insertSum: "範囲の合計を挿入",
  functions: "関数",
  pickRange: "セル範囲を選択…",
  formatting: "書式設定",
  undo: "元に戻す (Ctrl+Z)",
  redo: "やり直し (Ctrl+Shift+Z)",
};

const LOCALES: Record<string, Dict> = { en, fr, ja };

let active: Dict | null = null;

function detect(): Dict {
  const prefs = (typeof navigator !== "undefined" && navigator.languages) || ["en"];
  for (const tag of prefs) {
    const base = tag.toLowerCase().split("-")[0]!;
    if (LOCALES[base]) return LOCALES[base]!;
  }
  return en;
}

/** Force a locale (host escape hatch). Unknown codes fall back to English. */
export function setLocale(code: string): void {
  active = LOCALES[code.toLowerCase().split("-")[0]!] ?? en;
}

export function t(key: string, params?: Record<string, string | number>): string {
  if (!active) active = detect();
  let s = active[key] ?? en[key] ?? key;
  if (params) s = s.replace(/\{(\w+)\}/g, (_, k: string) => (k in params ? String(params[k]) : `{${k}}`));
  return s;
}
