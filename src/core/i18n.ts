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
  rowInsAbove: "Insert row above",
  rowsInsAbove: "Insert {n} rows above",
  rowInsBelow: "Insert row below",
  rowsInsBelow: "Insert {n} rows below",
  rowDelOne: "Delete row",
  rowsDel: "Delete {n} rows",
  colInsBefore: "Insert column left",
  colsInsBefore: "Insert {n} columns left",
  colInsAfter: "Insert column right",
  colsInsAfter: "Insert {n} columns right",
  colDelOne: "Delete column",
  colsDel: "Delete {n} columns",
  convertXlsx: "Convert to XLSX",
  convertXlsxTitle: "Convert this file to a real workbook (styles, formats, sheets)",
  findReplace: "Find and replace (Ctrl+F)",
  frFind: "Find",
  frReplace: "Replace with",
  frPrev: "Previous match",
  frNext: "Next match",
  frReplaceOne: "Replace",
  frReplaceOneTitle: "Replace the current match",
  frReplaceAll: "Replace all",
  frReplaceAllTitle: "Replace every match on this sheet",
  frClose: "Close",
  frCount: "{i} / {n}",
  frNone: "No matches",
  calcName: "Unknown function name: this formula cannot be computed, the file's saved value is shown.",
  calcEval: "This formula could not be evaluated; the file's saved value is shown.",
  calcCircular: "Circular reference: this cell depends on its own result; its value may be stale.",
  numFormat: "Number format",
  fmtGeneral: "Automatic",
  fmtNumber: "Number (0.00)",
  fmtThousands: "Thousands (1,234.00)",
  fmtPercent: "Percent (%)",
  fmtCurrency: "Currency",
  fmtDate: "Date",
  fmtDateTime: "Date and time",
  fmtTime: "Time",
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
  rowInsAbove: "Insérer une ligne au-dessus",
  rowsInsAbove: "Insérer {n} lignes au-dessus",
  rowInsBelow: "Insérer une ligne en dessous",
  rowsInsBelow: "Insérer {n} lignes en dessous",
  rowDelOne: "Supprimer la ligne",
  rowsDel: "Supprimer {n} lignes",
  colInsBefore: "Insérer une colonne à gauche",
  colsInsBefore: "Insérer {n} colonnes à gauche",
  colInsAfter: "Insérer une colonne à droite",
  colsInsAfter: "Insérer {n} colonnes à droite",
  colDelOne: "Supprimer la colonne",
  colsDel: "Supprimer {n} colonnes",
  convertXlsx: "Convertir en XLSX",
  convertXlsxTitle: "Convertir ce fichier en vrai classeur (styles, formats, feuilles)",
  findReplace: "Rechercher et remplacer (Ctrl+F)",
  frFind: "Rechercher",
  frReplace: "Remplacer par",
  frPrev: "Résultat précédent",
  frNext: "Résultat suivant",
  frReplaceOne: "Remplacer",
  frReplaceOneTitle: "Remplacer le résultat courant",
  frReplaceAll: "Tout remplacer",
  frReplaceAllTitle: "Remplacer tous les résultats de cette feuille",
  frClose: "Fermer",
  frCount: "{i} / {n}",
  frNone: "Aucun résultat",
  calcName: "Nom de fonction inconnu : cette formule ne peut pas être calculée, la valeur enregistrée du fichier est affichée.",
  calcEval: "Cette formule n'a pas pu être évaluée ; la valeur enregistrée du fichier est affichée.",
  calcCircular: "Référence circulaire : cette cellule dépend de son propre résultat ; sa valeur peut être obsolète.",
  numFormat: "Format de nombre",
  fmtGeneral: "Automatique",
  fmtNumber: "Nombre (0,00)",
  fmtThousands: "Milliers (1 234,00)",
  fmtPercent: "Pourcentage (%)",
  fmtCurrency: "Monnaie",
  fmtDate: "Date",
  fmtDateTime: "Date et heure",
  fmtTime: "Heure",
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
  rowInsAbove: "上に行を挿入",
  rowsInsAbove: "上に {n} 行を挿入",
  rowInsBelow: "下に行を挿入",
  rowsInsBelow: "下に {n} 行を挿入",
  rowDelOne: "行を削除",
  rowsDel: "{n} 行を削除",
  colInsBefore: "左に列を挿入",
  colsInsBefore: "左に {n} 列を挿入",
  colInsAfter: "右に列を挿入",
  colsInsAfter: "右に {n} 列を挿入",
  colDelOne: "列を削除",
  colsDel: "{n} 列を削除",
  convertXlsx: "XLSX に変換",
  convertXlsxTitle: "このファイルを本来のワークブック（スタイル・書式・シート対応）に変換",
  findReplace: "検索と置換 (Ctrl+F)",
  frFind: "検索",
  frReplace: "置換後の文字列",
  frPrev: "前の結果",
  frNext: "次の結果",
  frReplaceOne: "置換",
  frReplaceOneTitle: "現在の結果を置換",
  frReplaceAll: "すべて置換",
  frReplaceAllTitle: "このシートのすべての結果を置換",
  frClose: "閉じる",
  frCount: "{i} / {n}",
  frNone: "一致なし",
  calcName: "不明な関数名：この数式は計算できないため、ファイルに保存された値を表示しています。",
  calcEval: "この数式は評価できませんでした。ファイルに保存された値を表示しています。",
  calcCircular: "循環参照：このセルは自身の結果に依存しているため、値が古い可能性があります。",
  numFormat: "表示形式",
  fmtGeneral: "自動",
  fmtNumber: "数値 (0.00)",
  fmtThousands: "桁区切り (1,234.00)",
  fmtPercent: "パーセント (%)",
  fmtCurrency: "通貨",
  fmtDate: "日付",
  fmtDateTime: "日付と時刻",
  fmtTime: "時刻",
};

const LOCALES: Record<string, Dict> = { en, fr, ja };

let active: Dict | null = null;
let activeCode = "en";

function detect(): Dict {
  const prefs = (typeof navigator !== "undefined" && navigator.languages) || ["en"];
  for (const tag of prefs) {
    const base = tag.toLowerCase().split("-")[0]!;
    if (LOCALES[base]) {
      activeCode = base;
      return LOCALES[base]!;
    }
  }
  activeCode = "en";
  return en;
}

/** Force a locale (host escape hatch). Unknown codes fall back to English. */
export function setLocale(code: string): void {
  const base = code.toLowerCase().split("-")[0]!;
  active = LOCALES[base] ?? en;
  activeCode = LOCALES[base] ? base : "en";
}

/** The active base locale code ("en", "fr", "ja"): number/date input conventions. */
export function localeCode(): string {
  if (!active) active = detect();
  return activeCode;
}

export function t(key: string, params?: Record<string, string | number>): string {
  if (!active) active = detect();
  let s = active[key] ?? en[key] ?? key;
  if (params) s = s.replace(/\{(\w+)\}/g, (_, k: string) => (k in params ? String(params[k]) : `{${k}}`));
  return s;
}
