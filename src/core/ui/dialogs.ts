import { colToLetters, getCell, parseA1Ref, type ShapeGeom, type Sheet, type SheetShape, type Workbook } from "../model";
import { t } from "../i18n";
import { formDialog, type FormField } from "./form-dialog";
import { setOdsComment, setOdsCondFormat, setOdsDataValidation, setOdsHyperlink, setOdsSparklineGroup } from "../../adapters/ods";
import { setXlsxComment, setXlsxCondFormat, setXlsxDataValidation, setXlsxHyperlink, setXlsxSparklineGroup } from "../../adapters/xlsx";
import type { CfSpec } from "../../adapters/xlsx/write";

// The single-cell / selection authoring dialogs (hyperlink, data validation, conditional formatting,
// sparkline, shape, note). Extracted from the editor god-closure; each is a plain function of a small
// context. formDialog is the shared modal builder.

type Rect = { r1: number; c1: number; r2: number; c2: number };
type SparkSpec = { type: "line" | "column" | "stacked"; color: string; negColor?: string; dataRef: string } | null;

export interface DialogCtx {
  wb: Workbook;
  active: () => number;
  wrap: HTMLElement;
  getSelRect: () => Rect;
  mark: () => void;
  renderGrid: () => void;
  refreshShapes: () => void;
  applySparkline: (sheet: Sheet, host: { r: number; c: number }, spec: SparkSpec) => void;
}

/** Parse an A1 range ("B2:F5") or single cell ("G2") into a normalised rect, or null. */
function parseRange(raw: string): Rect | null {
  const body = raw.trim().replace(/\$/g, "");
  const parts = body.split(":");
  const a = parseA1Ref(parts[0] ?? "");
  if (!a) return null;
  const b = parts[1] ? parseA1Ref(parts[1]) : a;
  if (!b) return null;
  return { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
}

export function setupDialogs(ctx: DialogCtx): {
  openLinkDialog: () => void;
  openDvDialog: () => void;
  openCfDialog: () => void;
  openSparkDialog: () => void;
  openShapeDialog: (existing?: SheetShape) => void;
  openNoteDialog: () => void;
} {
  const { wb } = ctx;
  const dialog = (title: string, fields: FormField[], onOk: (vals: Record<string, string | boolean>) => void): void => formDialog(ctx.wrap, title, fields, onOk);
  const sheetNow = (): Sheet => wb.sheets[ctx.active()]!;

  const openLinkDialog = (): void => {
    const s = ctx.getSelRect(); const r = s.r1, c = s.c1; const sheet = sheetNow();
    const cur = getCell(sheet, r, c)?.link;
    dialog(t("linkEdit"), [
      { key: "href", label: t("linkTarget"), type: "text", value: cur?.href ?? "" },
      { key: "tip", label: t("linkTip"), type: "text", value: cur?.tip ?? "" },
      { key: "internal", label: t("linkInternal"), type: "checkbox", value: !!cur?.internal },
    ], (v) => {
      const href = String(v.href).trim();
      const link = href ? { href, internal: !!v.internal, tip: String(v.tip).trim() || undefined } : null;
      if (wb.kind === "ods") setOdsHyperlink(sheet, r, c, link);
      else setXlsxHyperlink(wb, sheet, r, c, link);
      ctx.mark(); ctx.renderGrid();
    });
  };

  const openDvDialog = (): void => {
    const s = ctx.getSelRect(); const sheet = sheetNow();
    const ranges = [{ r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 }];
    dialog(t("dvEdit"), [
      { key: "values", label: t("dvList"), type: "text", value: "" },
      { key: "range", label: t("dvRange"), type: "text", value: "" },
      { key: "blank", label: t("dvAllowBlank"), type: "checkbox", value: true },
    ], (v) => {
      const values = String(v.values).split(",").map((x) => x.trim()).filter(Boolean);
      const range = String(v.range).trim();
      const spec = !values.length && !range ? null : { values: range ? undefined : values, rangeRef: range || undefined, allowBlank: !!v.blank };
      if (wb.kind === "ods") setOdsDataValidation(wb, sheet, ranges, spec);
      else setXlsxDataValidation(sheet, ranges, spec);
      ctx.mark(); ctx.renderGrid();
    });
  };

  const openCfDialog = (): void => {
    const s = ctx.getSelRect(); const sheet = sheetNow();
    const ranges = [{ r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 }];
    // The graphical rules (colour scale / data bar / icon set) and the text / rank / formula rules
    // have no interoperable ODF form (calcext is dropped by LibreOffice on external author), so ODS
    // offers only the standard highlight (cellIs, incl. between).
    const ods = wb.kind === "ods";
    const kindOpts = [
      { value: "cellIs", label: t("cfHighlight") },
      ...(ods ? [] : [
        { value: "text", label: t("cfTextRule") },
        { value: "top", label: t("cfTopBottom") },
        { value: "average", label: t("cfAverageRule") },
        { value: "dupUnique", label: t("cfDupUnique") },
        { value: "expression", label: t("cfFormulaRule") },
        { value: "timePeriod", label: t("cfTimePeriodRule") },
        { value: "colorScale", label: t("cfColorScale") },
        { value: "dataBar", label: t("cfDataBar") },
        { value: "iconSet", label: t("cfIconSet") },
      ]),
    ];
    const periodOpts = ["today", "yesterday", "tomorrow", "last7Days", "thisWeek", "lastWeek", "nextWeek", "thisMonth", "lastMonth", "nextMonth"]
      .map((p) => ({ value: p, label: t(("cfPeriod_" + p) as Parameters<typeof t>[0]) }));
    const opOpts = [["greaterThan", "> "], ["lessThan", "< "], ["equal", "= "], ["notEqual", "≠ "], ["greaterThanOrEqual", "≥ "], ["lessThanOrEqual", "≤ "], ["between", t("cfBetween")], ["notBetween", t("cfNotBetween")]];
    const textOps = [["containsText", t("cfContains")], ["notContainsText", t("cfNotContains")], ["beginsWith", t("cfBegins")], ["endsWith", t("cfEnds")]];
    const iconSets = ["3TrafficLights1", "3Arrows", "3Symbols", "3Flags", "4Arrows", "4Rating", "5Arrows", "5Quarters", "5Rating"];
    const hasFill = ["cellIs", "text", "top", "average", "dupUnique", "expression", "timePeriod"];
    dialog(t("cfEdit"), [
      { key: "kind", label: t("cfKind"), type: "select", value: "cellIs", options: kindOpts },
      { key: "operator", label: t("cfOperator"), type: "select", value: "greaterThan", options: opOpts.map(([v, l]) => ({ value: v!, label: l! })), showFor: { key: "kind", values: ["cellIs"] } },
      { key: "value", label: t("cfValue"), type: "text", value: "", showFor: { key: "kind", values: ["cellIs"] } },
      { key: "value2", label: t("cfValue2"), type: "text", value: "", showFor: { key: "operator", values: ["between", "notBetween"] } },
      { key: "textOp", label: t("cfOperator"), type: "select", value: "containsText", options: textOps.map(([v, l]) => ({ value: v!, label: l! })), showFor: { key: "kind", values: ["text"] } },
      { key: "text", label: t("cfText"), type: "text", value: "", showFor: { key: "kind", values: ["text"] } },
      { key: "rank", label: t("cfRank"), type: "text", value: "10", showFor: { key: "kind", values: ["top"] } },
      { key: "bottom", label: t("cfBottom"), type: "checkbox", value: false, showFor: { key: "kind", values: ["top"] } },
      { key: "percent", label: t("cfPercent"), type: "checkbox", value: false, showFor: { key: "kind", values: ["top"] } },
      { key: "below", label: t("cfBelow"), type: "checkbox", value: false, showFor: { key: "kind", values: ["average"] } },
      { key: "equal", label: t("cfEqualAvg"), type: "checkbox", value: false, showFor: { key: "kind", values: ["average"] } },
      { key: "unique", label: t("cfUnique"), type: "checkbox", value: false, showFor: { key: "kind", values: ["dupUnique"] } },
      { key: "formula", label: t("cfFormula"), type: "text", value: "", showFor: { key: "kind", values: ["expression"] } },
      { key: "period", label: t("cfPeriod"), type: "select", value: "today", options: periodOpts, showFor: { key: "kind", values: ["timePeriod"] } },
      { key: "iconset", label: t("cfIconSet"), type: "select", value: "3TrafficLights1", options: iconSets.map((v) => ({ value: v, label: v })), showFor: { key: "kind", values: ["iconSet"] } },
      { key: "color", label: t("cfColour"), type: "color", value: "#ffc7ce", showFor: { key: "kind", values: [...hasFill, "dataBar"] } },
    ], (v) => {
      const kind = String(v.kind), color = String(v.color);
      const fill = color || "#ffc7ce";
      let spec: CfSpec;
      if (kind === "colorScale") spec = { kind: "colorScale", colors: ["#f8696b", "#ffeb84", "#63be7b"] };
      else if (kind === "dataBar") spec = { kind: "dataBar", color: color || "#638ec6" };
      else if (kind === "iconSet") { const set = String(v.iconset); spec = { kind: "iconSet", set, count: Number(set[0]) || 3 }; }
      else if (kind === "text") spec = { kind: "text", operator: String(v.textOp) as "containsText", text: String(v.text), fill };
      else if (kind === "top") spec = { kind: "top", rank: Math.max(1, Number(v.rank) || 10), percent: !!v.percent, bottom: !!v.bottom, fill };
      else if (kind === "average") spec = { kind: "average", below: !!v.below, equal: !!v.equal, fill };
      else if (kind === "dupUnique") spec = { kind: "dupUnique", unique: !!v.unique, fill };
      else if (kind === "expression") spec = { kind: "expression", formula: String(v.formula), fill };
      else if (kind === "timePeriod") spec = { kind: "timePeriod", period: String(v.period), fill };
      else { const op = String(v.operator); spec = { kind: "cellIs", operator: op, value: String(v.value), value2: op === "between" || op === "notBetween" ? String(v.value2) : undefined, fill }; }
      if (wb.kind === "ods") setOdsCondFormat(wb, sheet, ranges, spec);
      else setXlsxCondFormat(wb, sheet, ranges, spec);
      ctx.mark(); ctx.renderGrid();
    });
  };

  const openSparkDialog = (): void => {
    const s = ctx.getSelRect(); const sheet = sheetNow();
    // Editing: the single focused cell already hosts a sparkline -> prefill from it.
    const single = s.r1 === s.r2 && s.c1 === s.c2;
    const cur = single ? sheet.sparklines?.find((sp) => sp.host.r === s.r1 && sp.host.c === s.c1) : undefined;
    // Creating: default location is one cell past the selection (right of a row, below a column).
    // A 2-D selection defaults to a group: a column of hosts to the right, one per data row.
    const wide = s.c2 - s.c1 >= s.r2 - s.r1;
    const twoD = s.r2 > s.r1 && s.c2 > s.c1;
    const dataRange = cur ? cur.dataRef : `${colToLetters(s.c1)}${s.r1}:${colToLetters(s.c2)}${s.r2}`;
    const locRange = cur
      ? `${colToLetters(cur.host.c)}${cur.host.r}`
      : twoD
        ? `${colToLetters(s.c2 + 1)}${s.r1}:${colToLetters(s.c2 + 1)}${s.r2}`
        : (() => { const h = wide ? { r: s.r1, c: s.c2 + 1 } : { r: s.r2 + 1, c: s.c1 }; return `${colToLetters(h.c)}${h.r}`; })();
    dialog(t("sparkEdit"), [
      { key: "data", label: t("sparkData"), type: "text", value: dataRange },
      { key: "loc", label: t("sparkLoc"), type: "text", value: locRange },
      { key: "type", label: t("sparkType"), type: "select", value: cur?.type ?? "line", options: [{ value: "line", label: t("sparkLine") }, { value: "column", label: t("sparkColumn") }, { value: "stacked", label: t("sparkWinLoss") }] },
      { key: "color", label: t("sparkColour"), type: "color", value: cur?.color ?? "#376092" },
      // The negative-point colour only applies to column / win-loss, so it is gated on the type.
      { key: "negColor", label: t("sparkNegColour"), type: "color", value: cur?.negColor ?? "#d00000", showFor: { key: "type", values: ["column", "stacked"] } },
    ], (v) => {
      const data = String(v.data).trim();
      const loc = parseRange(String(v.loc));
      if (!loc || !data) return;
      const type = String(v.type) as "line" | "column" | "stacked";
      const style = { type, color: String(v.color), negColor: type === "line" ? undefined : String(v.negColor) };
      const sheetName = data.includes("!") ? data.split("!")[0] : sheet.name;
      const dataBody = data.includes("!") ? data.split("!")[1]! : data;
      const dr = parseRange(dataBody);
      const locCells = (loc.r2 - loc.r1 + 1) * (loc.c2 - loc.c1 + 1);
      if (locCells > 1 && dr) {
        // Group: map each location cell to a data row (column of hosts) or column (row of hosts).
        const items: { host: { r: number; c: number }; dataRef: string }[] = [];
        if (loc.c1 === loc.c2) {
          for (let i = 0; loc.r1 + i <= loc.r2 && dr.r1 + i <= dr.r2; i++)
            items.push({ host: { r: loc.r1 + i, c: loc.c1 }, dataRef: `${sheetName}!${colToLetters(dr.c1)}${dr.r1 + i}:${colToLetters(dr.c2)}${dr.r1 + i}` });
        } else {
          for (let j = 0; loc.c1 + j <= loc.c2 && dr.c1 + j <= dr.c2; j++)
            items.push({ host: { r: loc.r1, c: loc.c1 + j }, dataRef: `${sheetName}!${colToLetters(dr.c1 + j)}${dr.r1}:${colToLetters(dr.c1 + j)}${dr.r2}` });
        }
        if (wb.kind === "ods") setOdsSparklineGroup(sheet, style, items);
        else setXlsxSparklineGroup(sheet, style, items);
      } else {
        const dataRef = data.includes("!") ? data : `${sheet.name}!${data}`;
        ctx.applySparkline(sheet, { r: loc.r1, c: loc.c1 }, { ...style, dataRef });
      }
      ctx.mark(); ctx.renderGrid();
    });
  };

  // Insert a new shape (no argument) or edit an existing one's geometry/fill/outline/text.
  const openShapeDialog = (existing?: SheetShape): void => {
    const sheet = sheetNow();
    const s = ctx.getSelRect();
    const geomOpts = [
      { value: "rect", label: t("shapeRect") },
      { value: "roundRect", label: t("shapeRoundRect") },
      { value: "ellipse", label: t("shapeEllipse") },
      { value: "triangle", label: t("shapeTriangle") },
      { value: "diamond", label: t("shapeDiamond") },
      { value: "parallelogram", label: t("shapeParallelogram") },
      { value: "pentagon", label: t("shapePentagon") },
      { value: "hexagon", label: t("shapeHexagon") },
      { value: "star", label: t("shapeStar") },
      { value: "rightArrow", label: t("shapeArrow") },
      { value: "line", label: t("shapeLine") },
    ];
    const filled = geomOpts.map((o) => o.value).filter((v) => v !== "line"); // fill / text apply to every non-line shape
    // When editing there is no geometry picker, so fill/text can't be gated on it: show them unless
    // the shape is a line. When creating, gate them live on the chosen geometry.
    const fillGate = existing ? undefined : ({ key: "geom", values: filled } as const);
    const showFillNow = !existing || existing.geom !== "line";
    dialog(existing ? t("shapeEdit") : t("shapeInsert"), [
      ...(existing ? [] : [{ key: "geom", label: t("shapeType"), type: "select" as const, value: "rect", options: geomOpts }]),
      ...(showFillNow ? [
        { key: "fill", label: t("shapeFill"), type: "color" as const, value: existing?.fill ?? "#4c8bf5", showFor: fillGate },
        { key: "noFill", label: t("shapeNoFill"), type: "checkbox" as const, value: existing ? !existing.fill : false, showFor: fillGate },
      ] : []),
      { key: "stroke", label: t("shapeOutline"), type: "color", value: existing?.stroke ?? "#1f3a5f" },
      { key: "strokeWidth", label: t("shapeOutlineWidth"), type: "text", value: String(existing?.strokeWidth ?? 1) },
      ...(showFillNow ? [{ key: "text", label: t("shapeText"), type: "text" as const, value: existing?.text ?? "", showFor: fillGate }] : []),
    ], (v) => {
      const geom = (existing?.geom ?? String(v.geom)) as ShapeGeom;
      const isLine = geom === "line";
      const fill = isLine || v.noFill ? undefined : String(v.fill);
      const stroke = String(v.stroke);
      const strokeWidth = Math.max(1, Number(v.strokeWidth) || 1);
      const text = isLine ? undefined : (String(v.text).trim() || undefined);
      if (existing) {
        existing.fill = fill; existing.stroke = stroke; existing.strokeWidth = strokeWidth; existing.text = text;
        existing.preset = undefined; // our geometry is authoritative once edited
        existing.dirty = true;
      } else {
        // Default location: the selection rect, or a ~3x4-cell box at the active cell.
        const twoD = s.r2 > s.r1 || s.c2 > s.c1;
        const anchor = twoD
          ? { fromCol: s.c1, fromRow: s.r1, fromColOff: 0, fromRowOff: 0, toCol: s.c2 + 1, toRow: s.r2 + 1, toColOff: 0, toRowOff: 0 }
          : { fromCol: s.c1, fromRow: s.r1, fromColOff: 0, fromRowOff: 0, toCol: s.c1 + 3, toRow: s.r1 + 4, toColOff: 0, toRowOff: 0 };
        (sheet.shapes ??= []).push({ geom, anchor, fill, stroke, strokeWidth, text, created: true, dirty: true });
      }
      ctx.mark(); ctx.refreshShapes();
    });
  };

  const openNoteDialog = (): void => {
    const s = ctx.getSelRect(); const r = s.r1, c = s.c1; const sheet = sheetNow();
    const cur = getCell(sheet, r, c)?.comments?.map((cm) => cm.text).join("\n") ?? "";
    dialog(t("noteEdit"), [{ key: "text", label: t("noteText"), type: "text", value: cur }], (v) => {
      const text = String(v.text).trim();
      if (wb.kind === "ods") setOdsComment(sheet, r, c, text || null);
      else setXlsxComment(wb, sheet, r, c, text || null);
      ctx.mark(); ctx.renderGrid();
    });
  };

  return { openLinkDialog, openDvDialog, openCfDialog, openSparkDialog, openShapeDialog, openNoteDialog };
}
