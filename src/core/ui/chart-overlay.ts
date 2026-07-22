import { cellDisplay, getCell, parseA1Ref, type Sheet, type Workbook } from "../model";
import { CHART_PALETTE, type ChartModel, type ChartRef } from "../chart-model";

// The chart layer: floats a Chart.js canvas over the grid for every chart on the active sheet,
// anchored to cells and kept glued while scrolling. The layer sits over the data area (below the
// column header, right of the row numbers) and translates its content by the scroll offset, so
// charts scroll with the cells and clip under the frozen header. Chart.js is lazy-imported the
// first time a chart is drawn, so it never weighs on chart-free workbooks.

export interface ChartGeom { xOfCol: (c: number) => number; yOfRow: (r: number) => number; rnW: number; headerH: number }

export interface ChartLayerDeps {
  wrap: HTMLElement;
  gridScroll: HTMLElement;
  getSheet: () => Sheet | undefined;
  getWorkbook: () => Workbook;
  geom: () => ChartGeom;
  onSelect?: (chart: ChartModel | null) => void;
}

type ChartCtor = new (ctx: CanvasRenderingContext2D, cfg: unknown) => { update(): void; destroy(): void; data: unknown; options: unknown };
let ChartJs: ChartCtor | null = null;
let loading: Promise<void> | null = null;
async function ensureChartJs(): Promise<void> {
  if (ChartJs) return;
  loading ??= import("chart.js/auto").then((m) => { ChartJs = (m.default ?? (m as { Chart: ChartCtor }).Chart) as ChartCtor; });
  await loading;
}

/** Resolve a data ref to values from the live sheet, falling back to the cached points. */
function resolveRange(wb: Workbook, ref: string): { row: number; col: number; sheet: Sheet }[] | null {
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(ref);
  const sheetName = m ? (m[1] ?? m[2]) : undefined;
  const body = (m ? m[3] : ref).replace(/\$/g, "");
  const sheet = sheetName ? wb.sheets.find((s) => s.name === sheetName) : wb.sheets[0];
  if (!sheet) return null;
  const [a, b] = body.split(":");
  const p1 = parseA1Ref(a ?? "");
  const p2 = b ? parseA1Ref(b) : p1;
  if (!p1 || !p2) return null;
  const out: { row: number; col: number; sheet: Sheet }[] = [];
  for (let r = Math.min(p1.row, p2.row); r <= Math.max(p1.row, p2.row); r++)
    for (let c = Math.min(p1.col, p2.col); c <= Math.max(p1.col, p2.col); c++) out.push({ row: r, col: c, sheet });
  return out;
}
const numbers = (wb: Workbook, ref: ChartRef | undefined): (number | null)[] => {
  if (ref?.ref) { const cells = resolveRange(wb, ref.ref); if (cells) return cells.map(({ sheet, row, col }) => { const v = getCell(sheet, row, col)?.value ?? ""; const n = Number(v); return v !== "" && Number.isFinite(n) ? n : null; }); }
  return (ref?.cache ?? []).map((v) => (typeof v === "number" ? v : v != null && v !== "" && Number.isFinite(Number(v)) ? Number(v) : null));
};
const labels = (wb: Workbook, ref: ChartRef | undefined): string[] => {
  if (ref?.ref) { const cells = resolveRange(wb, ref.ref); if (cells) return cells.map(({ sheet, row, col }) => { const c = getCell(sheet, row, col); return c ? cellDisplay(c) : ""; }); }
  return (ref?.cache ?? []).map((v) => (v == null ? "" : String(v)));
};
const nameOf = (wb: Workbook, name: ChartSeriesName): string | undefined => {
  if (typeof name === "string") return name;
  if (!name) return undefined;
  const l = labels(wb, name);
  return l[0] || undefined;
};
type ChartSeriesName = string | ChartRef | undefined;

/** Build a Chart.js config from the model + live data. */
function toConfig(model: ChartModel, wb: Workbook): unknown {
  const type = model.kind === "column" || model.kind === "bar" ? "bar" : model.kind === "area" ? "line" : model.kind;
  const cats = labels(wb, model.categories);
  const palette = (i: number, c?: string): string => c ?? CHART_PALETTE[i % CHART_PALETTE.length];
  const datasets = model.series.map((s, i) => {
    const base: Record<string, unknown> = { label: nameOf(wb, s.name), borderColor: palette(i, s.color), backgroundColor: palette(i, s.color) };
    if (model.kind === "scatter" || model.kind === "bubble") {
      const xs = numbers(wb, s.xValues); const ys = numbers(wb, s.values); const rs = s.sizes ? numbers(wb, s.sizes) : [];
      base.data = ys.map((y, j) => ({ x: xs[j] ?? j, y: y ?? 0, r: model.kind === "bubble" ? (rs[j] ?? 5) : undefined }));
    } else {
      base.data = numbers(wb, s.values).map((v) => v ?? 0);
      if (model.kind === "area") base.fill = true;
    }
    return base;
  });
  const stacked = model.stacked ? { stacked: true } : {};
  const scales = model.kind === "bar" ? { x: { ...stacked, beginAtZero: true }, y: { ...stacked } }
    : model.kind === "column" || model.kind === "line" || model.kind === "area" ? { x: { ...stacked }, y: { ...stacked, beginAtZero: model.kind !== "line" } }
    : undefined;
  return {
    type,
    data: { labels: cats.length ? cats : undefined, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      indexAxis: model.kind === "bar" ? "y" : "x",
      plugins: {
        legend: { display: model.legend?.show ?? true, position: model.legend?.pos ?? "top" },
        title: { display: !!model.title, text: model.title },
      },
      ...(scales ? { scales } : {}),
    },
  };
}

export function setupChartLayer(deps: ChartLayerDeps): { refresh(): void; update(): void; teardown(): void } {
  const { wrap, gridScroll } = deps;
  const layer = document.createElement("div");
  layer.className = "sheetedit-chartlayer";
  const inner = document.createElement("div");
  inner.className = "sheetedit-chartlayer-inner";
  layer.appendChild(inner);
  wrap.appendChild(layer);

  const instances = new Map<string, { box: HTMLElement; canvas: HTMLCanvasElement; chart: InstanceType<ChartCtor> | null }>();

  const positionLayer = (): void => {
    const g = deps.geom();
    const gr = gridScroll.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    layer.style.left = `${gr.left - wr.left + g.rnW}px`;
    layer.style.top = `${gr.top - wr.top + g.headerH}px`;
    layer.style.width = `${Math.max(0, gr.width - g.rnW)}px`;
    layer.style.height = `${Math.max(0, gr.height - g.headerH)}px`;
  };
  const syncScroll = (): void => { inner.style.transform = `translate(${-gridScroll.scrollLeft}px, ${-gridScroll.scrollTop}px)`; };

  const positionBox = (box: HTMLElement, model: ChartModel): void => {
    const g = deps.geom();
    const a = model.anchor;
    const x = g.xOfCol(a.fromCol) + a.fromColOff;
    const y = g.yOfRow(a.fromRow) + a.fromRowOff;
    const x2 = g.xOfCol(a.toCol) + a.toColOff;
    const y2 = g.yOfRow(a.toRow) + a.toRowOff;
    box.style.left = `${x}px`;
    box.style.top = `${y}px`;
    box.style.width = `${Math.max(60, x2 - x)}px`;
    box.style.height = `${Math.max(40, y2 - y)}px`;
  };

  let drawSeq = 0;
  const refresh = (): void => {
    const sheet = deps.getSheet();
    const charts = sheet?.charts ?? [];
    positionLayer();
    syncScroll();
    // Drop instances for charts no longer present.
    for (const [id, inst] of instances) if (!charts.some((c) => c.id === id)) { inst.chart?.destroy(); inst.box.remove(); instances.delete(id); }
    if (!charts.length) return;
    const wb = deps.getWorkbook();
    const seq = ++drawSeq;
    void ensureChartJs().then(() => {
      if (seq !== drawSeq || !ChartJs) return; // a newer refresh superseded this one
      for (const model of charts) {
        let inst = instances.get(model.id);
        if (!inst) {
          const box = document.createElement("div");
          box.className = "sheetedit-chartbox";
          const canvas = document.createElement("canvas");
          box.appendChild(canvas);
          box.addEventListener("mousedown", (e) => { e.stopPropagation(); deps.onSelect?.(model); });
          inner.appendChild(box);
          inst = { box, canvas, chart: null };
          instances.set(model.id, inst);
        }
        positionBox(inst.box, model);
        const cfg = toConfig(model, wb);
        if (inst.chart) { const c = inst.chart as { data: unknown; options: unknown; update(): void }; const nc = cfg as { data: unknown; options: unknown }; c.data = nc.data; c.options = nc.options; c.update(); }
        else inst.chart = new ChartJs(inst.canvas.getContext("2d")!, cfg);
      }
    });
  };

  // Data-only refresh (a cell edit changed a value a chart reads): re-resolve and update, no reposition.
  const update = (): void => {
    const wb = deps.getWorkbook();
    for (const model of deps.getSheet()?.charts ?? []) {
      const inst = instances.get(model.id);
      if (!inst?.chart) continue;
      const cfg = toConfig(model, wb) as { data: unknown; options: unknown };
      const c = inst.chart as { data: unknown; options: unknown; update(): void };
      c.data = cfg.data; c.options = cfg.options; c.update();
    }
  };

  gridScroll.addEventListener("scroll", syncScroll, { passive: true });
  window.addEventListener("resize", refresh);

  return {
    refresh,
    update,
    teardown(): void {
      gridScroll.removeEventListener("scroll", syncScroll);
      window.removeEventListener("resize", refresh);
      for (const inst of instances.values()) inst.chart?.destroy();
      layer.remove();
    },
  };
}
