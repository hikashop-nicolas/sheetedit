import { type Sheet, type Workbook } from "../model";
import { CHART_PALETTE, type ChartModel } from "../chart-model";
import { resolveNumbers, resolveLabels, seriesName } from "../chart-data";

// The chart layer: floats a Chart.js canvas over the grid for every chart on the active sheet,
// anchored to cells and kept glued while scrolling. The layer sits over the data area (below the
// column header, right of the row numbers) and translates its content by the scroll offset, so
// charts scroll with the cells and clip under the frozen header. Chart.js is lazy-imported the
// first time a chart is drawn, so it never weighs on chart-free workbooks.

export interface ChartGeom { xOfCol: (c: number) => number; yOfRow: (r: number) => number; colAt: (px: number) => number; rowAt: (px: number) => number; rnW: number; headerH: number }

export interface ChartLayerDeps {
  wrap: HTMLElement;
  gridScroll: HTMLElement;
  getSheet: () => Sheet | undefined;
  getWorkbook: () => Workbook;
  geom: () => ChartGeom;
  onSelect?: (chart: ChartModel | null) => void;
  /** After a move/resize: the model's anchor was updated and dirty set; the host marks + persists. */
  onEdit?: (chart: ChartModel) => void;
}

export type ChartCtor = (new (ctx: CanvasRenderingContext2D, cfg: unknown) => { update(): void; destroy(): void; data: unknown; options: unknown }) & { register(p: unknown): void };
let ChartJs: ChartCtor | null = null;
let loading: Promise<void> | null = null;
/** Lazy-load Chart.js and the data-labels plugin; resolves to the Chart constructor. */
export async function loadChartJs(): Promise<ChartCtor> {
  if (ChartJs) return ChartJs;
  loading ??= Promise.all([import("chart.js/auto"), import("chartjs-plugin-datalabels")]).then(([m, dl]) => {
    ChartJs = (m.default ?? (m as { Chart: ChartCtor }).Chart) as ChartCtor;
    ChartJs.register((dl.default ?? dl) as unknown); // registered globally; per-chart display is opt-in
  });
  await loading;
  return ChartJs!;
}
const ensureChartJs = async (): Promise<void> => { await loadChartJs(); };

const numbers = resolveNumbers;
const labels = resolveLabels;
const nameOf = seriesName;

/** Build a Chart.js config from the model + live data (shared by the overlay and the preview). */
export function chartConfig(model: ChartModel, wb: Workbook): unknown {
  return toConfig(model, wb);
}
const mapKind = (k: ChartModel["kind"]): string => (k === "column" || k === "bar" ? "bar" : k === "area" ? "line" : k);
function toConfig(model: ChartModel, wb: Workbook): unknown {
  const type = mapKind(model.kind);
  const cats = labels(wb, model.categories);
  const palette = (i: number, c?: string): string => c ?? CHART_PALETTE[i % CHART_PALETTE.length];
  const pieLike = model.kind === "pie" || model.kind === "doughnut";
  const hasSecondary = model.series.some((s) => s.secondaryAxis);
  const datasets = model.series.map((s, i) => {
    const base: Record<string, unknown> = { label: nameOf(wb, s.name), borderColor: palette(i, s.color), backgroundColor: palette(i, s.color) };
    if (model.kind === "scatter" || model.kind === "bubble") {
      const xs = numbers(wb, s.xValues); const ys = numbers(wb, s.values); const rs = s.sizes ? numbers(wb, s.sizes) : [];
      base.data = ys.map((y, j) => ({ x: xs[j] ?? j, y: y ?? 0, r: model.kind === "bubble" ? (rs[j] ?? 5) : undefined }));
    } else {
      base.data = numbers(wb, s.values).map((v) => v ?? 0);
      const effKind = s.type ?? model.kind;
      if (effKind === "area") base.fill = true;
      if (s.type && mapKind(s.type) !== type) base.type = mapKind(s.type); // combo: per-series type override
      if (s.secondaryAxis) base.yAxisID = "y1";
      // Pie / doughnut: colour each slice from the palette, not one colour for the whole series.
      if (pieLike) { base.backgroundColor = (base.data as number[]).map((_, j) => CHART_PALETTE[j % CHART_PALETTE.length]); base.borderColor = "#fff"; }
    }
    return base;
  });
  const stacked = model.stacked ? { stacked: true } : {};
  const axisTitle = (t?: string): object => (t ? { title: { display: true, text: t } } : {});
  const xt = axisTitle(model.axes?.x?.title);
  const yt = axisTitle(model.axes?.y?.title);
  const scales: Record<string, unknown> | undefined = model.kind === "bar" ? { x: { ...stacked, beginAtZero: true, ...yt }, y: { ...stacked, ...xt } }
    : model.kind === "column" || model.kind === "line" || model.kind === "area" ? { x: { ...stacked, ...xt }, y: { ...stacked, beginAtZero: model.kind !== "line", ...yt } }
    : undefined;
  if (scales && hasSecondary) scales.y1 = { position: "right", grid: { drawOnChartArea: false } };
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
        // The plugin is registered globally, so charts default it off; opt in per chart.
        datalabels: model.dataLabels ? { anchor: pieLike ? "center" : "end", align: pieLike ? "center" : "top", color: pieLike ? "#fff" : "#444", font: { size: 10 }, formatter: (v: unknown) => (typeof v === "object" && v ? (v as { y?: number }).y ?? "" : v) } : { display: false },
      },
      ...(scales ? { scales } : {}),
    },
  };
}

export function setupChartLayer(deps: ChartLayerDeps): { refresh(): void; update(): void; select(id: string | null): void; boxRect(id: string): DOMRect | null; teardown(): void } {
  const { wrap, gridScroll } = deps;
  const layer = document.createElement("div");
  layer.className = "sheetedit-chartlayer";
  const inner = document.createElement("div");
  inner.className = "sheetedit-chartlayer-inner";
  layer.appendChild(inner);
  wrap.appendChild(layer);

  const instances = new Map<string, { box: HTMLElement; canvas: HTMLCanvasElement; chart: InstanceType<ChartCtor> | null; kind?: string }>();
  let selectedId: string | null = null;

  const setSelected = (id: string | null): void => {
    selectedId = id;
    for (const [k, inst] of instances) inst.box.classList.toggle("sel", k === id);
    deps.onSelect?.(id ? (deps.getSheet()?.charts ?? []).find((c) => c.id === id) ?? null : null);
  };

  // Commit a dragged/resized pixel rect (content coords) back to the model's cell anchor.
  const rectToAnchor = (model: ChartModel, x: number, y: number, w: number, h: number): void => {
    const g = deps.geom();
    const set = (px: number, at: (p: number) => number, of: (i: number) => number): [number, number] => { const i = Math.max(1, at(px)); return [i, Math.max(0, px - of(i))]; };
    const [fc, fco] = set(x, g.colAt, g.xOfCol);
    const [fr, fro] = set(y, g.rowAt, g.yOfRow);
    const [tc, tco] = set(x + w, g.colAt, g.xOfCol);
    const [tr, tro] = set(y + h, g.rowAt, g.yOfRow);
    model.anchor = { fromCol: fc, fromRow: fr, fromColOff: fco, fromRowOff: fro, toCol: tc, toRow: tr, toColOff: tco, toRowOff: tro };
    model.dirty = true;
  };

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

  // Drag the box to move, or its corner handle to resize (mouse or touch); commit on release.
  function attachDrag(box: HTMLElement, handle: HTMLElement, model: ChartModel): void {
    const start = (e: PointerEvent, mode: "move" | "resize"): void => {
      e.preventDefault();
      e.stopPropagation();
      setSelected(model.id);
      const sx = e.clientX;
      const sy = e.clientY;
      const x0 = parseFloat(box.style.left) || 0;
      const y0 = parseFloat(box.style.top) || 0;
      const w0 = box.offsetWidth;
      const h0 = box.offsetHeight;
      const onMove = (ev: PointerEvent): void => {
        const dx = ev.clientX - sx;
        const dy = ev.clientY - sy;
        if (mode === "move") { box.style.left = `${Math.max(0, x0 + dx)}px`; box.style.top = `${Math.max(0, y0 + dy)}px`; }
        else { box.style.width = `${Math.max(80, w0 + dx)}px`; box.style.height = `${Math.max(60, h0 + dy)}px`; }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        rectToAnchor(model, parseFloat(box.style.left) || 0, parseFloat(box.style.top) || 0, box.offsetWidth, box.offsetHeight);
        deps.onEdit?.(model);
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    box.addEventListener("pointerdown", (e) => { if (e.target !== handle) start(e, "move"); });
    handle.addEventListener("pointerdown", (e) => start(e, "resize"));
  }

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
          const handle = document.createElement("div");
          handle.className = "sheetedit-chart-resize";
          box.appendChild(handle);
          attachDrag(box, handle, model);
          inner.appendChild(box);
          inst = { box, canvas, chart: null };
          instances.set(model.id, inst);
        }
        positionBox(inst.box, model);
        const cfg = toConfig(model, wb);
        // Chart.js can't change its type via update(); recreate when the kind changed.
        if (inst.chart && inst.kind !== model.kind) { inst.chart.destroy(); inst.chart = null; }
        if (inst.chart) { const c = inst.chart as { data: unknown; options: unknown; update(): void }; const nc = cfg as { data: unknown; options: unknown }; c.data = nc.data; c.options = nc.options; c.update(); }
        else inst.chart = new ChartJs(inst.canvas.getContext("2d")!, cfg);
        inst.kind = model.kind;
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
  // Tap/click on empty grid deselects.
  gridScroll.addEventListener("pointerdown", () => { if (selectedId) setSelected(null); });

  /** Screen rect of a chart's box (for anchoring an external edit toolbar). */
  const boxRect = (id: string): DOMRect | null => instances.get(id)?.box.getBoundingClientRect() ?? null;

  return {
    refresh,
    update,
    select: setSelected,
    boxRect,
    teardown(): void {
      gridScroll.removeEventListener("scroll", syncScroll);
      window.removeEventListener("resize", refresh);
      for (const inst of instances.values()) inst.chart?.destroy();
      layer.remove();
    },
  };
}
