import { beforeAll, describe, expect, it } from "vitest";
import { strToU8 } from "fflate";
import { createSheetEditor } from "./editor";

// jsdom stubs (no layout engine): the editor falls back to a 600x1200 viewport.
beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
});

const frame = () => new Promise<void>((res) => setTimeout(res, 30));

const bigCsv = (rows: number): Uint8Array => {
  const lines: string[] = [];
  for (let i = 1; i <= rows; i++) lines.push(`r${i},${i}`);
  return strToU8(lines.join("\n") + "\n");
};

describe("virtualized grid", () => {
  it("renders a window, not the whole 50k-row sheet", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, bigCsv(50000), { formatHint: "csv" });
    const dataRows = host.querySelectorAll("td[data-rc]").length;
    expect(dataRows).toBeGreaterThan(0);
    expect(dataRows).toBeLessThan(2000); // window only, never 50k rows of DOM
    expect(host.querySelector('td[data-rc="1:1"]')).toBeTruthy();
    expect(host.querySelector('td[data-rc="25000:1"]')).toBeNull();
    expect(host.querySelector('td[data-rc="50000:1"]')).toBeNull();
    ed.destroy();
    host.remove();
  });

  it("scrolling re-renders the window at the new position", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, bigCsv(50000), { formatHint: "csv" });
    const grid = host.querySelector(".sheetedit-grid") as HTMLElement;
    grid.scrollTop = 24999 * 24; // ROW_H
    grid.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    const mid = host.querySelector('td[data-rc="25000:1"] input') as HTMLInputElement;
    expect(mid).toBeTruthy();
    expect(mid.value).toBe("r25000");
    expect(host.querySelector('td[data-rc="1:1"]')).toBeNull(); // scrolled out
    ed.destroy();
    host.remove();
  });

  it("virtualizes columns too: a 2000-column row renders a slice and scrolls", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const cols = Array.from({ length: 2000 }, (_v, i) => `c${i + 1}`);
    const ed = createSheetEditor(host, strToU8(cols.join(",") + "\n"), { formatHint: "csv" });
    expect(host.querySelector('td[data-rc="1:1"]')).toBeTruthy();
    expect(host.querySelector('td[data-rc="1:1000"]')).toBeNull();
    const heads = host.querySelectorAll("th.colhead").length;
    expect(heads).toBeLessThan(60);
    const grid = host.querySelector(".sheetedit-grid") as HTMLElement;
    grid.scrollLeft = 999 * 96; // COL_W
    grid.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    const far = host.querySelector('td[data-rc="1:1000"] input') as HTMLInputElement;
    expect(far).toBeTruthy();
    expect(far.value).toBe("c1000");
    ed.destroy();
    host.remove();
  });

  it("keeps an in-progress edit alive across a nearby window shift", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, bigCsv(200), { formatHint: "csv" });
    const grid = host.querySelector(".sheetedit-grid") as HTMLElement;
    const inp = host.querySelector('td[data-rc="2:2"] input') as HTMLInputElement;
    inp.dispatchEvent(new FocusEvent("focus"));
    inp.focus();
    inp.value = "pending edit";
    // Scroll a few rows: the edited cell stays within the pin range.
    grid.scrollTop = 24 * 10;
    grid.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    const again = host.querySelector('td[data-rc="2:2"] input') as HTMLInputElement;
    expect(again).toBeTruthy();
    expect(again.value).toBe("pending edit");
    ed.destroy();
    host.remove();
  });

  it("aggregates over huge ranges stay linear (the old path took ~30s)", () => {
    const lines: string[] = [];
    for (let i = 1; i <= 100000; i++) lines.push(`${i},${i % 10}`);
    lines.push(",=SUM(B1:B100000)");
    lines.push(",=AVERAGE(B1:B100000)");
    lines.push(",=COUNT(B1:B100000)");
    lines.push(",=MAX(B1:B100000)");
    const t = Date.now();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8(lines.join("\n") + "\n"), { formatHint: "csv" });
    expect(Date.now() - t).toBeLessThan(5000);
    const text = ed.getText()!;
    expect(text).toContain("=SUM(B1:B100000)");
    ed.destroy();
    host.remove();
  });

  it("aggregate overrides keep Excel semantics on small inputs", async () => {
    const { readCsv } = await import("../adapters/csv/read");
    const { recalc } = await import("./recalc");
    const { getCell } = await import("./model");
    const wb = readCsv(
      [
        "1,text,3", // B1 is text: ignored by range aggregates
        '=SUM(A1:C1),=AVERAGE(A1:C1),=COUNT(A1:C1)',
        '=COUNTA(A1:C1),=MIN(A1:C1),=MAX(A1:C1)',
        '"=SUM(1,""2"",TRUE)",=AVERAGE(D1:D1),x', // literals coerce; empty range divides by zero
      ].join("\n") + "\n",
    );
    recalc(wb);
    const s = wb.sheets[0]!;
    expect(getCell(s, 2, 1)?.value).toBe("4"); // SUM ignores the text
    expect(getCell(s, 2, 2)?.value).toBe("2"); // AVERAGE over the two numbers
    expect(getCell(s, 2, 3)?.value).toBe("2"); // COUNT numbers only
    expect(getCell(s, 3, 1)?.value).toBe("3"); // COUNTA counts the text too
    expect(getCell(s, 3, 2)?.value).toBe("1");
    expect(getCell(s, 3, 3)?.value).toBe("3");
    expect(getCell(s, 4, 1)?.value).toBe("4"); // 1 + "2" + TRUE
    expect(getCell(s, 4, 2)?.value).toBe("#DIV/0!");
  });

  it("the bottom of a 100k sheet renders the right rows", async () => {
    const lines = ["id,total"];
    for (let i = 1; i <= 100000; i++) lines.push(`${i},${i % 7}`);
    lines.push("total,=SUM(B2:B100001)");
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8(lines.join("\n") + "\n"), { formatHint: "csv" });
    const grid = host.querySelector(".sheetedit-grid") as HTMLElement;
    grid.scrollTop = 100008 * 24; // beyond the end: clamps to the bottom
    grid.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    const a = (r: number, c: number) => (host.querySelector(`td[data-rc="${r}:${c}"] input`) as HTMLInputElement | null)?.value;
    expect(a(100001, 1)).toBe("100000");
    expect(a(100002, 1)).toBe("total");
    expect(Number(a(100002, 2))).toBeGreaterThan(0); // the SUM computed on open
    ed.destroy();
    host.remove();
  });

  it("a big sheet still saves correctly after an edit deep in the file", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, bigCsv(30000), { formatHint: "csv" });
    const grid = host.querySelector(".sheetedit-grid") as HTMLElement;
    grid.scrollTop = 19999 * 24;
    grid.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    const inp = host.querySelector('td[data-rc="20000:2"] input') as HTMLInputElement;
    inp.dispatchEvent(new FocusEvent("focus"));
    inp.value = "999999";
    inp.dispatchEvent(new Event("input", { bubbles: true }));
    inp.dispatchEvent(new FocusEvent("blur"));
    const text = ed.getText()!;
    expect(text).toContain("r20000,999999\n");
    expect(text).toContain("r19999,19999\n"); // neighbours untouched
    expect(text.split("\n").length).toBe(30001);
    ed.destroy();
    host.remove();
  });
});
