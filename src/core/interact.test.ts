import { beforeAll, describe, expect, it } from "vitest";
import { strToU8 } from "fflate";
import { createSheetEditor } from "./editor";

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

/** jsdom has no DataTransfer: a minimal clipboard event stub that bubbles. */
const clipEvent = (type: "copy" | "paste", data: Record<string, string> = {}) => {
  const evt = new Event(type, { bubbles: true, cancelable: true }) as ClipboardEvent & {
    clipboardData: { setData(t: string, v: string): void; getData(t: string): string };
  };
  Object.defineProperty(evt, "clipboardData", {
    value: {
      setData: (t: string, v: string) => (data[t] = v),
      getData: (t: string) => data[t] ?? "",
    },
  });
  return { evt, data };
};

describe("scroll preservation across rebuilds", () => {
  it("a row op from the header menu keeps the scroll position", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, bigCsv(50000), { formatHint: "csv" });
    const grid = host.querySelector(".sheetedit-grid") as HTMLElement;
    grid.scrollTop = 24999 * 24;
    grid.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    const rn = [...host.querySelectorAll("th.rownum")].find((th) => th.textContent === "25000")!;
    rn.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    ([...document.querySelectorAll(".sheetedit-pop-item")][1] as HTMLButtonElement).click(); // insert below
    expect(grid.scrollTop).toBe(24999 * 24); // no jump to the top
    expect(host.querySelector('td[data-rc="25000:1"]')).toBeTruthy(); // window stayed in place
    ed.destroy();
    host.remove();
  });
});

describe("clipboard", () => {
  it("copies a drag-style selection without any focused input", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("a,b\nc,d\n"), { formatHint: "csv" });
    // Select A1:B2 the way a drag does: anchor + shift-extend, then blur.
    (host.querySelector('td[data-rc="1:1"] input') as HTMLInputElement).dispatchEvent(new FocusEvent("focus"));
    (host.querySelector('td[data-rc="2:2"] input') as HTMLInputElement).dispatchEvent(
      new MouseEvent("mousedown", { bubbles: true, cancelable: true, shiftKey: true }),
    );
    (document.activeElement as HTMLElement | null)?.blur?.();
    const { evt, data } = clipEvent("copy");
    document.body.dispatchEvent(evt);
    expect(data["text/plain"]).toBe("a\tb\nc\td");
    ed.destroy();
    host.remove();
  });

  it("single-cell copy copies the cell's raw content (formula text included)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("5,=A1*2\n"), { formatHint: "csv" });
    const inp = host.querySelector('td[data-rc="1:2"] input') as HTMLInputElement;
    inp.dispatchEvent(new FocusEvent("focus"));
    const { evt, data } = clipEvent("copy");
    inp.dispatchEvent(evt);
    expect(data["text/plain"]).toBe("=A1*2");
    ed.destroy();
    host.remove();
  });

  it("pastes a TSV block at the focused cell and keeps working after rebuild", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("a,b\nc,d\n"), { formatHint: "csv" });
    const inp = host.querySelector('td[data-rc="1:1"] input') as HTMLInputElement;
    inp.dispatchEvent(new FocusEvent("focus"));
    const { evt } = clipEvent("paste", { "text/plain": "x\ty\nz\tw" });
    inp.dispatchEvent(evt);
    expect(ed.getText()).toBe("x,y\nz,w\n");
    ed.destroy();
    host.remove();
  });
});

describe("row-number column width", () => {
  it("grows with the digit count of the last row", () => {
    const small = document.createElement("div");
    document.body.appendChild(small);
    const ed1 = createSheetEditor(small, strToU8("a\n"), { formatHint: "csv" });
    const w1 = parseFloat((small.querySelector("colgroup col") as HTMLElement).style.width);
    const big = document.createElement("div");
    document.body.appendChild(big);
    const ed2 = createSheetEditor(big, bigCsv(100000), { formatHint: "csv" });
    const w2 = parseFloat((big.querySelector("colgroup col") as HTMLElement).style.width);
    expect(w1).toBe(44);
    expect(w2).toBeGreaterThanOrEqual(44 + 16); // six digits need real room
    ed1.destroy();
    ed2.destroy();
    small.remove();
    big.remove();
  });
});

// Delete over a selection used to be handled on the cell's own input, and a drag selection
// blurs it: selecting a block and pressing Delete did nothing at all. Cut had no handler, so
// there was no way to move a block either.
describe("clearing and moving a selection", () => {
  const select = (host: HTMLElement, from: string, to: string) => {
    (host.querySelector(`td[data-rc="${from}"] input`) as HTMLInputElement).dispatchEvent(new FocusEvent("focus"));
    (host.querySelector(`td[data-rc="${to}"] input`) as HTMLInputElement).dispatchEvent(
      new MouseEvent("mousedown", { bubbles: true, cancelable: true, shiftKey: true }),
    );
    (document.activeElement as HTMLElement | null)?.blur?.();
  };
  const press = (key: string, mods: { metaKey?: boolean } = {}) =>
    document.body.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true, ...mods }));
  const valueAt = (host: HTMLElement, rc: string) => (host.querySelector(`td[data-rc="${rc}"] input`) as HTMLInputElement).value;

  it("Delete empties every cell of a block selected by dragging", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("a,b\nc,d\n"), { formatHint: "csv" });
    select(host, "1:1", "2:2");
    press("Delete");
    expect([valueAt(host, "1:1"), valueAt(host, "1:2"), valueAt(host, "2:1"), valueAt(host, "2:2")]).toEqual(["", "", "", ""]);
    ed.destroy();
    host.remove();
  });

  it("Delete empties the one selected cell too", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("a,b\n"), { formatHint: "csv" });
    const inp = host.querySelector('td[data-rc="1:2"] input') as HTMLInputElement;
    inp.dispatchEvent(new FocusEvent("focus"));
    inp.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true, cancelable: true }));
    expect(valueAt(host, "1:2")).toBe("");
    ed.destroy();
    host.remove();
  });

  it("leaves the caret alone while a cell is being edited", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("abc,b\n"), { formatHint: "csv" });
    const inp = host.querySelector('td[data-rc="1:1"] input') as HTMLInputElement;
    inp.dispatchEvent(new FocusEvent("focus"));
    inp.dispatchEvent(new KeyboardEvent("dblclick", { bubbles: true })); // now editing
    inp.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    inp.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true, cancelable: true }));
    expect(valueAt(host, "1:1")).toBe("abc");
    ed.destroy();
    host.remove();
  });

  it("cut takes the block's text and empties it, so paste moves it", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("a,b\nc,d\n"), { formatHint: "csv" });
    let written = "";
    const clip = navigator.clipboard;
    Object.defineProperty(navigator, "clipboard", {
      configurable: true,
      value: { writeText: (t: string) => ((written = t), Promise.resolve()) },
    });
    select(host, "1:1", "2:2");
    press("x", { metaKey: true });
    expect(written).toBe("a\tb\nc\td");
    expect([valueAt(host, "1:1"), valueAt(host, "2:2")]).toEqual(["", ""]);
    if (clip) Object.defineProperty(navigator, "clipboard", { configurable: true, value: clip });
    ed.destroy();
    host.remove();
  });
});

// Undo was handled on the cell's input as well, so after a drag, a move or a Delete, all of
// which leave the grid without focus, ctrl+Z reached nothing.
describe("undo without a focused cell", () => {
  it("takes back a Delete over a selection", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createSheetEditor(host, strToU8("a,b\nc,d\n"), { formatHint: "csv" });
    const inp = host.querySelector('td[data-rc="1:1"] input') as HTMLInputElement;
    inp.dispatchEvent(new FocusEvent("focus"));
    (host.querySelector('td[data-rc="2:2"] input') as HTMLInputElement).dispatchEvent(
      new MouseEvent("mousedown", { bubbles: true, cancelable: true, shiftKey: true }),
    );
    (document.activeElement as HTMLElement | null)?.blur?.();
    document.body.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true, cancelable: true }));
    expect((host.querySelector('td[data-rc="1:1"] input') as HTMLInputElement).value).toBe("");
    document.body.dispatchEvent(new KeyboardEvent("keydown", { key: "z", metaKey: true, bubbles: true, cancelable: true }));
    expect((host.querySelector('td[data-rc="1:1"] input') as HTMLInputElement).value).toBe("a");
    ed.destroy();
    host.remove();
  });
});
