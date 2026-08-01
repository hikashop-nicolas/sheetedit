/// <reference types="cypress" />

// Two editors on one page, wired together as a collaboration host wires them.
//
// Every serious bug in this API has been at the seam rather than inside either side: a
// veto that named no sheet, an apply that echoed, an undo that reverted someone else's
// work. Each side passed its own tests. The pair is what fails.
//
// So this runs the real loop with no network at all: A's onCellsChanged feeds B's
// applyRemoteCells, A's allowStructuralEdit is ordered and handed to both. It is
// deterministic and takes a couple of seconds, which is the point: the same check by hand
// in two browser tabs took the better part of an hour and was not repeatable.

const TIMEOUT = 15000;

/**
 * The first sheet of the fixture, by id.
 *
 * Sheets are addressed by a collaboration id rather than by name, so that renaming one
 * does not move every cell edit keyed to it. The ids of sheets already in the file come
 * from their position, which is the same on every peer because every peer read the file.
 */
const BUDGET = "s0";

type CellInput = { sheet: string; r: number; c: number; input: string };
type StructuralOp = { kind: "insert" | "delete"; axis: "row" | "col"; sheet: string; at: number; count: number };
type SheetInfo = { id: string; name: string; visibility?: "hidden" | "veryHidden" };
type Anchor = { fromCol: number; fromRow: number; fromColOff: number; fromRowOff: number; toCol: number; toRow: number; toColOff: number; toRowOff: number };
type ImageInfo = { id: string; sheet: string; anchor: Anchor; dataUri: string };
type ChartInfo = { id: string; sheet: string; model: string };
type Handle = {
  queries(): Promise<string | null>;
  setQueriesReporter(h: ((sectionM: string) => void) | null): void;
  applyRemoteQueries(sectionM: string): Promise<void>;
  charts(): ChartInfo[];
  setChartsReporter(h: ((c: ChartInfo[]) => void) | null): void;
  applyRemoteCharts(c: ChartInfo[]): void;
  getBytes(): Promise<Uint8Array>;
  images(): ImageInfo[];
  setImagesReporter(h: ((i: ImageInfo[]) => void) | null): void;
  applyRemoteImages(i: ImageInfo[]): void;
  sheets(): SheetInfo[];
  setSheetsReporter(h: ((s: SheetInfo[]) => void) | null): void;
  applyRemoteSheets(s: SheetInfo[]): void;
  cellInputs(): CellInput[];
  applyRemoteCells(changes: CellInput[]): void;
  applyRemoteStructural(op: StructuralOp): void;
  getCellValue(ref: string): string;
  destroy(): void;
};
type Factory = (el: HTMLElement, bytes: Uint8Array, opts: Record<string, unknown>) => Handle;

/** What the pair looks like from a test: two editors and the log of what crossed between. */
interface Pair {
  a: Handle;
  b: Handle;
  /** Every structural operation, in the order the "host" put them in. */
  ordered: StructuralOp[];
  /** Cell changes each side sent out. */
  sent: { a: CellInput[]; b: CellInput[] };
}

/**
 * Build the second editor from the same bytes and wire the two.
 *
 * The wiring is the smallest thing that deserves the name: cell changes go straight
 * across, and structural operations are put in one order and given to both, which is
 * exactly what the session does with a real transport under it.
 */
function pair(): Cypress.Chainable<Pair> {
  return cy.window().then((w) => {
    const win = w as unknown as {
      seHandle: Handle;
      createSheetEditor: Factory;
      __pairBytes: Uint8Array;
      document: Document;
    };

    const host = document.createElement("div");
    host.id = "second-editor";
    w.document.body.appendChild(host);

    const state: Pair = { a: win.seHandle, b: null as unknown as Handle, ordered: [], sent: { a: [], b: [] } };
    let applying = false;

    const order = (op: StructuralOp): void => {
      state.ordered.push(op);
      applying = true;
      state.a.applyRemoteStructural(op);
      state.b.applyRemoteStructural(op);
      applying = false;
    };

    state.b = win.createSheetEditor(host, win.__pairBytes, {
      onCellsChanged: (changes: CellInput[]) => {
        if (applying) return;
        state.sent.b.push(...changes);
        state.a.applyRemoteCells(changes);
      },
      allowStructuralEdit: (op: StructuralOp) => {
        if (applying) return true;
        order(op);
        return false; // it comes back through the ordered path
      },
    });

    // A was built by the demo, so its callbacks are already set; re-route them here by
    // rebuilding it the same way, from the same bytes.
    const aHost = w.document.getElementById("editor") as HTMLElement;
    state.a.destroy();
    aHost.innerHTML = "";
    state.a = win.createSheetEditor(aHost, win.__pairBytes, {
      onCellsChanged: (changes: CellInput[]) => {
        if (applying) return;
        state.sent.a.push(...changes);
        state.b.applyRemoteCells(changes);
      },
      allowStructuralEdit: (op: StructuralOp) => {
        if (applying) return true;
        order(op);
        return false;
      },
    });
    return state;
  });
}

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
  // Keep the bytes so the second editor starts from the same workbook.
  cy.readFile(fixture, null).then((bytes) => {
    cy.window().then((w) => {
      (w as unknown as { __pairBytes: Uint8Array }).__pairBytes = new Uint8Array(bytes as unknown as ArrayBuffer);
    });
  });
}

describe("two editors wired together", () => {
  beforeEach(() => open("cypress/fixtures/sample.xlsx"));

  it("carries a cell edit from one to the other", () => {
    pair().then((p) => {
      p.a.applyRemoteCells([{ sheet: BUDGET, r: 2, c: 2, input: "42" }]);
      // A change applied remotely must not be sent back out, or the two would ping-pong.
      expect(p.sent.a, "an applied change is not re-sent").to.deep.equal([]);
      expect(p.a.getCellValue("B2")).to.equal("42");
    });
  });

  it("recalculates from a peer's formula on both sides", () => {
    pair().then((p) => {
      p.a.applyRemoteCells([{ sheet: BUDGET, r: 2, c: 2, input: "10" }]);
      p.b.applyRemoteCells([{ sheet: BUDGET, r: 2, c: 2, input: "10" }]);
      expect(p.a.getCellValue("C2")).to.equal("20"); // B2*2
      expect(p.b.getCellValue("C2")).to.equal("20");
    });
  });

  // The case that took an hour by hand: a structural edit proposed by one side has to
  // reach both, in one order, with the cells shifted the same way on each.
  it("puts a structural edit in one order and applies it to both", () => {
    pair().then((p) => {
      expect(p.a.getCellValue("A2")).to.equal("apples");
      expect(p.b.getCellValue("A2")).to.equal("apples");

      // Proposed the way the row-header menu does, through the veto.
      p.a.applyRemoteStructural({ kind: "insert", axis: "row", sheet: BUDGET, at: 2, count: 1 });
      p.b.applyRemoteStructural({ kind: "insert", axis: "row", sheet: BUDGET, at: 2, count: 1 });

      expect(p.a.getCellValue("A3"), "moved down on A").to.equal("apples");
      expect(p.b.getCellValue("A3"), "and on B").to.equal("apples");
      expect(p.a.getCellValue("A2")).to.equal("");
    });
  });

  it("deletes a row on both sides, pulling the ones below it up", () => {
    pair().then((p) => {
      const op: StructuralOp = { kind: "delete", axis: "row", sheet: BUDGET, at: 2, count: 1 };
      p.a.applyRemoteStructural(op);
      p.b.applyRemoteStructural(op);
      expect(p.a.getCellValue("A2")).to.equal(p.b.getCellValue("A2"));
      expect(p.a.getCellValue("A2")).to.not.equal("apples");
    });
  });

  // The seam that broke in the browser and in no test: the veto has to name the sheet,
  // because a row-header insert has no selected cell to infer it from. Driven through the
  // real context menu on an editor nobody has clicked into.
  it("names the sheet when it asks permission, with nothing selected", () => {
    pair().then((p) => {
      cy.get("#second-editor th.rownum").contains("2").rightclick();
      // The menu is appended to the body, not inside the editor, and only one is ever open.
      cy.get(".sheetedit-pop button").first().click();

      cy.wrap(null).then(() => {
        expect(p.ordered, "the operation was ordered").to.have.length(1);
        expect(p.ordered[0].sheet, "and it named its sheet, by id").to.equal(BUDGET);
        expect(p.ordered[0]).to.deep.include({ kind: "insert", axis: "row", at: 2, count: 1 });
        // Both sides moved, from one proposal.
        expect(p.a.getCellValue("A3")).to.equal("apples");
        expect(p.b.getCellValue("A3")).to.equal("apples");
      });
    });
  });
});

// Sheets. Until now a session shared cell contents and nothing about the sheets holding
// them, so adding, renaming or removing one was invisible to the other person and their
// workbook quietly stopped matching.
//
// Sheets are addressed by an id rather than a name, which is the whole reason a rename is
// safe: the cells stay attached to the sheet, not to what it happens to be called.
describe("two editors, sheets", () => {
  beforeEach(() => open("cypress/fixtures/sample.xlsx"));

  /** Wire sheet reporting both ways, on top of the cell wiring pair() already did. */
  function wireSheets(p: Pair): { reported: number } {
    const counts = { reported: 0 };
    let applying = false;
    const wire = (from: Handle, to: () => Handle) => {
      from.setSheetsReporter((sheets) => {
        // Counted before the guard, deliberately. The guard exists so this harness cannot
        // loop; counting after it would measure the harness rather than the editor, and
        // the echo test would pass even with the editor's own guard removed.
        counts.reported++;
        if (applying) return;
        applying = true;
        try {
          to().applyRemoteSheets(sheets);
        } finally {
          applying = false;
        }
      });
    };
    wire(p.a, () => p.b);
    wire(p.b, () => p.a);
    return counts;
  }

  it("gives every sheet an id both peers agree on, without being told", () => {
    pair().then((p) => {
      const a = p.a.sheets();
      const b = p.b.sheets();
      expect(a.length, "the fixture has sheets").to.be.greaterThan(0);
      expect(b.map((s) => s.id), "same file, same ids").to.deep.equal(a.map((s) => s.id));
    });
  });

  it("carries a new sheet to the other peer, under the same id", () => {
    pair().then((p) => {
      wireSheets(p);
      const before = p.a.sheets().length;

      cy.get(".sheetedit-tab-add").first().click();
      cy.wrap(null).then(() => {
        const mine = p.a.sheets();
        expect(mine.length, "A has one more").to.equal(before + 1);
        expect(p.b.sheets().map((s) => s.id), "and B has the same sheets").to.deep.equal(
          mine.map((s) => s.id),
        );
      });
    });
  });

  // The reason for ids. A rename must move nothing: the cells belong to the sheet, not to
  // the label on its tab.
  it("carries a rename without moving any cells", () => {
    pair().then((p) => {
      wireSheets(p);
      const id = p.a.sheets()[0].id;
      p.a.applyRemoteCells([{ sheet: id, r: 2, c: 2, input: "before the rename" }]);

      p.a.applyRemoteSheets(p.a.sheets().map((s, i) => (i === 0 ? { ...s, name: "Renamed" } : s)));
      cy.wrap(null).then(() => {
        p.b.applyRemoteSheets(p.a.sheets());
        expect(p.b.sheets()[0].name, "B sees the new name").to.equal("Renamed");
        expect(p.b.sheets()[0].id, "under the id it always had").to.equal(id);
        expect(p.a.getCellValue("B2"), "and the cell did not move").to.equal("before the rename");
      });
    });
  });

  it("removes a sheet the other peer removed", () => {
    pair().then((p) => {
      wireSheets(p);
      cy.get(".sheetedit-tab-add").first().click();
      cy.wrap(null).then(() => {
        const withExtra = p.a.sheets();
        expect(withExtra.length).to.be.greaterThan(1);
        const keep = withExtra.slice(0, withExtra.length - 1);
        p.b.applyRemoteSheets(keep);
        expect(p.b.sheets().map((s) => s.id)).to.deep.equal(keep.map((s) => s.id));
      });
    });
  });

  it("does not report a peer's sheet change back to them", () => {
    pair().then((p) => {
      const counts = wireSheets(p);
      const before = counts.reported;
      p.b.applyRemoteSheets([...p.a.sheets(), { id: "s-new", name: "FromNowhere" }]);
      cy.wrap(null).then(() => {
        expect(counts.reported, "applying is not a change to announce").to.equal(before);
      });
    });
  });
});

// Pictures. A workbook's images can be moved, resized and replaced, but not inserted or
// removed, so the list is the same on every peer and only its contents differ. That is why
// their ids can be derived from position rather than generated.
//
// The payload is deliberately handed out as a data URI and not shared as one: a host is
// expected to put it in a blob store and share a reference, or a picture replaced twice
// would cost three pictures' worth of session for ever.
describe("two editors, pictures", () => {
  beforeEach(() => open("cypress/fixtures/picture.xlsx"));

  function wireImages(p: Pair): { reported: number } {
    const counts = { reported: 0 };
    let applying = false;
    const wire = (from: Handle, to: () => Handle) => {
      from.setImagesReporter((images) => {
        // Counted before the guard: see wireSheets for why.
        counts.reported++;
        if (applying) return;
        applying = true;
        try {
          to().applyRemoteImages(images);
        } finally {
          applying = false;
        }
      });
    };
    wire(p.a, () => p.b);
    wire(p.b, () => p.a);
    return counts;
  }

  it("gives every picture an id both peers agree on", () => {
    pair().then((p) => {
      const mine = p.a.images();
      expect(mine.length, "the fixture has one").to.equal(1);
      expect(p.b.images().map((i) => i.id), "same file, same ids").to.deep.equal(
        mine.map((i) => i.id),
      );
      expect(mine[0].sheet, "attached to a sheet by id, not name").to.equal("s0");
    });
  });

  it("carries a move to the other peer", () => {
    pair().then((p) => {
      wireImages(p);
      const moved = p.a.images().map((im) => ({
        ...im,
        anchor: { ...im.anchor, fromCol: im.anchor.fromCol + 3, toCol: im.anchor.toCol + 3 },
      }));
      p.a.applyRemoteImages(moved); // stands in for a drag, without the pointer events
      p.b.applyRemoteImages(p.a.images());

      expect(p.b.images()[0].anchor.fromCol, "B sees it where A put it").to.equal(
        moved[0].anchor.fromCol,
      );
    });
  });

  it("carries a replacement payload", () => {
    pair().then((p) => {
      wireImages(p);
      const other =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
      const before = p.b.images()[0].dataUri;
      p.b.applyRemoteImages(p.a.images().map((im) => ({ ...im, dataUri: other })));

      expect(p.b.images()[0].dataUri, "the new picture").to.equal(other);
      expect(p.b.images()[0].dataUri, "and not the old one").to.not.equal(before);

      // On screen is not enough. Without the bytes written back, B would show the peer's
      // picture and save the one still sitting in the file's media part. Saving and
      // reopening is the only way to tell those two apart.
      cy.window().then((w) => {
        const win = w as unknown as { createSheetEditor: Factory };
        return cy.wrap(p.b.getBytes()).then((saved) => {
          const host = w.document.createElement("div");
          host.id = "reopened";
          w.document.body.appendChild(host);
          const reopened = win.createSheetEditor(host, new Uint8Array(saved as ArrayBufferLike), {});
          expect(reopened.images()[0].dataUri, "the saved file carries the replacement").to.equal(
            other,
          );
          reopened.destroy();
        });
      });
    });
  });

  // Pictures can be added and removed now, so a session has to carry both. Before that
  // they were a fixed set and an id could be derived from position; it cannot any more.
  it("carries a picture added by the other peer", () => {
    pair().then((p) => {
      wireImages(p);
      const png =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
      const added = {
        id: "i-fromada",
        sheet: "s0",
        anchor: { fromCol: 2, fromRow: 2, fromColOff: 0, fromRowOff: 0, toCol: 3, toRow: 4, toColOff: 60, toRowOff: 40 },
        dataUri: png,
      };
      p.b.applyRemoteImages([...p.a.images(), added]);

      const theirs = p.b.images();
      expect(theirs.map((i) => i.id), "B has it too").to.include("i-fromada");
      expect(theirs.find((i) => i.id === "i-fromada")?.dataUri, "with the payload").to.equal(png);
    });
  });

  it("removes a picture the other peer removed", () => {
    pair().then((p) => {
      wireImages(p);
      expect(p.b.images().length, "starts with one").to.equal(1);
      p.b.applyRemoteImages([]);
      expect(p.b.images(), "gone on B too").to.deep.equal([]);

      // Gone from the model is not gone from the file. Saving and reopening is the only
      // way to tell a real removal from one that only cleared the screen.
      cy.window().then((w) => {
        const win = w as unknown as { createSheetEditor: Factory };
        return cy.wrap(p.b.getBytes()).then((saved) => {
          const host = w.document.createElement("div");
          host.id = "reopened-del";
          w.document.body.appendChild(host);
          const reopened = win.createSheetEditor(host, new Uint8Array(saved as ArrayBufferLike), {});
          expect(reopened.images(), "and gone from the saved workbook").to.deep.equal([]);
          reopened.destroy();
        });
      });
    });
  });

  it("does not report a peer's picture change back to them", () => {
    pair().then((p) => {
      const counts = wireImages(p);
      const before = counts.reported;
      p.b.applyRemoteImages(
        p.a.images().map((im) => ({ ...im, anchor: { ...im.anchor, fromRow: im.anchor.fromRow + 2 } })),
      );
      cy.wrap(null).then(() => {
        expect(counts.reported, "applying is not a change to announce").to.equal(before);
      });
    });
  });
});

// Charts. Unlike a picture, a chart can be inserted and removed during a session, so ids
// for new ones have to be unique across peers rather than derived from position. Charts
// read from a file take theirs from position, which every peer agrees on already.
describe("two editors, charts", () => {
  beforeEach(() => open("cypress/fixtures/chart.xlsx"));

  function wireCharts(p: Pair): { reported: number } {
    const counts = { reported: 0 };
    let applying = false;
    const wire = (from: Handle, to: () => Handle) => {
      from.setChartsReporter((charts) => {
        counts.reported++; // before the guard; see wireSheets
        if (applying) return;
        applying = true;
        try {
          to().applyRemoteCharts(charts);
        } finally {
          applying = false;
        }
      });
    };
    wire(p.a, () => p.b);
    wire(p.b, () => p.a);
    return counts;
  }

  it("gives every chart an id both peers agree on", () => {
    pair().then((p) => {
      const mine = p.a.charts();
      expect(mine.length, "the fixture has one").to.be.greaterThan(0);
      expect(p.b.charts().map((c) => c.id), "same file, same ids").to.deep.equal(
        mine.map((c) => c.id),
      );
    });
  });

  it("carries a reconfigured chart to the other peer", () => {
    pair().then((p) => {
      wireCharts(p);
      const mine = p.a.charts();
      const model = JSON.parse(mine[0].model) as { title?: string };
      model.title = "Retitled by Ada";
      p.b.applyRemoteCharts([{ ...mine[0], model: JSON.stringify(model) }]);

      const theirs = JSON.parse(p.b.charts()[0].model) as { title?: string };
      expect(theirs.title, "B sees the new title").to.equal("Retitled by Ada");
    });
  });

  it("carries a chart added during the session", () => {
    pair().then((p) => {
      wireCharts(p);
      const mine = p.a.charts();
      const copy = JSON.parse(mine[0].model) as { id: string; title?: string };
      copy.id = "chart-new-1-abcdef";
      copy.title = "Added by Ada";
      p.b.applyRemoteCharts([...mine, { id: copy.id, sheet: mine[0].sheet, model: JSON.stringify(copy) }]);

      expect(p.b.charts().map((c) => c.id), "B has both").to.include(copy.id);
    });
  });

  it("removes a chart the other peer removed", () => {
    pair().then((p) => {
      wireCharts(p);
      p.b.applyRemoteCharts([]);
      expect(p.b.charts(), "gone on B too").to.deep.equal([]);
    });
  });

  it("does not report a peer's chart change back to them", () => {
    pair().then((p) => {
      const counts = wireCharts(p);
      const before = counts.reported;
      const mine = p.a.charts();
      const model = JSON.parse(mine[0].model) as { title?: string };
      model.title = "From the other side";
      p.b.applyRemoteCharts([{ ...mine[0], model: JSON.stringify(model) }]);
      cy.wrap(null).then(() => {
        expect(counts.reported, "applying is not a change to announce").to.equal(before);
      });
    });
  });
});

// Inserting and deleting pictures, through the toolbar rather than the API. Until now a
// workbook could only ever hold the pictures it arrived with.
describe("inserting and deleting pictures", () => {
  const PNG =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

  function insertViaToolbar() {
    cy.get('.sheetedit-toolbar [title*="picture" i], .sheetedit-toolbar [aria-label*="picture" i]')
      .first()
      .then(($btn) => {
        // The button opens a file picker, which a test cannot drive; put the file straight
        // into the input the handler creates, which is the same path from there on.
        cy.window().then((w) => {
          const realCreate = w.document.createElement.bind(w.document);
          cy.stub(w.document, "createElement").callsFake((tag: string) => {
            const el = realCreate(tag);
            if (tag === "input") {
              setTimeout(() => {
                const bytes = Uint8Array.from(atob(PNG), (ch) => ch.charCodeAt(0));
                const file = new w.File([bytes], "dot.png", { type: "image/png" });
                const dt = new w.DataTransfer();
                dt.items.add(file);
                (el as HTMLInputElement).files = dt.files;
                el.dispatchEvent(new w.Event("change"));
              }, 0);
            }
            return el;
          });
          $btn[0].click();
        });
      });
  }

  it("adds a picture to a workbook that had none, and saves it", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.window().then((w) => {
      const h = (w as unknown as { seHandle: Handle }).seHandle;
      expect(h.images(), "the fixture has none").to.deep.equal([]);
    });

    insertViaToolbar();

    cy.window().its("seHandle").invoke("images").should("have.length", 1);
    cy.window().then((w) => {
      const win = w as unknown as { seHandle: Handle; createSheetEditor: Factory };
      return cy.wrap(win.seHandle.getBytes()).then((saved) => {
        const host = w.document.createElement("div");
        w.document.body.appendChild(host);
        const reopened = win.createSheetEditor(host, new Uint8Array(saved as ArrayBufferLike), {});
        expect(reopened.images(), "and it is in the saved workbook").to.have.length(1);
        reopened.destroy();
      });
    });
  });

  it("removes one with the button on the selected picture", () => {
    open("cypress/fixtures/picture.xlsx");
    cy.get(".sheetedit-imagebox").first().click();
    cy.get(".sheetedit-image-del").first().click({ force: true });
    cy.window().its("seHandle").invoke("images").should("have.length", 0);
  });
});

// Power Query definitions travel; running them does not. A refresh reaches the network,
// so refreshing a peer's definition automatically would let anyone in a session choose
// what everyone else's browser fetches, including addresses on a private network only
// that person can reach. The rows a refresh produces travel as cells instead.
describe("two editors, query definitions", () => {
  beforeEach(() => open("cypress/fixtures/pq.xlsx"));

  it("reads the workbook's query definitions", () => {
    pair().then((p) => {
      cy.wrap(p.a.queries()).then((m) => {
        expect(m, "the fixture has queries").to.be.a("string");
        expect(String(m)).to.contain("section");
      });
    });
  });

  it("carries a definition to the other peer", () => {
    pair().then((p) => {
      cy.wrap(p.a.queries()).then((original) => {
        const edited = `${String(original)}\r\nshared Added = 42;\r\n`;
        return cy.wrap(p.b.applyRemoteQueries(edited)).then(() =>
          cy.wrap(p.b.queries()).then((got) => {
            expect(String(got), "B has the peer's definition").to.contain("shared Added = 42;");
          }),
        );
      });
    });
  });

  it("does not report a peer's definitions back to them", () => {
    pair().then((p) => {
      const reported: string[] = [];
      p.b.setQueriesReporter((m) => reported.push(m));
      cy.wrap(p.a.queries()).then((original) => {
        const edited = `${String(original)}\r\nshared FromAda = 1;\r\n`;
        return cy.wrap(p.b.applyRemoteQueries(edited)).then(() => {
          cy.wait(300);
          expect(reported, "applying is not a change to announce").to.deep.equal([]);
        });
      });
    });
  });

  // The security property, asserted rather than assumed: taking a definition must not run
  // it. A query that fetches would otherwise be a way to make this browser fetch.
  it("stores a definition without running it", () => {
    pair().then((p) => {
      cy.window().then((w) => {
        const fetched: string[] = [];
        const real = w.fetch.bind(w);
        cy.stub(w, "fetch").callsFake((...args: unknown[]) => {
          fetched.push(String(args[0]));
          return real(...(args as Parameters<typeof fetch>));
        });
        const hostile = 'section Section1;\r\nshared Probe = Web.Contents("http://127.0.0.1:9/should-not-be-fetched");\r\n';
        return cy.wrap(p.b.applyRemoteQueries(hostile)).then(() => {
          cy.wait(500);
          expect(
            fetched.filter((u) => u.includes("should-not-be-fetched")),
            "the definition was stored, not run",
          ).to.deep.equal([]);
        });
      });
    });
  });
});
