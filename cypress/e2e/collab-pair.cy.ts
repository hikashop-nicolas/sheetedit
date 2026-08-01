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
type Handle = {
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
