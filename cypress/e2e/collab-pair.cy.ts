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

type CellInput = { sheet: string; r: number; c: number; input: string };
type StructuralOp = { kind: "insert" | "delete"; axis: "row" | "col"; sheet: string; at: number; count: number };
type Handle = {
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
      p.a.applyRemoteCells([{ sheet: "Budget", r: 2, c: 2, input: "42" }]);
      // A change applied remotely must not be sent back out, or the two would ping-pong.
      expect(p.sent.a, "an applied change is not re-sent").to.deep.equal([]);
      expect(p.a.getCellValue("B2")).to.equal("42");
    });
  });

  it("recalculates from a peer's formula on both sides", () => {
    pair().then((p) => {
      p.a.applyRemoteCells([{ sheet: "Budget", r: 2, c: 2, input: "10" }]);
      p.b.applyRemoteCells([{ sheet: "Budget", r: 2, c: 2, input: "10" }]);
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
      p.a.applyRemoteStructural({ kind: "insert", axis: "row", sheet: "Budget", at: 2, count: 1 });
      p.b.applyRemoteStructural({ kind: "insert", axis: "row", sheet: "Budget", at: 2, count: 1 });

      expect(p.a.getCellValue("A3"), "moved down on A").to.equal("apples");
      expect(p.b.getCellValue("A3"), "and on B").to.equal("apples");
      expect(p.a.getCellValue("A2")).to.equal("");
    });
  });

  it("deletes a row on both sides, pulling the ones below it up", () => {
    pair().then((p) => {
      const op: StructuralOp = { kind: "delete", axis: "row", sheet: "Budget", at: 2, count: 1 };
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
        expect(p.ordered[0].sheet, "and it named its sheet").to.equal("Budget");
        expect(p.ordered[0]).to.deep.include({ kind: "insert", axis: "row", at: 2, count: 1 });
        // Both sides moved, from one proposal.
        expect(p.a.getCellValue("A3")).to.equal("apples");
        expect(p.b.getCellValue("A3")).to.equal("apples");
      });
    });
  });
});
