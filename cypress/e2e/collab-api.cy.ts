/// <reference types="cypress" />

// The API a collaboration host drives, in a real browser.
//
// Sharing a workbook means sharing what people typed, not what was computed: the formula
// engine is deterministic, so every peer recalculates the same results from the same
// inputs, and the shared state stays tiny. So the two things tested here are that the
// inputs are what comes out, and that a peer's edit arriving does NOT look like a local
// one, because that would echo back to whoever sent it.

const TIMEOUT = 15000;

type CellInput = { sheet: string; r: number; c: number; input: string };
type Handle = {
  cellInputs(): CellInput[];
  applyRemoteCells(changes: CellInput[]): void;
  sheetNames(): string[];
  getCellValue(ref: string): string;
  setPeerCells(peers: { id: string; colour: string; name: string; sheet: string; r: number; c: number }[]): void;
  applyRemoteStructural(op: { kind: "insert" | "delete"; axis: "row" | "col"; sheet: string; at: number; count: number }): void;
};

const win = (): Cypress.Chainable<Record<string, unknown>> =>
  cy.window().then((w) => w as unknown as Record<string, unknown>);
const handle = (): Cypress.Chainable<Handle> =>
  cy.window().then((w) => (w as unknown as { seHandle: Handle }).seHandle);

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("the collaboration API", () => {
  it("reports what was typed, not what was computed", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      const inputs = h.cellInputs();
      const c2 = inputs.find((i) => i.r === 2 && i.c === 3);
      // C2 shows 6 in the grid; what travels is the formula that produced it.
      expect(c2?.input, "a formula cell sends its formula").to.match(/^=/);
      expect(h.getCellValue("C2")).to.equal("6");
      // And a literal sends its literal.
      expect(inputs.find((i) => i.r === 1 && i.c === 1)?.input).to.equal("item");
      // Empty cells are not worth sending.
      expect(inputs.every((i) => i.input !== "")).to.equal(true);
    });
  });

  it("names its sheets, which is how addresses are qualified", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      expect(h.sheetNames()).to.include("Budget");
    });
  });

  it("announces which cells changed, and to what", () => {
    open("cypress/fixtures/sample.xlsx");
    win().then((w) => (w.seCellChanges = []));
    cy.get('input[aria-label="B2"]').clear().type("5").blur();

    win().then((w) => {
      const changes = w.seCellChanges as CellInput[];
      expect(changes.length, "a local edit is announced").to.be.greaterThan(0);
      const b2 = changes.find((c) => c.r === 2 && c.c === 2);
      expect(b2?.input).to.equal("5");
      expect(b2?.sheet).to.equal("Budget");
    });
  });

  it("applies a remote edit, and recalculates from it", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.applyRemoteCells([{ sheet: "Budget", r: 2, c: 2, input: "10" }]);
    });
    cy.get('input[aria-label="B2"]').should("have.value", "10");
    cy.get('input[aria-label="C2"]').should("have.value", "20"); // B2*2, recalculated locally
  });

  it("applies a remote formula, not only a literal", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.applyRemoteCells([{ sheet: "Budget", r: 2, c: 5, input: "=B2*10" }]);
    });
    cy.get('input[aria-label="E2"]').should("have.value", "30"); // B2 is 3 in the fixture
  });

  // The no-echo guarantee, which is the whole reason applyRemoteCells exists separately
  // from setCellValue.
  it("does not report a remote edit as a local change", () => {
    open("cypress/fixtures/sample.xlsx");
    win().then((w) => {
      w.seChangeCount = 0;
      w.seCellChanges = [];
      const h = (w as unknown as { seHandle: Handle }).seHandle;
      h.applyRemoteCells([{ sheet: "Budget", r: 2, c: 2, input: "42" }]);
      expect(w.seChangeCount, "onChange must not fire for a remote edit").to.equal(0);
      expect((w.seCellChanges as CellInput[]).length, "nor onCellsChanged").to.equal(0);
    });
    cy.get('input[aria-label="B2"]').should("have.value", "42"); // it did apply, though
  });

  // A remote edit must not become a step in this person's undo history.
  it("keeps a remote edit out of the local undo stack", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="B2"]').clear().type("7").blur();
    cy.get('input[aria-label="B2"]').should("have.value", "7");

    handle().then((h) => h.applyRemoteCells([{ sheet: "Budget", r: 3, c: 2, input: "99" }]));
    cy.get('input[aria-label="B3"]').should("have.value", "99");

    // Undo is bound on the cell input, not the document, so it has to be typed there.
    cy.get('input[aria-label="B2"]').focus().type("{ctrl}z");

    // My edit was taken back; theirs was not, because it was never on my stack.
    cy.get('input[aria-label="B2"]').should("not.have.value", "7");
    cy.get('input[aria-label="B3"]').should("have.value", "99");
  });

  it("ignores a cell for a sheet this workbook does not have", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.applyRemoteCells([{ sheet: "NoSuchSheet", r: 1, c: 1, input: "boom" }]);
      expect(h.getCellValue("A1")).to.equal("item"); // untouched
    });
  });
});

// Seeing where the other people are. Entirely visual, so this is the only place to test it.
describe("peer presence", () => {
  it("reports which cell this person moved to", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="C3"]').focus();
    win().then((w) => {
      expect(w.seSelection).to.deep.equal({ sheet: "Budget", r: 3, c: 3 });
    });
  });

  it("marks the cell another person is on, in their colour", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.setPeerCells([{ id: "p1", colour: "rgb(255, 0, 0)", name: "Ada", sheet: "Budget", r: 2, c: 2 }]);
    });

    cy.get('input[aria-label="B2"]').parent().should("have.class", "sheetedit-peer");
    cy.get('input[aria-label="B2"]').parent().find(".sheetedit-peerflag")
      .should("have.length", 1).and("have.text", "Ada");
    cy.get('input[aria-label="B2"]')
      .parent()
      .should("have.css", "box-shadow")
      .and("contain", "rgb(255, 0, 0)");
    cy.get('input[aria-label="B3"]').parent().should("not.have.class", "sheetedit-peer");
  });

  it("moves the marker when they move, and clears it when they leave", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.setPeerCells([{ id: "p1", colour: "rgb(0, 128, 0)", name: "Ada", sheet: "Budget", r: 2, c: 2 }]);
      h.setPeerCells([{ id: "p1", colour: "rgb(0, 128, 0)", name: "Ada", sheet: "Budget", r: 4, c: 2 }]);
    });
    cy.get('input[aria-label="B2"]').parent().should("not.have.class", "sheetedit-peer");
    cy.get('input[aria-label="B4"]').parent().should("have.class", "sheetedit-peer");

    handle().then((h) => h.setPeerCells([]));
    cy.get('input[aria-label="B4"]').parent().should("not.have.class", "sheetedit-peer");
  });

  // Several people on one cell: the cell has one outline, so it can only carry one colour.
  // Each name gets its own badge in that person's colour, or two peers on the same cell
  // would be indistinguishable.
  it("gives everyone on the same cell their own badge, in their own colour", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.setPeerCells([
        { id: "p1", colour: "rgb(255, 0, 0)", name: "Ada", sheet: "Budget", r: 2, c: 2 },
        { id: "p2", colour: "rgb(0, 0, 255)", name: "Grace", sheet: "Budget", r: 2, c: 2 },
      ]);
    });

    const cell = () => cy.get('input[aria-label="B2"]').parent();
    cell().find(".sheetedit-peerflag").should("have.length", 2);
    cell().find(".sheetedit-peerflag").eq(0).should("have.text", "Ada");
    cell().find(".sheetedit-peerflag").eq(1).should("have.text", "Grace");
    cell().find(".sheetedit-peerflag").eq(0)
      .should("have.css", "background-color", "rgb(255, 0, 0)");
    cell().find(".sheetedit-peerflag").eq(1)
      .should("have.css", "background-color", "rgb(0, 0, 255)");
    cell().should("have.class", "sheetedit-peer"); // still one outline
  });

  it("removes a badge when only one of two peers moves away", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.setPeerCells([
        { id: "p1", colour: "rgb(255, 0, 0)", name: "Ada", sheet: "Budget", r: 2, c: 2 },
        { id: "p2", colour: "rgb(0, 0, 255)", name: "Grace", sheet: "Budget", r: 2, c: 2 },
      ]);
      h.setPeerCells([
        { id: "p1", colour: "rgb(255, 0, 0)", name: "Ada", sheet: "Budget", r: 2, c: 2 },
        { id: "p2", colour: "rgb(0, 0, 255)", name: "Grace", sheet: "Budget", r: 3, c: 2 },
      ]);
    });
    cy.get('input[aria-label="B2"]').parent().find(".sheetedit-peerflag")
      .should("have.length", 1).and("have.text", "Ada");
    cy.get('input[aria-label="B3"]').parent().find(".sheetedit-peerflag")
      .should("have.length", 1).and("have.text", "Grace");
  });

  // The grid is virtualized, so a marker applied once would vanish on the next render.
  it("keeps the marker when a remote edit re-renders the grid", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.setPeerCells([{ id: "p1", colour: "rgb(255, 0, 0)", name: "Ada", sheet: "Budget", r: 2, c: 2 }]);
      h.applyRemoteCells([{ sheet: "Budget", r: 3, c: 2, input: "77" }]);
    });
    cy.get('input[aria-label="B3"]').should("have.value", "77");
    cy.get('input[aria-label="B2"]').parent().should("have.class", "sheetedit-peer");
  });

  it("does not mark a cell on a sheet nobody is looking at", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.setPeerCells([{ id: "p1", colour: "rgb(255, 0, 0)", name: "Ada", sheet: "NoSuchSheet", r: 2, c: 2 }]);
    });
    cy.get('input[aria-label="B2"]').parent().should("not.have.class", "sheetedit-peer");
  });
});

// Structural edits during a session. Cells are addressed by row and column, so inserting a
// row shifts every address below it on one side only: the two documents would diverge with
// nobody told, which is the failure mode worth spending code to prevent.
describe("structural edits", () => {
  function insertRowViaMenu() {
    cy.get("th.rownum").contains("2").rightclick();
    cy.get(".sheetedit-pop button").first().click(); // "insert row above"
  }

  it("asks the host before inserting a row", () => {
    open("cypress/fixtures/sample.xlsx");
    win().then((w) => {
      w.seStructural = [];
      w.seBlockStructural = false;
    });
    insertRowViaMenu();
    win().then((w) => {
      const asked = w.seStructural as { kind: string; axis: string }[];
      expect(asked.length, "the host is asked").to.be.greaterThan(0);
      expect(asked[0].axis).to.equal("row");
      expect(asked[0].kind).to.equal("insert");
      // The sheet comes from the editor: a row-header insert has no selected cell, so a
      // host inferring it from the selection would refuse the operation silently.
      expect((asked[0] as unknown as { sheet: string }).sheet).to.equal("Budget");
    });
  });

  it("does not insert when the host refuses", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A2"]').should("have.value", "apples");
    win().then((w) => (w.seBlockStructural = true));

    insertRowViaMenu();

    // A2 still holds what it held: nothing shifted down.
    cy.get('input[aria-label="A2"]').should("have.value", "apples");
  });
});

// A structural edit decided elsewhere. In a shared session one peer puts these in order for
// everyone, and the others apply the result rather than deciding for themselves.
describe("a structural edit from elsewhere", () => {
  it("inserts a row and moves the cells below it down", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A2"]').should("have.value", "apples");

    handle().then((h) => {
      h.applyRemoteStructural({ kind: "insert", axis: "row", sheet: "Budget", at: 2, count: 1 });
    });

    cy.get('input[aria-label="A2"]').should("have.value", ""); // the new blank row
    cy.get('input[aria-label="A3"]').should("have.value", "apples"); // pushed down
  });

  it("deletes a row and pulls the ones below it up", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.applyRemoteStructural({ kind: "delete", axis: "row", sheet: "Budget", at: 2, count: 1 });
    });
    cy.get('input[aria-label="A2"]').should("not.have.value", "apples");
  });

  it("does not ask permission: the decision was already made", () => {
    open("cypress/fixtures/sample.xlsx");
    win().then((w) => {
      w.seStructural = [];
      w.seBlockStructural = true; // this peer refuses its OWN structural edits
      const h = (w as unknown as { seHandle: Handle }).seHandle;
      h.applyRemoteStructural({ kind: "insert", axis: "row", sheet: "Budget", at: 2, count: 1 });
      expect(w.seStructural, "the veto is for local edits only").to.deep.equal([]);
    });
    cy.get('input[aria-label="A3"]').should("have.value", "apples"); // it still applied
  });

  it("ignores an operation for a sheet this workbook does not have", () => {
    open("cypress/fixtures/sample.xlsx");
    handle().then((h) => {
      h.applyRemoteStructural({ kind: "delete", axis: "row", sheet: "NoSuchSheet", at: 1, count: 5 });
    });
    cy.get('input[aria-label="A2"]').should("have.value", "apples");
  });
});
