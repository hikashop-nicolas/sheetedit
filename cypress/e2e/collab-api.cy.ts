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
