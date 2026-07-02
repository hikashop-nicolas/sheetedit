/// <reference types="cypress" />

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("sheetedit", () => {
  it("renders an .xlsx as a grid with computed formula values", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('.sheetedit-tab[aria-selected="true"]').should("contain.text", "Budget");
    cy.get('input[aria-label="A1"]').should("have.value", "item");
    cy.get('input[aria-label="C2"]').should("have.value", "6"); // B2*2
    cy.get('input[aria-label="C4"]').should("have.value", "14"); // SUM(C2:C3)
  });

  it("shows formatted numbers but edits the raw value", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="D2"]').should("have.value", "$3.50"); // currency format applied
    cy.get('input[aria-label="D2"]').focus().should("have.value", "3.5"); // raw value when editing
    cy.get('input[aria-label="D2"]').clear().type("9.9").blur();
    cy.get('input[aria-label="D2"]').should("have.value", "$9.90"); // typed value keeps the format
  });

  it("recalculates formulas when a dependency changes", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="B2"]').clear().type("5").blur();
    cy.get('input[aria-label="C2"]').should("have.value", "10"); // 5*2
    cy.get('input[aria-label="C4"]').should("have.value", "18"); // 10 + 8
  });

  it("shows the formula when a formula cell is focused", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="C2"]').focus().should("have.value", "=B2*2");
  });

  it("Escape cancels an edit without committing the display text", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="C2"]').focus().clear().type("=B2*99").type("{esc}");
    cy.get('input[aria-label="C2"]').should("have.value", "6"); // display restored
    cy.get('input[aria-label="C2"]').focus().should("have.value", "=B2*2"); // formula intact
    // Escape on a formatted cell must not commit the "$3.50" display as a string.
    cy.get('input[aria-label="D2"]').focus().type("{esc}");
    cy.get('input[aria-label="D2"]').should("have.value", "$3.50");
    cy.get('input[aria-label="D2"]').focus().should("have.value", "3.5");
  });

  it("edits, exports a valid .xlsx, and round-trips with recalculated values", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.window().then((win) => {
      (win as unknown as { __exported: Uint8Array | null }).__exported = null;
      const orig = win.URL.createObjectURL.bind(win.URL);
      win.URL.createObjectURL = (b: Blob) => {
        if (b instanceof win.Blob)
          void b.arrayBuffer().then((ab) => ((win as unknown as { __exported: Uint8Array }).__exported = new Uint8Array(ab)));
        return orig(b);
      };
    });
    cy.get('input[aria-label="B2"]').clear().type("5").blur();
    cy.get("#save").click();
    cy.window().its("__exported").should("exist");
    cy.window().then((win) => {
      const bytes = (win as unknown as { __exported: Uint8Array }).__exported;
      const file = new win.File([bytes as BlobPart], "x.xlsx", {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const dt = new win.DataTransfer();
      dt.items.add(file);
      const inp = win.document.getElementById("file") as HTMLInputElement;
      inp.files = dt.files;
      inp.dispatchEvent(new win.Event("change", { bubbles: true }));
    });
    cy.get('input[aria-label="B2"]', { timeout: TIMEOUT }).should("have.value", "5");
    cy.get('input[aria-label="C2"]').should("have.value", "10");
    cy.get('input[aria-label="C4"]').should("have.value", "18");
  });

  it("renders and edits an .ods workbook", () => {
    open("cypress/fixtures/sample.ods");
    cy.get('input[aria-label="C2"]').should("have.value", "6"); // B2*2
    cy.get('input[aria-label="B2"]').clear().type("10").blur();
    cy.get('input[aria-label="C2"]').should("have.value", "20");
  });
});
