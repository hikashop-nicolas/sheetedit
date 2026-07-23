/// <reference types="cypress" />

// Cypress runs in en-US, so dialog labels/titles are the English strings.
const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("sparklines", () => {
  it("creates a sparkline from a range through the insert dialog", () => {
    open("cypress/fixtures/sample.xlsx");
    // Select the numeric row B2:D2.
    cy.get('input[aria-label="B2"]').focus();
    cy.get('input[aria-label="D2"]').trigger("mousedown", { shiftKey: true });
    cy.get('.sheetedit-toolbar [title="Sparkline"]').click();
    cy.get(".sheetedit-form-modal", { timeout: TIMEOUT }).should("be.visible");
    // Data range prefilled from the selection; location prefilled just past it.
    cy.get('.sheetedit-form-modal [data-field="data"]').should("have.value", "B2:D2");
    cy.get('.sheetedit-form-modal [data-field="loc"]').clear().type("E2");
    cy.get('.sheetedit-form-modal [data-role="ok"]').click();
    cy.get(".sheetedit-form-modal").should("not.exist");
    // The host cell (E2 = row 2, col 5) now carries a sparkline canvas.
    cy.get('td[data-rc="2:5"] canvas.sheetedit-spark').should("exist");
  });

  it("shows a negative-colour field only for win/loss and column types", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="B2"]').focus();
    cy.get('.sheetedit-toolbar [title="Sparkline"]').click();
    cy.get(".sheetedit-form-modal", { timeout: TIMEOUT }).should("be.visible");
    // Line (default): negative colour hidden.
    cy.get('.sheetedit-form-modal [data-field="negColor"]').should("not.be.visible");
    cy.get('.sheetedit-form-modal [data-field="type"]').select("stacked");
    cy.get('.sheetedit-form-modal [data-field="negColor"]').should("be.visible");
    cy.get('.sheetedit-form-modal [data-role="cancel"]').click();
  });

  it("deletes a sparkline from the float bar", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="B2"]').focus();
    cy.get('input[aria-label="D2"]').trigger("mousedown", { shiftKey: true });
    cy.get('.sheetedit-toolbar [title="Sparkline"]').click();
    cy.get('.sheetedit-form-modal [data-field="loc"]').clear().type("E2");
    cy.get('.sheetedit-form-modal [data-role="ok"]').click();
    cy.get('td[data-rc="2:5"] canvas.sheetedit-spark').should("exist");
    // Select the host cell and hover it to reveal the float bar with sparkline actions.
    cy.get('input[aria-label="E2"]').focus();
    cy.get('td[data-rc="2:5"]').trigger("mousemove", "center");
    cy.get(".sheetedit-floatbar", { timeout: TIMEOUT }).should("be.visible");
    cy.get('.sheetedit-floatbar [title="Delete sparkline"]').click();
    cy.get('td[data-rc="2:5"] canvas.sheetedit-spark').should("not.exist");
  });
});
