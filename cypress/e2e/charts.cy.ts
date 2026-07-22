/// <reference types="cypress" />

// Cypress runs in en-US, so the toolbar titles and dialog labels are the English strings.
const TIMEOUT = 20000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("charts", () => {
  it("renders an existing chart over the grid", () => {
    open("cypress/fixtures/chart.xlsx");
    cy.get(".sheetedit-chartbox", { timeout: TIMEOUT }).should("have.length", 1);
    cy.get(".sheetedit-chartbox canvas").should("exist");
  });

  it("creates a new chart from a range through the insert dialog", () => {
    open("cypress/fixtures/chart.xlsx");
    cy.get(".sheetedit-chartbox", { timeout: TIMEOUT }).should("have.length", 1);
    cy.get('[title="Insert chart"]').click();
    cy.get(".sheetedit-chart-modal", { timeout: TIMEOUT }).should("be.visible");
    cy.get(".sheetedit-chart-modal input[type=text]").first().clear().type("Sheet1!A1:C4");
    cy.contains(".sheetedit-chart-type", "Line").click();
    cy.contains(".sheetedit-chart-actions button", "Insert").click();
    cy.get(".sheetedit-chart-modal").should("not.exist");
    cy.get(".sheetedit-chartbox", { timeout: TIMEOUT }).should("have.length", 2);
  });
});
