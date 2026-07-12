/// <reference types="cypress" />

const TIMEOUT = 15000;
function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("keyboard accessibility", () => {
  it("border popover: opens focused, arrows roam, Escape closes and refocuses", () => {
    cy.viewport(1400, 800);
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A1"]').click();
    cy.get('.sheetedit-toolbar [aria-label="Borders"]').click();
    cy.get('.sheetedit-pop[role="menu"]').should("be.visible");
    cy.focused().should("have.class", "sheetedit-pop-item"); // focus moved into the menu
    cy.focused().type("{downArrow}");
    cy.focused().should("have.class", "sheetedit-pop-item"); // still on a menu item
    cy.focused().type("{esc}");
    cy.get(".sheetedit-pop").should("not.exist");
    cy.focused().should("have.attr", "aria-label", "Borders"); // focus returned to the trigger
  });
});
