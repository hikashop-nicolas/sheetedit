/// <reference types="cypress" />

// End-to-end for the Power Query editor. The Cypress browser runs in en-US, so the toolbar
// titles and dialog labels are the English strings.

const TIMEOUT = 20000;

function openEditor() {
  cy.visit("/");
  cy.get("#file").selectFile("cypress/fixtures/pq.xlsx", { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
  cy.get('[title="Edit queries"]', { timeout: TIMEOUT }).click();
  cy.get(".se-pqe", { timeout: TIMEOUT }).should("be.visible");
}

describe("Power Query editor", () => {
  it("lists queries and steps, previews, and a transform appends a step", () => {
    openEditor();
    cy.get(".se-pqe-queries .se-pqe-item-name").should("have.length.greaterThan", 0);
    cy.get(".se-pqe-settings .se-pqe-item-name").its("length").should("be.gt", 1);
    cy.get(".se-pqe-ptable", { timeout: TIMEOUT }).should("exist");

    cy.contains(".se-pqe-rbtn", "Keep top rows").click();
    cy.get(".se-pqe-card input").clear().type("1");
    cy.contains(".se-pqe-card-actions button", "Apply").click();
    cy.get(".se-pqe-settings .se-pqe-item-name").should("contain.text", "Kept First Rows");
    cy.get(".se-pqe-ptable tbody tr").should("have.length", 1);
  });

  it("creates a query from a workbook table and loads it onto a new sheet", () => {
    openEditor();
    cy.get(".se-pqe-newq").click();
    cy.get(".se-pqe-card select").select("table");
    cy.contains(".se-pqe-card-actions button", "Apply").click();
    cy.get(".se-pqe-card select").first().select("Sales");
    cy.contains(".se-pqe-card-actions button", "Apply").click();
    cy.get(".se-pqe-queries .se-pqe-item.sel .se-pqe-item-name").should("contain.text", "Query1");
    cy.get(".se-pqe-ptable th .nm").should("contain.text", "Product");

    cy.contains(".se-pqe-bar button", "Load").click();
    cy.contains(".se-pqe-foot", "Loaded", { timeout: TIMEOUT });
    cy.contains(".se-pqe-bar button", "Cancel").click();
    cy.get(".sheetedit-tab").should("contain.text", "Query1");
  });

  it("collapses the side panels into drawers on a narrow viewport", () => {
    cy.viewport(414, 820);
    openEditor();
    // Ribbon labels are hidden (icon-only) and the pane toggles appear.
    cy.get(".se-pqe-rbtn span").first().should("not.be.visible");
    cy.get(".se-pqe-panetoggle").should("be.visible").and("have.length", 2);
    // The Queries drawer opens on toggle and closes when the preview is tapped.
    cy.get(".se-pqe").should("not.have.class", "show-queries");
    cy.get(".se-pqe-panetoggle").first().click();
    cy.get(".se-pqe").should("have.class", "show-queries");
    cy.get(".se-pqe-center").click("center");
    cy.get(".se-pqe").should("not.have.class", "show-queries");
  });
});
