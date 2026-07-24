/// <reference types="cypress" />

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("pivot tables", () => {
  it("outlines and labels the xlsx pivot output on its sheet", () => {
    open("cypress/fixtures/pivot.xlsx");
    // The pivot lands on the "Pivot" sheet; the Data sheet shows no outline.
    cy.get(".sheetedit-pivotbox").should("not.exist");
    cy.get(".sheetedit-tab").contains("Pivot").click();
    cy.get(".sheetedit-pivotbox", { timeout: TIMEOUT }).should("exist");
    cy.get(".sheetedit-pivottag").should("contain.text", "PivotTable1");
  });

  it("outlines the ODS data-pilot output on its sheet", () => {
    open("cypress/fixtures/pivot.ods");
    cy.get(".sheetedit-tab").contains("Pivot").click();
    cy.get(".sheetedit-pivotbox", { timeout: TIMEOUT }).should("exist");
    cy.get(".sheetedit-pivottag").should("contain.text", "PivotTable1");
  });

  it("authors a pivot from a selection through the insert dialog", () => {
    open("cypress/fixtures/pivot.xlsx"); // opens on the Data sheet
    // Select the source range A1:C7 (header + data).
    cy.get('input[aria-label="A1"]').focus();
    cy.get('input[aria-label="C7"]').trigger("mousedown", { shiftKey: true });
    cy.get('.sheetedit-toolbar [title="Insert pivot table"]').click();
    cy.get(".sheetedit-form-modal", { timeout: TIMEOUT }).should("be.visible");
    // Defaults: Region=Rows, Sales=Values(Sum). Make Product a column field for a crosstab.
    cy.get('.sheetedit-form-modal [data-field="role_1"]').select("columns");
    cy.get('.sheetedit-form-modal [data-role="ok"]').click();
    cy.get(".sheetedit-form-modal").should("not.exist");
    // A new sheet holds the pivot, outlined + labelled, with the grand total materialised.
    cy.get(".sheetedit-pivotbox", { timeout: TIMEOUT }).should("exist");
    cy.get(".sheetedit-pivottag").should("contain.text", "PivotTable");
    cy.get('input[aria-label="D4"]').should("have.value", "350");
  });

  it("authors a nested pivot with subtotals", () => {
    open("cypress/fixtures/pivot.xlsx");
    cy.get('input[aria-label="A1"]').focus();
    cy.get('input[aria-label="C7"]').trigger("mousedown", { shiftKey: true });
    cy.get('.sheetedit-toolbar [title="Insert pivot table"]').click();
    cy.get(".sheetedit-form-modal", { timeout: TIMEOUT }).should("be.visible");
    // Region + Product both on rows, Sum of Sales, subtotals on.
    cy.get('.sheetedit-form-modal [data-field="role_1"]').select("rows");
    cy.get('.sheetedit-form-modal [data-field="subtotals"]').check();
    cy.get('.sheetedit-form-modal [data-role="ok"]').click();
    cy.get(".sheetedit-form-modal").should("not.exist");
    // The nested output carries a per-group subtotal and a grand total.
    cy.get('input[aria-label="A4"]').should("have.value", "North Total");
    cy.get('input[aria-label="C4"]').should("have.value", "190");
    cy.get('input[aria-label="A8"]').should("have.value", "Grand Total");
    cy.get('input[aria-label="C8"]').should("have.value", "350");
  });

  it("edits an authored pivot in place via its tag menu", () => {
    open("cypress/fixtures/pivot.xlsx");
    cy.get('input[aria-label="A1"]').focus();
    cy.get('input[aria-label="C7"]').trigger("mousedown", { shiftKey: true });
    cy.get('.sheetedit-toolbar [title="Insert pivot table"]').click();
    cy.get(".sheetedit-form-modal", { timeout: TIMEOUT }).should("be.visible");
    cy.get('.sheetedit-form-modal [data-role="ok"]').click(); // Region rows, Sales values
    cy.get('input[aria-label="B2"]').should("have.value", "190"); // North total, one value column
    // Open the tag menu and choose Edit, then make Product a column field. (force: the row-1 tag
    // sits under the toolbar in Cypress's small default viewport; it is clickable in a real window.)
    cy.get(".sheetedit-pivottag").click({ force: true });
    cy.get(".sheetedit-pivot-menu").contains("Edit fields").click();
    cy.get(".sheetedit-form-modal", { timeout: TIMEOUT }).should("be.visible");
    cy.get('.sheetedit-form-modal [data-field="role_1"]').select("columns");
    cy.get('.sheetedit-form-modal [data-role="ok"]').click();
    // The pivot is now a crosstab in place.
    cy.get('input[aria-label="B1"]').should("have.value", "Apple");
    cy.get('input[aria-label="B2"]').should("have.value", "140");
    cy.get('input[aria-label="D4"]').should("have.value", "350");
  });
});
