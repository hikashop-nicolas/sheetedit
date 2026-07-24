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
});
