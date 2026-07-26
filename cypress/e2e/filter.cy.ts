/// <reference types="cypress" />
import { clickToolbar } from "../support/toolbar";

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("autofilter", () => {
  it("toggles a filter scoped to the data region and places carets inside each header cell", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A1"]').trigger("mousedown"); // no explicit range -> current region
    clickToolbar("Toggle filter");
    // A caret appears in each header-row cell of the region.
    cy.get("td.has-filter .sheetedit-filterbtn").should("have.length.greaterThan", 0);
    // Each caret sits at the right edge of its own cell and is one cell tall (guards the bug where
    // an unpositioned header cell let the caret escape to the grid edge, full height).
    cy.get("td.has-filter").each(($td) => {
      const cell = $td[0].getBoundingClientRect();
      const btn = $td[0].querySelector(".sheetedit-filterbtn") as HTMLElement;
      const b = btn.getBoundingClientRect();
      expect(b.right, "caret right aligns with cell right").to.be.closeTo(cell.right, 2);
      expect(b.height, "caret is one cell tall, not full grid").to.be.lessThan(cell.height + 4);
    });
  });

  it("opens the filter menu from a caret", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A1"]').trigger("mousedown");
    clickToolbar("Toggle filter");
    cy.get("td.has-filter .sheetedit-filterbtn").first().click({ force: true });
    cy.get(".sheetedit-filtermenu", { timeout: TIMEOUT }).should("be.visible");
  });

  it("removes the filter when toggled again", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A1"]').trigger("mousedown");
    clickToolbar("Toggle filter");
    cy.get("td.has-filter").should("have.length.greaterThan", 0);
    clickToolbar("Toggle filter");
    cy.get("td.has-filter").should("have.length", 0);
  });
});
