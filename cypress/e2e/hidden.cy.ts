/// <reference types="cypress" />

// hidden.xlsx (20 rows x 8 cols) has row 3 and column C hidden (<row hidden>, <col hidden>).
// Hidden lines collapse to zero size: their cells are not rendered, and the neighbours abut.

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

const rightOf = (ref: string) => cy.get(`input[aria-label="${ref}"]`).then(($el) => $el[0].getBoundingClientRect().right);
const leftOf = (ref: string) => cy.get(`input[aria-label="${ref}"]`).then(($el) => $el[0].getBoundingClientRect().left);
const topOf = (ref: string) => cy.get(`input[aria-label="${ref}"]`).then(($el) => $el[0].getBoundingClientRect().top);

describe("hidden rows and columns", () => {
  it("collapses a hidden row so its neighbours are adjacent", () => {
    open("cypress/fixtures/hidden.xlsx");
    cy.get('input[aria-label="A2"]').should("exist");
    cy.get('input[aria-label="A4"]').should("exist");
    cy.get('input[aria-label="A3"]').should("not.exist"); // hidden row is not rendered
    // Row 4 sits directly under row 2 (row 3 took no vertical space).
    topOf("A2").then((t2) => topOf("A4").then((t4) => expect(t4 - t2).to.be.lessThan(30))); // ~one ROW_H (24)
  });

  it("collapses a hidden column so its neighbours are adjacent", () => {
    open("cypress/fixtures/hidden.xlsx");
    cy.get('input[aria-label="B1"]').should("exist");
    cy.get('input[aria-label="D1"]').should("exist");
    cy.get('input[aria-label="C1"]').should("not.exist"); // hidden column is not rendered
    // Column D starts where column B ends (column C took no horizontal space).
    rightOf("B1").then((bR) => leftOf("D1").then((dL) => expect(Math.abs(dL - bR)).to.be.lessThan(2)));
  });
});
