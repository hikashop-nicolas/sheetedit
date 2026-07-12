/// <reference types="cypress" />
const TIMEOUT = 15000;
function open(f: string) {
  cy.visit("/");
  cy.get("#file").selectFile(f, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}
describe("wrap text", () => {
  it("grows the row and wraps the text of a wrap cell", () => {
    open("cypress/fixtures/wrap.xlsx"); // A1 = long text, wrap on, narrow column
    cy.get('input[aria-label="A1"]').closest("td").as("a1");
    cy.get("@a1").should("have.class", "has-wrap");
    cy.get("@a1").find(".sheetedit-cellwrap").should("contain.text", "wrap across several lines");
    // The row grew well beyond the default (~24px) to fit the wrapped text.
    cy.get("@a1").then(($td) => expect($td[0].getBoundingClientRect().height).to.be.greaterThan(45));
    cy.screenshot("wrap", { overwrite: true, capture: "viewport" });
  });
});
