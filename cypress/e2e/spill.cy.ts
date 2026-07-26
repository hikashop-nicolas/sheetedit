/// <reference types="cypress" />
const TIMEOUT = 15000;

// Text longer than its column runs on over the empty cells beside it, and stops at the first one
// that holds something. Both halves matter: a heading that vanishes at the gridline is wrong, and
// a heading that runs over a neighbour's own text is worse.
describe("text spill", () => {
  it("spills over empty neighbours and stops at an occupied one", () => {
    cy.visit("/");
    cy.get("#file").selectFile("cypress/fixtures/sample.xlsx", { force: true });
    cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");

    cy.get('input[aria-label="A10"]').click().type("A heading far too long for one column{enter}");
    cy.get('input[aria-label="A10"]').closest("td").as("a10");
    cy.get("@a10").should("have.class", "has-spill");
    cy.get("@a10")
      .find(".sheetedit-cellspill")
      .then(($ov) => {
        const ov = $ov[0]!.getBoundingClientRect();
        const td = Cypress.$('input[aria-label="A10"]').closest("td")[0]!.getBoundingClientRect();
        expect(ov.right, "the overlay reaches past its own cell").to.be.greaterThan(td.right);
      });

    // Something in the way ends the spill.
    // force: the floating format bar sits over the cell below the one just edited.
    cy.get('input[aria-label="B10"]').click({ force: true }).type("stop{enter}", { force: true });
    cy.get('input[aria-label="A10"]').closest("td").should("not.have.class", "has-spill");
  });
});
