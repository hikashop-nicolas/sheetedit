/// <reference types="cypress" />
const TIMEOUT = 15000;
function open(f: string) {
  cy.visit("/");
  cy.get("#file").selectFile(f, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}
const inTop = (ref: string) => cy.get(`input[aria-label="${ref}"]`).then(($e) => $e[0].getBoundingClientRect().top);

describe("vertical alignment", () => {
  it("aligns cell text top / middle / bottom within a tall row", () => {
    open("cypress/fixtures/valign.xlsx"); // A1=top, B1=middle, C1=bottom, row 1 is tall
    inTop("A1").then((a) =>
      inTop("B1").then((b) =>
        inTop("C1").then((c) => {
          expect(a, "top above middle").to.be.lessThan(b);
          expect(b, "middle above bottom").to.be.lessThan(c);
        }),
      ),
    );
    cy.screenshot("valign", { overwrite: true, capture: "viewport" });
  });
});
