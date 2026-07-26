/// <reference types="cypress" />
const TIMEOUT = 15000;
function open(f: string) {
  cy.visit("/");
  cy.get("#file").selectFile(f, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}
// The editor covers its whole cell when the text is middle-aligned and shrinks to its own height
// at the top or the bottom, so the box CENTRE is where the text actually sits; the box top is not.
const inTop = (ref: string) =>
  cy.get(`input[aria-label="${ref}"]`).then(($e) => {
    const r = $e[0].getBoundingClientRect();
    return r.top + r.height / 2;
  });

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
