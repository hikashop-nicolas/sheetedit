/// <reference types="cypress" />
const TIMEOUT = 15000;
function open(f: string) {
  cy.visit("/");
  cy.get("#file").selectFile(f, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}
describe("style patch (no full rebuild)", () => {
  it("applies bold by patching the existing cell node, not rebuilding the grid", () => {
    cy.viewport(1400, 800);
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A1"]').click();
    // Tag the cell's DOM node; a full grid rebuild would replace it and lose the tag.
    cy.get('input[aria-label="A1"]').then(($el) => (($el[0] as unknown as { __m?: string }).__m = "keep"));
    cy.get('.sheetedit-toolbar [aria-label="Bold"]').click();
    cy.get('input[aria-label="A1"]').should("have.css", "font-weight").and("match", /700|bold/);
    cy.get('input[aria-label="A1"]').then(($el) => {
      expect(($el[0] as unknown as { __m?: string }).__m, "same DOM node -> patched in place").to.eq("keep");
    });
    // Un-bold removes it (the reset path) on the same node.
    cy.get('.sheetedit-toolbar [aria-label="Bold"]').click();
    cy.get('input[aria-label="A1"]').should("have.css", "font-weight").and("match", /400|normal/);
  });
});
