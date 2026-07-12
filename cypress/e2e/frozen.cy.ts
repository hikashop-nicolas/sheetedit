/// <reference types="cypress" />

// frozen.xlsx has the top row and first column frozen (<pane state="frozen">). Each cell's
// text is its own A1 ref. These tests scroll the grid and assert the frozen row/column stay
// on screen while the rest scrolls under them (real layout, so this runs in Chrome, not jsdom).

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("frozen panes", () => {
  // Viewport top of the cell input for the given A1 ref.
  const topOf = (ref: string) => cy.get(`input[aria-label="${ref}"]`).then(($el) => $el[0].getBoundingClientRect().top);

  it("keeps the top row pinned while scrolling down", () => {
    open("cypress/fixtures/frozen.xlsx");
    cy.get('.sheetedit-tab[aria-selected="true"]').should("contain.text", "Frozen");
    cy.get('input[aria-label="A1"]').should("have.value", "A1");

    cy.get(".sheetedit-grid").scrollTo(0, 700);
    cy.get('input[aria-label="D35"]', { timeout: TIMEOUT }).should("be.visible"); // scrolled down
    cy.get('input[aria-label="D1"]').should("be.visible"); // frozen top row still on screen
    // The frozen row sits above the visible body rows (pinned at the top, not floating mid-grid).
    topOf("D1").then((rowTop) => topOf("D35").then((bodyTop) => expect(rowTop).to.be.lessThan(bodyTop)));
    cy.get('input[aria-label="A1"]').should("be.visible"); // frozen corner still on screen
  });

  it("keeps the first column pinned while scrolling right", () => {
    open("cypress/fixtures/frozen.xlsx");
    cy.get(".sheetedit-grid").scrollTo(0, 400); // bring some lower rows into the window
    cy.get(".sheetedit-grid").scrollTo(600, 400);
    cy.get('input[aria-label="J20"]', { timeout: TIMEOUT }).should("be.visible"); // scrolled right
    cy.get('input[aria-label="A20"]').should("be.visible"); // frozen first column still on screen
  });
});
