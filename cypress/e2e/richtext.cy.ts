/// <reference types="cypress" />

// Styling part of a cell's text. sample.xlsx A1 is the plain string "item"; bolding the first two
// letters turns the cell into a two-run one, and the result has to be on screen at once. It was not:
// the caret was still in the cell, and while the cell held the caret the grid handed it to the plain
// <input>, so the new run only appeared once the caret moved somewhere else.

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("styling part of a cell's text", () => {
  it("shows the new run straight away, without moving the caret out", () => {
    cy.viewport(1400, 800); // wide enough that the style cluster stays inline
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A1"]').should("have.value", "item").click();
    cy.get('input[aria-label="A1"]').click(); // a second click puts the caret in the text: now editing
    cy.get('input[aria-label="A1"]').closest("td").should("have.class", "editing");
    // Select "it" inside the cell: a sub-range is what makes the change a run rather than a
    // whole-cell style.
    cy.get('input[aria-label="A1"]').then(($i) => {
      const el = $i[0] as HTMLInputElement;
      el.focus();
      el.setSelectionRange(0, 2);
    });
    cy.contains(".sheetedit-toolbar button", "B").click();
    cy.get('input[aria-label="A1"]').closest("td").as("a1");
    cy.get("@a1").should("have.class", "has-rich");
    cy.get("@a1").should("not.have.class", "editing"); // the styled text, not the input, is what shows
    cy.get("@a1").find(".sheetedit-cellrich").should("be.visible").should("have.text", "item");
    cy.get("@a1").find(".sheetedit-cellrich span").first().should("have.css", "font-weight", "700");
    cy.get("@a1").find(".sheetedit-cellrich span").last().should("not.have.css", "font-weight", "700");
  });
});
