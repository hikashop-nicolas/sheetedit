/// <reference types="cypress" />

// Authoring a shape, and the grid it goes on.
//
// A new shape used to be anchored at the selection, which on a sheet nobody had clicked in yet is
// A1: scroll down a hundred rows, insert a shape, and it lands somewhere off screen. And the grid
// itself used to stop at the last row the file mentions, with a "+ rows" button as the only way
// past it, so a picture anchored below the data hung past the end of the sheet.

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

/** The scroller that holds the cells. */
const grid = () => cy.get(".sheetedit-grid").first();

describe("a new shape", () => {
  it("lands in the middle of what is on screen, not at the top of the sheet", () => {
    cy.viewport(1900, 420); // wide enough for the whole toolbar, short enough that the grid scrolls
    open("cypress/fixtures/sample.xlsx");
    // Somewhere well down the sheet, so "where the user is looking" and "A1" are far apart.
    grid().scrollTo(0, 200);
    cy.get('.sheetedit-toolbar [aria-label="Insert shape"]').click();
    cy.get(".sheetedit-shapegallery").should("be.visible");
    cy.get('.sheetedit-shapegallery [aria-label="Rectangle"]').first().click();
    cy.get(".sheetedit-shapebox").should("have.length", 1);
    // On screen: inside the grid's own rectangle, not above it or below it.
    cy.get(".sheetedit-shapebox").then(($box) => {
      const box = $box[0].getBoundingClientRect();
      cy.get(".sheetedit-grid").first().then(($g) => {
        const view = $g[0].getBoundingClientRect();
        expect(box.top, "below the top of the grid").to.be.greaterThan(view.top);
        expect(box.bottom, "above the bottom of it").to.be.lessThan(view.bottom);
        expect(box.left).to.be.greaterThan(view.left);
        expect(box.right).to.be.lessThan(view.right);
      });
    });
  });

  it("keeps its rotation grip clear of the float bar", () => {
    cy.viewport(1900, 800); // room above the shape for both the grip and the bar
    open("cypress/fixtures/sample.xlsx");
    cy.get('.sheetedit-toolbar [aria-label="Insert shape"]').click();
    cy.get('.sheetedit-shapegallery [aria-label="Rectangle"]').first().click();
    cy.get(".sheetedit-shapebox").click();
    cy.get(".sheetedit-shapebar").should("be.visible");
    cy.get(".sheetedit-shape-rotate").should("be.visible");
    // The grip hangs above the shape and the bar sits above the grip: a bar drawn right on top of
    // the shape covered it, and a handle under a toolbar cannot be grabbed.
    cy.get(".sheetedit-shape-rotate").then(($g) => {
      const grip = $g[0].getBoundingClientRect();
      cy.get(".sheetedit-shapebar").then(($b) => {
        const bar = $b[0].getBoundingClientRect();
        expect(bar.bottom, `the bar clears the grip (bar ${Math.round(bar.bottom)}, grip ${Math.round(grip.top)})`)
          .to.be.at.most(grip.top);
      });
    });
  });
});

describe("the end of the grid", () => {
  it("moves down as the scroll reaches it", () => {
    cy.viewport(1900, 420);
    open("cypress/fixtures/sample.xlsx");
    grid().then(($g) => {
      const before = $g[0].scrollHeight;
      grid().scrollTo("bottom");
      // Arriving at the bottom added a chunk of rows, so there is further to scroll than there was.
      grid().should(($after) => expect($after[0].scrollHeight, "the sheet got longer").to.be.greaterThan(before));
      // And the rows themselves run past where the file stopped.
      grid().scrollTo("bottom");
      cy.get(".sheetedit-table th.rownum", { timeout: TIMEOUT }).last().invoke("attr", "data-r").then(Number)
        .should("be.greaterThan", 24);
    });
  });

  it("has no + row / + column buttons left on the toolbar", () => {
    cy.viewport(1900, 420);
    open("cypress/fixtures/sample.xlsx");
    cy.get(".sheetedit-toolbar").should("not.contain.text", "+ Row");
    cy.get(".sheetedit-toolbar").should("not.contain.text", "+ Col");
  });
});
