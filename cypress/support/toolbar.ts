/// <reference types="cypress" />

// The toolbar folds whatever does not fit into a "⋯" menu, so a button's presence on screen
// depends on the viewport and on how many buttons exist today. Clicking one by title has to cope
// with both places, or a spec starts failing the day an unrelated button is added.

/** Click a toolbar button by its title, through the overflow menu when it has been folded away. */
export function clickToolbar(title: string): void {
  cy.get(".sheetedit-toolbar").then(($bar) => {
    const btn = $bar.find(`[title="${title}"]`);
    if (btn.length && btn.is(":visible")) {
      cy.get(`.sheetedit-toolbar [title="${title}"]`).click();
      return;
    }
    cy.get('.sheetedit-toolbar [title="More"]').click();
    // Each overflow row repeats the button's label as text and forwards the click to the original.
    cy.get(".sheetedit-tb-moremenu").contains("button", title).click();
  });
}
