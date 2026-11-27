/// <reference types="cypress" />

// On a phone the styling controls fold into the "Aa" menu. A dropdown opened from there used to
// be measured after that menu closed, so it landed above the editor, under the host's header.

const TIMEOUT = 15000;

/** The element's first item sits inside the visible area and is the topmost thing at its centre. */
function firstItemReachable(menuSel: string, itemSel: string) {
  cy.get(menuSel).should("be.visible").then(($m) => {
    const win = $m[0]!.ownerDocument.defaultView!;
    const doc = $m[0]!.ownerDocument;
    const editorTop = doc.querySelector(".sheetedit-wrap")!.getBoundingClientRect().top;
    const vv = win.visualViewport;
    const vh = vv ? vv.offsetTop + vv.height : win.innerHeight;
    const vw = vv ? vv.offsetLeft + vv.width : win.innerWidth;
    const m = $m[0]!.getBoundingClientRect();
    expect(m.top, "menu below the editor's top").to.be.at.least(editorTop);
    expect(m.bottom, "menu inside the viewport").to.be.at.most(vh);
    expect(m.right, "menu inside the viewport width").to.be.at.most(vw);
    expect($m[0]!.scrollTop, "opens at the top").to.eq(0);
    const first = $m[0]!.querySelector(itemSel)!;
    const r = first.getBoundingClientRect();
    expect(r.top).to.be.at.least(m.top);
    expect(r.bottom).to.be.at.most(Math.min(m.bottom, vh));
    const hit = doc.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
    expect(hit === first || first.contains(hit), "first item is on top").to.eq(true);
  });
}

describe("toolbar menus at phone width", () => {
  for (const [w, h] of [[375, 812], [390, 844]] as const) {
    it(`open under the toolbar, fully reachable, at ${w}x${h}`, () => {
      cy.viewport(w, h);
      cy.visit("/");
      cy.get("#file").selectFile("cypress/fixtures/sample.xlsx", { force: true });
      cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
      cy.get('.sheetedit-toolbar [title="Formatting"]').should("be.visible");

      for (const [title, menu, item] of [
        ["Font", ".sheetedit-fontmenu", ".sheetedit-font-item"],
        ["Font size", ".sheetedit-tb-listmenu:not(.sheetedit-fontmenu)", "button"],
        ["Number format", ".sheetedit-fmtmenu", "button"],
      ] as const) {
        const openIt = () => {
          cy.get(".sheetedit-wrap").then(($w) => {
            for (const m of $w[0]!.querySelectorAll<HTMLElement>(".sheetedit-tb-groupmenu")) m.hidden = true;
          });
          cy.get('.sheetedit-toolbar [title="Formatting"]').click();
          cy.get(`.sheetedit-tb-groupmenu [title="${title}"]`).click();
        };
        openIt();
        firstItemReachable(menu, item);
        // Scrolled, closed and reopened: back at the top, not where it was left.
        cy.get(menu).then(($m) => { $m[0]!.scrollTop = 200; });
        openIt();
        firstItemReachable(menu, item);
      }
    });
  }
});
