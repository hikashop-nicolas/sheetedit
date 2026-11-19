/// <reference types="cypress" />

// Selecting with a finger, as in Excel mobile: a long-press starts a range, round grips on two
// corners resize it, and lifting the finger offers Cut, Copy, Paste, Clear and Fill.

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

type Win = Window & typeof globalThis;

/** Fire a touch pointer event at the centre of an element (or at x, y). */
function touch(doc: Document, type: string, el: Element, at?: { x: number; y: number }) {
  const r = el.getBoundingClientRect();
  const x = at?.x ?? r.left + r.width / 2;
  const y = at?.y ?? r.top + r.height / 2;
  const target = type === "pointerup" ? doc.defaultView! : el;
  target.dispatchEvent(new (doc.defaultView as Win).PointerEvent(type, {
    pointerType: "touch", pointerId: 7, isPrimary: true, buttons: type === "pointerup" ? 0 : 1, bubbles: true, cancelable: true, clientX: x, clientY: y,
  }));
}
const td = (doc: Document, ref: string) => doc.querySelector(`input[aria-label="${ref}"]`)!.closest("td")!;

describe("touch selection", () => {
  it("long-press selects, the grips resize, and the bar clears", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.document().then((doc) => touch(doc, "pointerdown", td(doc, "B2")));
    cy.wait(350); // the long-press
    cy.document().then((doc) => {
      touch(doc, "pointermove", td(doc, "B3"));
      touch(doc, "pointerup", td(doc, "B3"));
    });
    cy.get("td.sheetedit-sel").should("have.length", 2);
    cy.get(".sheetedit-grip").should("have.length", 2);
    cy.get(".sheetedit-fillhandle").should("not.exist");
    cy.get(".sheetedit-touchbar button").should("have.length", 5);

    // the bottom grip grows the range from B2:B3 to B2:C4
    cy.document().then((doc) => {
      const grip = doc.querySelector(".sheetedit-grip-br")!;
      touch(doc, "pointerdown", grip);
      const c4 = td(doc, "C4").getBoundingClientRect();
      touch(doc, "pointermove", grip, { x: c4.left + c4.width / 2, y: c4.top + c4.height / 2 });
      touch(doc, "pointerup", grip);
    });
    cy.get("td.sheetedit-sel").should("have.length", 6);
    cy.get(".sheetedit-touchbar").contains("button", "Clear").click();
    cy.get(".sheetedit-touchbar").should("not.exist");
    cy.get('input[aria-label="B2"]').should("have.value", "");
    cy.get('input[aria-label="C4"]').should("have.value", "");
  });

  it("Fill turns the corner grip into a fill handle", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.document().then((doc) => touch(doc, "pointerdown", td(doc, "B2")));
    cy.wait(350);
    cy.document().then((doc) => touch(doc, "pointerup", td(doc, "B2")));
    cy.get(".sheetedit-touchbar").contains("button", "Fill").click();
    cy.get(".sheetedit-grip-br").should("have.class", "fill");
    cy.get(".sheetedit-grip-tl").should("not.exist");
    cy.document().then((doc) => {
      const grip = doc.querySelector(".sheetedit-grip-br")!;
      touch(doc, "pointerdown", grip);
      const b3 = td(doc, "B3").getBoundingClientRect();
      const at = { x: b3.left + b3.width / 2, y: b3.top + b3.height / 2 };
      touch(doc, "pointermove", grip, at);
      touch(doc, "pointerup", grip, at);
    });
    cy.get('input[aria-label="B3"]').should("have.value", "3"); // B2's value, copied down
    cy.get('input[aria-label="C3"]').should("have.value", "6"); // and C3 = B3*2 follows
    cy.get(".sheetedit-grip-br").should("not.have.class", "fill");
  });
});
