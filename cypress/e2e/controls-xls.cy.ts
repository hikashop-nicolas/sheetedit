/// <reference types="cypress" />

// Form controls in a legacy .xls: they drive their linked cells, save as BIFF8 and reopen as saved.
const TIMEOUT = 20000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

/** Catch the bytes the demo hands to a download, so they can be opened again. */
function captureExport() {
  cy.window().then((win) => {
    (win as unknown as { __exported: Uint8Array | null }).__exported = null;
    const orig = win.URL.createObjectURL.bind(win.URL);
    win.URL.createObjectURL = (b: Blob) => {
      if (b instanceof win.Blob)
        void b.arrayBuffer().then((ab) => ((win as unknown as { __exported: Uint8Array }).__exported = new Uint8Array(ab)));
      return orig(b);
    };
  });
}

function reopenExport() {
  cy.window().its("__exported").should("exist");
  cy.window().then((win) => {
    const bytes = (win as unknown as { __exported: Uint8Array }).__exported;
    const file = new win.File([bytes as BlobPart], "controls.xls", { type: "application/vnd.ms-excel" });
    const dt = new win.DataTransfer();
    dt.items.add(file);
    const inp = win.document.getElementById("file") as HTMLInputElement;
    inp.files = dt.files;
    inp.dispatchEvent(new win.Event("change", { bubbles: true }));
  });
}

describe("form controls in a .xls", () => {
  it("toggles a check box, picks from a drop-down, saves and reopens", () => {
    open("cypress/fixtures/controls.xls");
    cy.get(".sheetedit-ctrl-check input[type=checkbox]", { timeout: TIMEOUT }).should("not.be.checked");
    cy.get('[title="Form control"]').should("exist");
    captureExport();
    cy.get(".sheetedit-ctrl-check input[type=checkbox]").check();
    cy.get('input[aria-label="F2"]').should("have.value", "TRUE");
    // Pick the second item the way the browser reports a choice: the value, then a change event.
    cy.get(".sheetedit-ctrl-select").should("have.length", 1).then(($s) => {
      ($s[0] as HTMLSelectElement).value = "2";
      $s[0]!.dispatchEvent(new Event("change", { bubbles: true }));
    });
    cy.get('input[aria-label="F3"]').should("have.value", "2");
    cy.get("#save").click();
    reopenExport();
    cy.get(".sheetedit-ctrl-check input[type=checkbox]", { timeout: TIMEOUT }).should("be.checked");
    cy.get(".sheetedit-ctrl-select").first().should("have.value", "2");
    cy.get('input[aria-label="F2"]').should("have.value", "TRUE");
    cy.get('input[aria-label="F3"]').should("have.value", "2");
  });
});
