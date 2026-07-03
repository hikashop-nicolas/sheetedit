/// <reference types="cypress" />

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("sheetedit", () => {
  it("renders an .xlsx as a grid with computed formula values", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('.sheetedit-tab[aria-selected="true"]').should("contain.text", "Budget");
    cy.get('input[aria-label="A1"]').should("have.value", "item");
    cy.get('input[aria-label="C2"]').should("have.value", "6"); // B2*2
    cy.get('input[aria-label="C4"]').should("have.value", "14"); // SUM(C2:C3)
  });

  it("shows formatted numbers but edits the raw value", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="D2"]').should("have.value", "$3.50"); // currency format applied
    cy.get('input[aria-label="D2"]').focus().should("have.value", "3.5"); // raw value when editing
    cy.get('input[aria-label="D2"]').clear().type("9.9").blur();
    cy.get('input[aria-label="D2"]').should("have.value", "$9.90"); // typed value keeps the format
  });

  it("recalculates formulas when a dependency changes", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="B2"]').clear().type("5").blur();
    cy.get('input[aria-label="C2"]').should("have.value", "10"); // 5*2
    cy.get('input[aria-label="C4"]').should("have.value", "18"); // 10 + 8
  });

  it("shows the formula when a formula cell is focused", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="C2"]').focus().should("have.value", "=B2*2");
  });

  it("navigates cells with arrow keys", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A1"]').focus();
    cy.focused().type("{downArrow}");
    cy.focused().should("have.attr", "aria-label", "A2");
    cy.focused().type("{upArrow}");
    cy.focused().should("have.attr", "aria-label", "A1");
  });

  // Note: TSV paste is verified manually in a real browser; Cypress's synthetic
  // ClipboardEvent does not reach the page's paste listener with usable data.

  it("clears a multi-cell selection with Delete", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A2"]').focus(); // anchor
    cy.get('input[aria-label="B3"]').trigger("mousedown", { shiftKey: true }); // extend A2:B3
    cy.focused().trigger("keydown", { key: "Delete" });
    cy.get('input[aria-label="A2"]').should("have.value", "");
    cy.get('input[aria-label="B2"]').should("have.value", "");
    cy.get('input[aria-label="A3"]').should("have.value", "");
    cy.get('input[aria-label="B3"]').should("have.value", "");
  });

  it("inserts =SUM(range) after a selected column run", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="B2"]').focus(); // anchor
    cy.get('input[aria-label="B3"]').trigger("mousedown", { shiftKey: true }); // extend B2:B3
    cy.get(".sheetedit-fxsum").click();
    // The result cell is focused for review, showing the editable formula.
    cy.focused().should("have.attr", "aria-label", "B4");
    cy.focused().should("have.value", "=SUM(B2:B3)");
    cy.focused().blur();
    cy.get('input[aria-label="B4"]').should("have.value", "7"); // 3 + 4
  });

  it("picks a range for a pending formula while editing", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="E1"]').focus();
    cy.get(".sheetedit-fxsum").click(); // editing a single cell: enter pick mode
    cy.get(".sheetedit-fxbar").should("have.class", "is-picking");
    cy.get('input[aria-label="B2"]').closest("td").trigger("pointerdown", { pointerType: "mouse", buttons: 1 });
    cy.get('input[aria-label="B3"]').closest("td").trigger("pointermove", { pointerType: "mouse", buttons: 1 });
    cy.window().then((win) => win.dispatchEvent(new win.PointerEvent("pointerup", { pointerType: "mouse" })));
    cy.get('input[aria-label="E1"]').should("have.value", "=SUM(B2:B3)");
    cy.get(".sheetedit-fxbar").should("not.have.class", "is-picking");
    cy.focused().type("{enter}");
    cy.get('input[aria-label="E1"]').should("have.value", "7");
  });

  it("edits through the formula bar", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="C2"]').focus();
    cy.get(".sheetedit-fxref").should("have.text", "C2");
    cy.get(".sheetedit-fxinput").should("have.value", "=B2*2");
    cy.get(".sheetedit-fxinput").clear().type("=B2*3{enter}");
    cy.get('input[aria-label="C2"]').should("have.value", "9");
  });

  it("collapses the style cluster into a group button when narrow", () => {
    cy.viewport(340, 600);
    open("cypress/fixtures/sample.xlsx");
    cy.get(".sheetedit-tb-slot").contains("button", "Aa"); // collapsed to the group button
    cy.get(".sheetedit-tb-slot button").contains("Aa").click();
    cy.get(".sheetedit-tb-groupmenu").should("be.visible").find("button").should("have.length.greaterThan", 4);
  });

  it("undoes and redoes a cell edit", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="B2"]').clear().type("42").blur();
    cy.get('input[aria-label="C2"]').should("have.value", "84"); // B2*2 recalculated
    cy.get('input[aria-label="A1"]').focus().type("{cmd+z}");
    cy.get('input[aria-label="B2"]').should("have.value", "3"); // original restored
    cy.get('input[aria-label="C2"]').should("have.value", "6");
    cy.get('input[aria-label="A1"]').focus().type("{cmd+shift+z}");
    cy.get('input[aria-label="B2"]').should("have.value", "42");
    cy.get('input[aria-label="C2"]').should("have.value", "84");
  });

  it("undoes a range clear via the toolbar button", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="A2"]').focus(); // anchor
    cy.get('input[aria-label="B3"]').trigger("mousedown", { shiftKey: true }); // extend A2:B3
    cy.focused().trigger("keydown", { key: "Delete" });
    cy.get('input[aria-label="B2"]').should("have.value", "");
    cy.get('button[aria-label="Undo (Ctrl+Z)"]').click();
    cy.get('input[aria-label="B2"]').should("have.value", "3");
    cy.get('input[aria-label="A2"]').should("have.value", "apples");
  });

  it("Escape cancels an edit without committing the display text", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.get('input[aria-label="C2"]').focus().clear().type("=B2*99").type("{esc}");
    cy.get('input[aria-label="C2"]').should("have.value", "6"); // display restored
    cy.get('input[aria-label="C2"]').focus().should("have.value", "=B2*2"); // formula intact
    // Escape on a formatted cell must not commit the "$3.50" display as a string.
    cy.get('input[aria-label="D2"]').focus().type("{esc}");
    cy.get('input[aria-label="D2"]').should("have.value", "$3.50");
    cy.get('input[aria-label="D2"]').focus().should("have.value", "3.5");
  });

  it("edits, exports a valid .xlsx, and round-trips with recalculated values", () => {
    open("cypress/fixtures/sample.xlsx");
    cy.window().then((win) => {
      (win as unknown as { __exported: Uint8Array | null }).__exported = null;
      const orig = win.URL.createObjectURL.bind(win.URL);
      win.URL.createObjectURL = (b: Blob) => {
        if (b instanceof win.Blob)
          void b.arrayBuffer().then((ab) => ((win as unknown as { __exported: Uint8Array }).__exported = new Uint8Array(ab)));
        return orig(b);
      };
    });
    cy.get('input[aria-label="B2"]').clear().type("5").blur();
    cy.get("#save").click();
    cy.window().its("__exported").should("exist");
    cy.window().then((win) => {
      const bytes = (win as unknown as { __exported: Uint8Array }).__exported;
      const file = new win.File([bytes as BlobPart], "x.xlsx", {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const dt = new win.DataTransfer();
      dt.items.add(file);
      const inp = win.document.getElementById("file") as HTMLInputElement;
      inp.files = dt.files;
      inp.dispatchEvent(new win.Event("change", { bubbles: true }));
    });
    cy.get('input[aria-label="B2"]', { timeout: TIMEOUT }).should("have.value", "5");
    cy.get('input[aria-label="C2"]').should("have.value", "10");
    cy.get('input[aria-label="C4"]').should("have.value", "18");
  });

  it("renders and edits an .ods workbook", () => {
    open("cypress/fixtures/sample.ods");
    cy.get('input[aria-label="C2"]').should("have.value", "6"); // B2*2
    cy.get('input[aria-label="B2"]').clear().type("10").blur();
    cy.get('input[aria-label="C2"]').should("have.value", "20");
  });
});
