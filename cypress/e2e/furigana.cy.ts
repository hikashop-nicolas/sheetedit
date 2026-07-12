/// <reference types="cypress" />

// furigana.xlsx: A1 is a shared string carrying a phonetic (rPh) run ("東京" + reading
// "トウキョウ"); B1 is a plain string. The furigana must render as ruby over the cell, with
// the cell value staying the base text, while a plain cell gets no ruby.

const TIMEOUT = 15000;

function open(fixture: string) {
  cy.visit("/");
  cy.get("#file").selectFile(fixture, { force: true });
  cy.get(".sheetedit-table", { timeout: TIMEOUT }).should("exist");
}

describe("furigana (phonetic ruby)", () => {
  it("renders the reading as ruby while the cell value stays the base text", () => {
    open("cypress/fixtures/furigana.xlsx");
    // The editable value is the base text only (not "東京トウキョウ").
    cy.get('input[aria-label="A1"]').should("have.value", "東京");
    // The reading is shown as ruby over the cell.
    cy.get('input[aria-label="A1"]').closest("td").as("cellA1");
    cy.get("@cellA1").should("have.class", "has-ruby");
    cy.get("@cellA1").find("ruby").should("contain.text", "東京");
    cy.get("@cellA1").find("ruby rt").should("have.text", "トウキョウ");
    // A plain string cell gets no ruby.
    cy.get('input[aria-label="B1"]').should("have.value", "plain");
    cy.get('input[aria-label="B1"]').closest("td").should("not.have.class", "has-ruby");
  });

  it("shows the plain base text for editing when the cell is focused", () => {
    open("cypress/fixtures/furigana.xlsx");
    cy.get('input[aria-label="A1"]').focus();
    // On focus the ruby overlay is hidden (the input is editable with the base text).
    cy.get('input[aria-label="A1"]').closest("td").find(".sheetedit-ruby").should("not.be.visible");
    cy.get('input[aria-label="A1"]').should("have.value", "東京");
  });

  it("adds and removes furigana on a cell via the toolbar", () => {
    cy.viewport(1400, 800); // wide enough that the style cluster stays inline
    open("cypress/fixtures/sample.xlsx"); // A1 = "item" (a text cell)
    cy.get('input[aria-label="A1"]').click(); // select the cell
    cy.contains(".sheetedit-toolbar button", "ふ").click();
    cy.get(".sheetedit-furi-input").should("be.visible").type("アイテム");
    cy.contains(".sheetedit-furi-pop button", "Set").click();
    // The reading now renders as ruby over the cell.
    cy.get('input[aria-label="A1"]').closest("td").as("a1");
    cy.get("@a1").should("have.class", "has-ruby");
    cy.get("@a1").find("ruby rt").should("have.text", "アイテム");
    cy.get('input[aria-label="A1"]').should("have.value", "item"); // value unchanged
    // Remove it again.
    cy.contains(".sheetedit-toolbar button", "ふ").click();
    cy.contains(".sheetedit-furi-pop button", "Remove").click();
    cy.get('input[aria-label="A1"]').closest("td").should("not.have.class", "has-ruby");
  });

  it("renders ODF text:ruby the same way", () => {
    open("cypress/fixtures/furigana.ods");
    cy.get('input[aria-label="A1"]').should("have.value", "東京");
    cy.get('input[aria-label="A1"]').closest("td").as("cellA1");
    cy.get("@cellA1").should("have.class", "has-ruby");
    cy.get("@cellA1").find("ruby rt").should("have.text", "トウキョウ");
    cy.get('input[aria-label="B1"]').closest("td").should("not.have.class", "has-ruby");
  });
});
