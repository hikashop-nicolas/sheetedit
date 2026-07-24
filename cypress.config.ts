import { defineConfig } from "cypress";

export default defineConfig({
  e2e: {
    baseUrl: "http://localhost:5173",
    supportFile: "cypress/support/e2e.ts",
    video: false,
    defaultCommandTimeout: 8000,
    // Wide enough that the toolbar's authoring controls stay inline (they fold into a "⋯" overflow
    // menu on narrow widths); specs that test narrow layouts set their own cy.viewport().
    viewportWidth: 1600,
    viewportHeight: 900,
  },
});
