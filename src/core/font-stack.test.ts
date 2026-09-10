import { describe, expect, it } from "vitest";
import { fontStack } from "./model";

// A bare family name is a trap: an uninstalled font drops the browser to its default standard face,
// usually a serif and always a different width, and the file's column widths were measured against
// the font it names. Substituting a full-width sans for a condensed one pushed text out of its cell.

describe("the font stack for a name the file gives", () => {
  it("puts the file's own font first", () => {
    expect(fontStack("Calibri")).toMatch(/^"Calibri", /);
  });

  it("falls back to a sans, never to the browser's default", () => {
    expect(fontStack("Calibri")).toContain("sans-serif");
  });

  it("keeps a condensed face condensed when it is not installed", () => {
    const narrow = fontStack("Aptos Narrow");
    expect(narrow).toMatch(/^"Aptos Narrow", /);
    expect(narrow).toContain("Arial Narrow");
    expect(fontStack("Helvetica Condensed")).toContain("Arial Narrow");
    expect(fontStack("Calibri")).not.toContain("Arial Narrow");
  });

  it("is empty for no font, so the grid keeps its own", () => {
    expect(fontStack(undefined)).toBe("");
    expect(fontStack("")).toBe("");
  });

  // The name goes into a style attribute, so a quote in it must not close the string early.
  it("cannot break out of the declaration", () => {
    expect(fontStack('Ev"il')).toBe('"Evil", ui-sans-serif, system-ui, sans-serif');
  });
});
