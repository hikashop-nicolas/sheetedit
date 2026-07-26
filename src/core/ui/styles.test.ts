import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { SHEETEDIT_CSS } from "./styles.generated";

// The stylesheet is authored as real CSS and compiled into a string module, because the package
// ships through plain tsc which cannot import a .css file. That leaves one failure mode: editing
// the CSS and forgetting to regenerate. This catches it in the normal test run.

describe("stylesheet", () => {
  it("the generated module is in step with sheetedit.css", () => {
    const css = readFileSync("src/sheetedit.css", "utf8");
    // Every selector in the source has to be present in what actually ships.
    const selectors = [...css.matchAll(/^([.#][\w-][^{\n]*)\{/gm)].map((m) => m[1]!.trim());
    expect(selectors.length).toBeGreaterThan(50);
    const missing = selectors.filter((s) => !SHEETEDIT_CSS.includes(s.replace(/\s+/g, " ")));
    expect(missing, "run `npm run css` after editing sheetedit.css").toEqual([]);
  });

  it("ships no comments, so the injected string stays small", () => {
    expect(SHEETEDIT_CSS).not.toContain("/*");
  });

  it("routes every colour through a custom property", () => {
    // A literal colour in a rule cannot be restyled by a host or a theme. Literals belong in the
    // token block (and in the code-editor palette, which defines its own --se-* set per scheme).
    const rules = SHEETEDIT_CSS.split("\n").filter((l) => !/^\s*--/.test(l));
    const bare = rules.join("\n").replace(/var\([^)]*\)/g, "").match(/#[0-9a-fA-F]{3,8}\b/g) ?? [];
    expect(bare, "define a --sheetedit-* token instead of a literal colour").toEqual([]);
  });

  it("defines a default for every token it references", () => {
    const defined = new Set([...SHEETEDIT_CSS.matchAll(/^\s*(--sheetedit-[\w-]+)\s*:/gm)].map((m) => m[1]!));
    const used = new Set([...SHEETEDIT_CSS.matchAll(/var\((--sheetedit-[\w-]+)/g)].map((m) => m[1]!));
    expect([...used].filter((v) => !defined.has(v)), "add it to the :root token block").toEqual([]);
  });
});
