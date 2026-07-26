// Turns src/sheetedit.css into src/core/ui/styles.generated.ts.
//
// The package ships through plain tsc (no bundler), which cannot import or emit a .css file, and
// the widget injects its own <style> tag so a host needs no build step at all. So the stylesheet is
// authored as real CSS and compiled into a string module here. The generated file is committed so
// a fresh clone, the test run and the Vite dev server all work without running this first; the
// build and prepare scripts regenerate it, and `npm run css:check` fails if it is stale.
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const SRC = join(root, "src/sheetedit.css");
const OUT = join(root, "src/core/ui/styles.generated.ts");

/** Strip comments and collapse whitespace: the shipped string is never read by a human. */
function minify(css) {
  return css
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .replace(/\s*\n\s*/g, "\n")
    .replace(/\n{2,}/g, "\n")
    .trim();
}

const css = readFileSync(SRC, "utf8");
const body = minify(css)
  .replace(/\\/g, "\\\\")
  .replace(/`/g, "\\`")
  .replace(/\$\{/g, "\\${");

const out = `// GENERATED FROM src/sheetedit.css - DO NOT EDIT. Run \`npm run css\` after changing the CSS.
export const SHEETEDIT_CSS = \`${body}\`;
`;

const check = process.argv.includes("--check");
if (check) {
  let current = "";
  try {
    current = readFileSync(OUT, "utf8");
  } catch {
    /* missing counts as stale */
  }
  if (current !== out) {
    console.error("styles.generated.ts is out of date with sheetedit.css - run `npm run css`.");
    process.exit(1);
  }
  console.log("styles.generated.ts is up to date.");
} else {
  writeFileSync(OUT, out);
  console.log(`styles.generated.ts written (${body.length} chars).`);
}
