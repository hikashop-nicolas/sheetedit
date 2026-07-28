#!/usr/bin/env node
// Run the openpyxl cross-check with whichever Python has openpyxl.
//
// CI installs it into the system interpreter; a local checkout more likely wants a virtual
// environment, since a Homebrew or Debian python refuses to install into itself. This prefers
// .cache/py (which `npm run check:openpyxl:setup` creates) and falls back to python3, so the
// command works the same either way instead of failing with an import error and no advice.

import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { join } from "node:path";

const ROOT = new URL("..", import.meta.url).pathname;
const venv = join(ROOT, ".cache", "py", "bin", "python");
const python = existsSync(venv) ? venv : "python3";

try {
  execFileSync(python, [join(ROOT, "scripts", "openpyxl-check.py")], { stdio: "inherit" });
} catch (e) {
  if (e.status === 2 && python === "python3") {
    console.error("\nopenpyxl is not installed. Either `pip install openpyxl`, or run");
    console.error("`npm run check:openpyxl:setup` to put it in a virtual environment under .cache/.");
  }
  process.exit(e.status ?? 1);
}
