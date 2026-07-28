#!/usr/bin/env node
// Validate what sheetedit WRITES against the official ECMA-376 schemas.
//
// The point is not "is this file schema-perfect": real workbooks are not. Excel writes
// xml:space on <t>, which the schema does not allow, and it writes mc:AlternateContent in
// worksheets, which the worksheet schema does not mention. A validator run against any real
// file reports those, and they are not defects.
//
// So this compares. Every demo workbook is read, edited and written back, and each XML part of
// the OUTPUT is validated against the same part of the INPUT. Only an error the input did not
// have is reported: that is one this project introduced. Everything else is the schema being
// stricter than the format's practice, which is not ours to fix.
//
// The schemas come from ECMA-376 Part 4 (transitional, which is what Excel writes and what
// sheetedit preserves). They are downloaded once into .cache/ and not committed: they are
// ECMA's, and a checkout should not carry a copy of someone else's standard.
//
// Needs xmllint (libxml2), which macOS carries and Debian/Ubuntu package as libxml2-utils.

import { execFileSync, execSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync, readdirSync, statSync } from "node:fs";
import { join, relative } from "node:path";
import { tmpdir } from "node:os";

const ROOT = new URL("..", import.meta.url).pathname;
const CACHE = join(ROOT, ".cache", "ooxml-xsd");
const PART4 = "https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip";

function have(cmd) {
  try { execSync(`command -v ${cmd}`, { stdio: "ignore" }); return true; } catch { return false; }
}

/** Fetch and unpack the transitional schemas, once. */
function ensureSchemas() {
  if (existsSync(join(CACHE, "sml.xsd"))) return;
  console.log("fetching the ECMA-376 transitional schemas (once, into .cache/) ...");
  mkdirSync(CACHE, { recursive: true });
  const work = join(tmpdir(), `ooxml-xsd-${process.pid}`);
  mkdirSync(work, { recursive: true });
  execFileSync("curl", ["-sSL", "-o", join(work, "part4.zip"), PART4]);
  execFileSync("unzip", ["-o", "-q", join(work, "part4.zip"), "OfficeOpenXML-XMLSchema-Transitional.zip", "-d", work]);
  execFileSync("unzip", ["-o", "-q", join(work, "OfficeOpenXML-XMLSchema-Transitional.zip"), "-d", CACHE]);
  rmSync(work, { recursive: true, force: true });
}

/** The schema a part is described by, or null for parts this check does not cover. */
function schemaFor(part) {
  if (/xl\/(worksheets\/sheet[^/]*\.xml|workbook\.xml|styles\.xml|sharedStrings\.xml)$/.test(part)) return "sml.xsd";
  if (/xl\/(tables|pivotTables|pivotCache|queryTables)\/[^/]*\.xml$/.test(part)) return "sml.xsd";
  if (/xl\/connections\.xml$/.test(part)) return "sml.xsd";
  return null; // drawings, charts, slicers, timelines: their own schemas or extensions, not here
}

/** Validate one part; returns the error lines, with the file name and line numbers stripped so
    the same complaint from two files compares equal. */
function validate(file, schema) {
  let out = "";
  try {
    out = execFileSync("xmllint", ["--noout", "--schema", join(CACHE, schema), file], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });
  } catch (e) {
    out = `${e.stdout ?? ""}${e.stderr ?? ""}`;
  }
  return out
    .split("\n")
    .filter((l) => l.includes("Schemas validity error"))
    .map((l) => l.replace(/^[^:]*:\d+:\s*/, "").trim());
}

function unzipTo(zip, dir) {
  mkdirSync(dir, { recursive: true });
  execFileSync("unzip", ["-o", "-q", zip, "-d", dir]);
  const out = [];
  const walk = (d) => {
    for (const name of readdirSync(d)) {
      const p = join(d, name);
      if (statSync(p).isDirectory()) walk(p);
      else if (p.endsWith(".xml")) out.push(p);
    }
  };
  walk(dir);
  return out;
}

if (!have("xmllint")) {
  console.error("xmllint is needed (macOS has it; Debian/Ubuntu: apt-get install libxml2-utils)");
  process.exit(2);
}
ensureSchemas();

// With no arguments, check every workbook the corpus test wrote (see
// src/adapters/xlsx/schema-corpus.test.ts, which `npm run check:schema` runs first) against the
// demo fixture it came from.
let inputs = process.argv.slice(2);
if (!inputs.length) {
  const corpus = join(ROOT, ".cache", "schema-corpus");
  if (!existsSync(corpus)) {
    console.error("no corpus: run `npm run check:schema`, which writes it first");
    process.exit(2);
  }
  inputs = readdirSync(corpus)
    .filter((n) => n.endsWith(".xlsx"))
    .sort()
    .flatMap((n) => [join(ROOT, "demo", n), join(corpus, n)])
    .filter((p, i, a) => (i % 2 === 1 ? true : existsSync(p) && existsSync(a[i + 1])));
  console.log(`checking ${inputs.length / 2} workbooks ...`);
}

let introduced = 0;
for (let i = 0; i < inputs.length; i += 2) {
  const [before, after] = [inputs[i], inputs[i + 1]];
  const work = join(tmpdir(), `schemacheck-${process.pid}-${i}`);
  const beforeParts = unzipTo(before, join(work, "in"));
  const afterParts = unzipTo(after, join(work, "out"));
  // The baseline is every complaint the INPUT draws, from any of its parts. A workbook that
  // already writes xml:space on a <t> in its shared strings is a workbook where that complaint
  // says nothing about us, wherever it then turns up in the output. What is worth reporting is a
  // KIND of violation the file did not have before.
  const baseline = new Set();
  for (const p of beforeParts) {
    const rel = relative(join(work, "in"), p);
    const schema = schemaFor(rel);
    if (schema) for (const e of validate(p, schema)) baseline.add(e);
  }
  for (const p of afterParts) {
    const rel = relative(join(work, "out"), p);
    const schema = schemaFor(rel);
    if (!schema) continue;
    const now = validate(p, schema);
    const fresh = now.filter((e) => !baseline.has(e));
    for (const e of fresh) {
      introduced++;
      console.log(`${before.split("/").pop()} -> ${rel}\n  ${e}`);
    }
  }
  rmSync(work, { recursive: true, force: true });
}

if (introduced) {
  console.error(`\n${introduced} schema violation(s) introduced by writing.`);
  process.exit(1);
}
console.log("no schema violations introduced by writing.");
