#!/usr/bin/env node
import { readFile, writeFile } from "node:fs/promises";
import { openDocx, scan, fix } from "./index.js";

const USAGE = `Usage:
  docx-doctor scan <file> [--rules a,b,c] [--ci]
  docx-doctor fix <file> -o <out> [--rules a,b,c]

  --ci     exit with a nonzero status if any error-severity finding is present
           (warnings alone don't fail the build) — for use in a CI pipeline.
  --rules  comma-separated rule ids to run; defaults to all built-in rules.
  -o, --out   output path for "fix" (required).`;

function parseArgs(argv) {
  const [command, file, ...rest] = argv;
  const opts = { ci: false, out: null, rules: null };
  for (let i = 0; i < rest.length; i++) {
    const arg = rest[i];
    if (arg === "--ci") opts.ci = true;
    else if (arg === "-o" || arg === "--out") opts.out = rest[++i];
    else if (arg === "--rules") opts.rules = rest[++i].split(",").map((s) => s.trim()).filter(Boolean);
  }
  return { command, file, opts };
}

function printFindings(findings) {
  for (const f of findings) {
    const tag = f.severity === "error" ? "ERROR" : "WARN ";
    console.log(`[${tag}] ${f.rule}: ${f.message}`);
  }
}

async function main() {
  const { command, file, opts } = parseArgs(process.argv.slice(2));
  if (!command || !file || !["scan", "fix"].includes(command)) {
    console.error(USAGE);
    process.exit(1);
  }

  const buffer = await readFile(file);
  const doc = await openDocx(buffer);
  const ruleOpts = opts.rules ? { rules: opts.rules } : {};

  if (command === "scan") {
    const { findings } = scan(doc, ruleOpts);
    if (findings.length === 0) {
      console.log("No findings.");
    } else {
      printFindings(findings);
      console.log(`\n${findings.length} finding(s).`);
    }
    const hasError = findings.some((f) => f.severity === "error");
    if (opts.ci && hasError) process.exit(1);
    return;
  }

  // command === "fix"
  if (!opts.out) {
    console.error("fix requires -o <output file>\n\n" + USAGE);
    process.exit(1);
  }
  const { applied, skipped } = fix(doc, ruleOpts);
  console.log(`Applied ${applied.length} fix(es).`);
  if (skipped.length > 0) {
    console.log(`Skipped ${skipped.length} finding(s) (not autofixable):`);
    printFindings(skipped);
  }
  await writeFile(opts.out, doc.toBuffer());
  console.log(`Wrote ${opts.out}`);
}

main().catch((err) => {
  console.error(err.stack || String(err));
  process.exit(1);
});
