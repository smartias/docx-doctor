import { test } from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { writeFileSync, mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { openDocx, scan } from "../src/index.js";
import { buildTrailingBlanksDocx } from "../fixtures/trailing-blanks-docx.mjs";
import { buildHeadingPPrOrderDocx } from "../fixtures/heading-ppr-order-docx.mjs";

const CLI_PATH = fileURLToPath(new URL("../src/cli.js", import.meta.url));

function tmpFile(name, bytes) {
  const dir = mkdtempSync(join(tmpdir(), "docx-doctor-cli-test-"));
  const path = join(dir, name);
  writeFileSync(path, bytes);
  return path;
}

function runCli(args) {
  try {
    const stdout = execFileSync(process.execPath, [CLI_PATH, ...args], { encoding: "utf8" });
    return { status: 0, stdout };
  } catch (err) {
    // execFileSync throws on nonzero exit; recover status/stdout from it.
    return { status: err.status, stdout: err.stdout?.toString() ?? "" };
  }
}

test("scan: reports findings and exits 0 without --ci even when findings exist (severity is warn)", () => {
  const file = tmpFile("trailing-blanks.docx", buildTrailingBlanksDocx());
  const { status, stdout } = runCli(["scan", file, "--rules", "trailing-blank-pages"]);
  assert.equal(status, 0);
  assert.match(stdout, /trailing-blank-pages/);
  assert.match(stdout, /1 finding/);
});

test("scan --ci: exits nonzero when an error-severity finding is present", () => {
  const file = tmpFile("heading-ppr.docx", buildHeadingPPrOrderDocx());
  const { status, stdout } = runCli(["scan", file, "--rules", "heading-ppr-order", "--ci"]);
  assert.equal(status, 1);
  assert.match(stdout, /heading-ppr-order/);
});

test("scan --ci: exits 0 when findings are warn-only, even with --ci", () => {
  const file = tmpFile("trailing-blanks.docx", buildTrailingBlanksDocx());
  const { status } = runCli(["scan", file, "--rules", "trailing-blank-pages", "--ci"]);
  assert.equal(status, 0);
});

test("scan: reports no findings on a rule-scoped clean run", () => {
  const file = tmpFile("heading-ppr.docx", buildHeadingPPrOrderDocx());
  // This fixture has no trailing-blank-pages defect, only heading-pPr-order ones.
  const { status, stdout } = runCli(["scan", file, "--rules", "trailing-blank-pages"]);
  assert.equal(status, 0);
  assert.match(stdout, /No findings\./);
});

test("fix: writes a repaired file that re-scans clean", () => {
  const inFile = tmpFile("trailing-blanks.docx", buildTrailingBlanksDocx());
  const outFile = inFile.replace(".docx", ".fixed.docx");
  const { status, stdout } = runCli(["fix", inFile, "-o", outFile, "--rules", "trailing-blank-pages"]);
  assert.equal(status, 0);
  assert.match(stdout, /Applied 1 fix/);
  assert.match(stdout, new RegExp(`Wrote ${outFile.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`));

  const fixedBytes = readFileSync(outFile);
  return openDocx(fixedBytes).then(async (doc) => {
    const { findings } = scan(doc, { rules: ["trailing-blank-pages"] });
    assert.deepEqual(findings, []);
  });
});

test("fix: requires -o", () => {
  const file = tmpFile("trailing-blanks.docx", buildTrailingBlanksDocx());
  const { status } = runCli(["fix", file]);
  assert.equal(status, 1);
});

test("no command: prints usage and exits 1", () => {
  const { status } = runCli([]);
  assert.equal(status, 1);
});
