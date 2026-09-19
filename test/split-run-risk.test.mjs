import { test } from "node:test";
import assert from "node:assert/strict";
import { openDocx, scan, fix } from "../src/index.js";
import { buildMinimalDocx } from "../fixtures/minimal-docx.mjs";
import { buildSplitRunRiskDocx } from "../fixtures/split-run-risk-docx.mjs";

const RULE_OPTS = { rules: ["split-run-risk"] };

test("detect: finds no findings on a clean document (no false positives)", async () => {
  const doc = await openDocx(buildMinimalDocx());
  const { findings } = scan(doc, RULE_OPTS);
  assert.deepEqual(findings, []);
});

test("detect: flags the two split tokens, not the intact one or the token-free paragraphs", async () => {
  const doc = await openDocx(buildSplitRunRiskDocx());
  const { findings } = scan(doc, RULE_OPTS);

  assert.equal(findings.length, 2);
  const byIndex = Object.fromEntries(findings.map((f) => [f.location.paraIndex, f]));

  assert.ok(byIndex[1], "the {{PROJECT_NAME}} split across 2 runs should be flagged");
  assert.equal(byIndex[1].location.token, "{{PROJECT_NAME}}");
  assert.equal(byIndex[1].location.runsTouched, 2);
  assert.equal(byIndex[1].severity, "warn");
  assert.equal(byIndex[1].autofixable, false);

  assert.ok(byIndex[3], "the ${CLIENT_NAME} split across 3 runs should be flagged");
  assert.equal(byIndex[3].location.token, "${CLIENT_NAME}");
  assert.equal(byIndex[3].location.runsTouched, 3);

  assert.equal(byIndex[2], undefined, "the intact token (paragraph 2) must not be flagged");
  assert.equal(byIndex[0], undefined, "the heading (no tokens) must not be flagged");
  assert.equal(byIndex[4], undefined, "the token-free paragraph must not be flagged");
});

test("fix: never touches split-run-risk findings — detect-only by design", async () => {
  const doc = await openDocx(buildSplitRunRiskDocx());
  const before = doc.text("word/document.xml");

  const { applied, skipped } = fix(doc, RULE_OPTS);
  assert.deepEqual(applied, []);
  assert.equal(skipped.length, 2);

  assert.equal(doc.text("word/document.xml"), before, "fix() must not mutate anything for this rule");
});
