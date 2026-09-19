import { test } from "node:test";
import assert from "node:assert/strict";
import { openDocx, scan, fix } from "../src/index.js";
import { findParagraphs } from "../src/xml.js";
import { buildMinimalDocx } from "../fixtures/minimal-docx.mjs";
import { buildHeadingPPrOrderDocx } from "../fixtures/heading-ppr-order-docx.mjs";

const RULE_OPTS = { rules: ["heading-ppr-order"] };

test("detect: finds no findings on a clean document (no false positives)", async () => {
  const doc = await openDocx(buildMinimalDocx());
  const { findings } = scan(doc, RULE_OPTS);
  assert.deepEqual(findings, []);
});

test("detect: flags only the two broken paragraphs, not the correctly-ordered or rPr-less ones", async () => {
  const doc = await openDocx(buildHeadingPPrOrderDocx());
  const { findings } = scan(doc, RULE_OPTS);

  assert.equal(findings.length, 2);
  const byIndex = Object.fromEntries(findings.map((f) => [f.location.paraIndex, f]));

  assert.ok(byIndex[0], "paragraph 0 (rPr + 1 trailing element) should be flagged");
  assert.deepEqual(byIndex[0].location.afterRPr, ["w:spacing"]);
  assert.equal(byIndex[0].severity, "error");
  assert.equal(byIndex[0].autofixable, true);

  assert.ok(byIndex[4], "paragraph 4 (rPr + 2 trailing elements) should be flagged");
  assert.deepEqual(byIndex[4].location.afterRPr, ["w:ind", "w:spacing"]);

  assert.equal(byIndex[1], undefined, "correctly-ordered paragraph must not be flagged");
  assert.equal(byIndex[2], undefined, "paragraph with no rPr must not be flagged");
  assert.equal(byIndex[3], undefined, "paragraph with no pPr must not be flagged");
});

test("fix: reorders rPr to last, preserves the other elements' relative order, leaves untouched paragraphs untouched", async () => {
  const doc = await openDocx(buildHeadingPPrOrderDocx());
  const before = findParagraphs(doc.text("word/document.xml"));

  const { applied, skipped } = fix(doc, RULE_OPTS);
  assert.equal(applied.length, 2);
  assert.deepEqual(skipped, []);

  const after = findParagraphs(doc.text("word/document.xml"));

  // Paragraph 0: was rPr, spacing → now spacing, rPr.
  assert.equal(
    after[0].xml,
    "<w:p><w:pPr><w:spacing w:before=\"240\" w:after=\"120\"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Broken heading spacing</w:t></w:r></w:p>"
  );

  // Paragraph 4: was rPr, ind, spacing → now ind, spacing, rPr (order preserved, not reversed).
  assert.equal(
    after[4].xml,
    "<w:p><w:pPr><w:ind w:left=\"720\"/><w:spacing w:before=\"240\" w:after=\"120\"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Broken with two trailing elements</w:t></w:r></w:p>"
  );

  // Untouched paragraphs (1, 2, 3) must be byte-for-byte unchanged.
  for (const i of [1, 2, 3]) {
    assert.equal(after[i].xml, before[i].xml, `paragraph ${i} should be unchanged`);
  }

  // A second scan of the fixed document should be clean.
  const { findings: afterFix } = scan(doc, RULE_OPTS);
  assert.deepEqual(afterFix, []);
});
