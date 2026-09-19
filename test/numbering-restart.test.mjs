import { test } from "node:test";
import assert from "node:assert/strict";
import { openDocx, scan, fix } from "../src/index.js";
import { findParagraphs, attrValue, topLevelElements, elementInner } from "../src/xml.js";
import { buildMinimalDocx } from "../fixtures/minimal-docx.mjs";
import { buildNumberingRestartDocx } from "../fixtures/numbering-restart-docx.mjs";

const RULE_OPTS = { rules: ["numbering-restart"] };

function numIdOf(paraXml) {
  const paraInner = elementInner(paraXml);
  const pPr = topLevelElements(paraInner).find((el) => el.name === "w:pPr");
  const pPrInner = elementInner(pPr.xml);
  const numPr = topLevelElements(pPrInner).find((el) => el.name === "w:numPr");
  const numIdEl = topLevelElements(elementInner(numPr.xml)).find((el) => el.name === "w:numId");
  return attrValue(numIdEl.xml, "w:val");
}

test("detect: finds no findings on a clean document (no false positives)", async () => {
  // minimal-docx.mjs's 2-item list uses a single numId throughout.
  const doc = await openDocx(buildMinimalDocx());
  const { findings } = scan(doc, RULE_OPTS);
  assert.deepEqual(findings, []);
});

test("detect: flags the numId-2 block (same abstractNumId as the run's canonical numId 1), nothing else", async () => {
  const doc = await openDocx(buildNumberingRestartDocx());
  const { findings } = scan(doc, RULE_OPTS);

  assert.equal(findings.length, 1);
  const [finding] = findings;
  assert.equal(finding.rule, "numbering-restart");
  assert.equal(finding.severity, "error");
  assert.equal(finding.autofixable, true);
  // paragraphs: 0 heading, 1 body, 2-3 numId1, 4-5 numId2 (the restart), 6 numId3, 7 plain, 8 numId1 (lone)
  assert.deepEqual(finding.location, {
    paraIndexStart: 4,
    paraIndexEnd: 5,
    canonicalNumId: "1",
    restartedNumId: "2",
  });
});

test("fix: retargets only the restarted block to the canonical numId, leaves the different list type and the lone item alone", async () => {
  const doc = await openDocx(buildNumberingRestartDocx());
  const { applied, skipped } = fix(doc, RULE_OPTS);
  assert.equal(applied.length, 1);
  assert.deepEqual(skipped, []);

  const paragraphs = findParagraphs(doc.text("word/document.xml"));
  // 2, 3, 4, 5 should now all read numId 1 (the restarted block retargeted).
  assert.equal(numIdOf(paragraphs[2].xml), "1");
  assert.equal(numIdOf(paragraphs[3].xml), "1");
  assert.equal(numIdOf(paragraphs[4].xml), "1");
  assert.equal(numIdOf(paragraphs[5].xml), "1");
  // paragraph 6 (numId 3, a genuinely different list) must be untouched.
  assert.equal(numIdOf(paragraphs[6].xml), "3");
  // paragraph 8 (lone numId-1 item after the plain-paragraph break) untouched.
  assert.equal(numIdOf(paragraphs[8].xml), "1");

  // A second scan of the fixed document should be clean.
  const { findings: afterFix } = scan(doc, RULE_OPTS);
  assert.deepEqual(afterFix, []);
});
