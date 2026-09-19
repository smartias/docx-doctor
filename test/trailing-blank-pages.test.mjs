import { test } from "node:test";
import assert from "node:assert/strict";
import { openDocx, scan, fix } from "../src/index.js";
import { buildMinimalDocx } from "../fixtures/minimal-docx.mjs";
import {
  buildTrailingBlanksDocx,
  TRAILING_BLANKS_REAL_CONTENT,
} from "../fixtures/trailing-blanks-docx.mjs";

const RULE_OPTS = { rules: ["trailing-blank-pages"] };

test("detect: finds no findings on a clean document (no false positives)", async () => {
  const doc = await openDocx(buildMinimalDocx());
  const { findings } = scan(doc, RULE_OPTS);
  assert.deepEqual(findings, []);
});

test("detect: finds the trailing run of empty paragraphs on a damaged document", async () => {
  const doc = await openDocx(buildTrailingBlanksDocx());
  const { findings } = scan(doc, RULE_OPTS);
  assert.equal(findings.length, 1);
  const [finding] = findings;
  assert.equal(finding.rule, "trailing-blank-pages");
  assert.equal(finding.severity, "warn");
  assert.equal(finding.autofixable, true);
  // heading(0), real paragraph(1), then 3 empties(2,3,4)
  assert.deepEqual(finding.location, { paraIndexStart: 2, paraIndexEnd: 4 });
});

test("fix: removes the trailing empties and leaves real content untouched", async () => {
  const doc = await openDocx(buildTrailingBlanksDocx());
  const { applied, skipped } = fix(doc, RULE_OPTS);
  assert.equal(applied.length, 1);
  assert.deepEqual(skipped, []);

  const xml = doc.text("word/document.xml");
  for (const text of TRAILING_BLANKS_REAL_CONTENT) {
    assert.ok(xml.includes(text), `expected surviving text: ${text}`);
  }
  // Exactly the two real paragraphs should remain — the fix shouldn't
  // leave a trailing empty behind on top of them, nor eat past them.
  const remainingParas = [...xml.matchAll(/<w:p(?:\s[^>]*)?>|<w:p(?:\s[^>]*)?\/>/g)];
  assert.equal(remainingParas.length, 2);

  // A second scan of the fixed document should be clean.
  const { findings: afterFix } = scan(doc, RULE_OPTS);
  assert.deepEqual(afterFix, []);
});

test("fix: never removes every paragraph, even if the whole body is empty", async () => {
  // Degenerate case: nothing but empty paragraphs before sectPr. The
  // safety rule must leave the first one behind rather than emptying the
  // body entirely.
  const { buildDocxFromDocument } = await import("../fixtures/parts.mjs");
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p/><w:p/><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
</w:document>`;
  const { buffer } = buildDocxFromDocument(documentXml);
  const doc = await openDocx(buffer);

  fix(doc, RULE_OPTS);

  const xml = doc.text("word/document.xml");
  const remainingParas = [...xml.matchAll(/<w:p(?:\s[^>]*)?>|<w:p(?:\s[^>]*)?\/>/g)];
  assert.equal(remainingParas.length, 1, "must leave at least one paragraph in the body");
});
