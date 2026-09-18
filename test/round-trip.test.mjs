import { test } from "node:test";
import assert from "node:assert/strict";
import { openDocx } from "../src/index.js";
import { buildMinimalDocx, MINIMAL_DOCX_PARTS } from "../fixtures/minimal-docx.mjs";

// The most important test in the project (see ROADMAP.md Weekend 1): open
// a .docx and save it straight back out must reproduce every part
// byte-for-byte. Every later rule's repair() depends on this holding —
// if open/save themselves mutate anything, "fixed" output can't be trusted.

test("openDocx reads every part the fixture was built from", async () => {
  const doc = await openDocx(buildMinimalDocx());
  assert.deepEqual(
    [...doc.parts.keys()].sort(),
    Object.keys(MINIMAL_DOCX_PARTS).sort()
  );
  assert.equal(doc.text("word/document.xml"), MINIMAL_DOCX_PARTS["word/document.xml"]);
});

test("round-trip: save() then re-open reproduces every part unchanged", async () => {
  const doc = await openDocx(buildMinimalDocx());
  const savedBuffer = doc.toBuffer();
  const reopened = await openDocx(savedBuffer);

  assert.equal(reopened.parts.size, doc.parts.size);
  for (const [path, originalBytes] of doc.parts) {
    const roundTripped = reopened.parts.get(path);
    assert.ok(roundTripped, `part missing after round-trip: ${path}`);
    assert.equal(
      Buffer.from(roundTripped).toString("utf8"),
      Buffer.from(originalBytes).toString("utf8"),
      `part changed after round-trip: ${path}`
    );
  }
});

test("round-trip: saved buffer is itself a well-formed zip", async () => {
  const doc = await openDocx(buildMinimalDocx());
  const buf = doc.toBuffer();
  assert.ok(buf.length > 0);
  const reopened = await openDocx(buf);
  assert.ok(reopened.text("word/document.xml").includes("This is a heading"));
});
