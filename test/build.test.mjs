import { test } from "node:test";
import assert from "node:assert/strict";
import { createDocument, paragraph, token, table, pageBreak } from "../src/build/index.js";
import { openDocx, scan } from "../src/index.js";
import { findParagraphs, topLevelElements, elementInner, attrValue } from "../src/xml.js";

test("createDocument().build(): an empty document is already clean (zero findings)", async () => {
  const buffer = createDocument().build();
  const doc = await openDocx(buffer);
  const { findings } = scan(doc);
  assert.deepEqual(findings, []);
});

test("paragraph(): rPr is always last in pPr regardless of option order, so heading-pPr-order never fires", async () => {
  const doc = createDocument();
  // Deliberately combine bold (-> pPr rPr) with spacing/ind/list — the
  // exact shape that caused the Weekend 3 bug when hand-assembled.
  const scope = doc.numberedList({ format: "decimal" });
  doc.append(
    paragraph({ text: "Heading", bold: true, spacing: { before: 240, after: 120 }, ind: { left: 100 } }),
    paragraph({ text: "First", list: scope }),
  );
  const buffer = doc.build();

  const opened = await openDocx(buffer);
  const { findings } = scan(opened, { rules: ["heading-pPr-order"] });
  assert.deepEqual(findings, []);

  // Confirm it's not just "no finding" but structurally correct: rPr is
  // the LAST child of pPr in the raw XML.
  const [heading] = findParagraphs(opened.text("word/document.xml"));
  const pPr = topLevelElements(elementInner(heading.xml)).find((el) => el.name === "w:pPr");
  const children = topLevelElements(elementInner(pPr.xml));
  assert.equal(children.at(-1).name, "w:rPr");
});

test("numberedList(): reusing the same handle keeps numbering continuous; a second call is a genuinely separate list", async () => {
  const doc = createDocument();
  const scopeA = doc.numberedList({ format: "decimal" });
  doc.append(
    paragraph({ text: "A1", list: scopeA }),
    paragraph({ text: "A2", list: scopeA }),
  );
  const scopeB = doc.numberedList({ format: "lowerLetter" });
  doc.append(paragraph({ text: "B1", list: scopeB }));

  const buffer = doc.build();
  const opened = await openDocx(buffer);

  // Reusing scopeA must never trigger numbering-restart against itself...
  const { findings } = scan(opened, { rules: ["numbering-restart"] });
  assert.deepEqual(findings, []);
  // ...and scopeA/scopeB really are different numIds (a real second list,
  // not an accidental duplicate of the first).
  assert.notEqual(scopeA.numId, scopeB.numId);
});

test("token(): always lands as one complete run, so split-run-risk never fires", async () => {
  const doc = createDocument();
  doc.append(paragraph({ children: ["The project is ", token("PROJECT_NAME"), ", due soon."] }));
  const buffer = doc.build();

  const opened = await openDocx(buffer);
  const { findings } = scan(opened, { rules: ["split-run-risk"] });
  assert.deepEqual(findings, []);

  // Confirm the token text is whole within a single <w:t>, not just that
  // the rule didn't flag it.
  const [para] = findParagraphs(opened.text("word/document.xml"));
  const runs = topLevelElements(elementInner(para.xml)).filter((el) => el.name === "w:r");
  const tokenRun = runs.find((r) => r.xml.includes("PROJECT_NAME"));
  assert.match(tokenRun.xml, /<w:t[^>]*>\{\{PROJECT_NAME\}\}<\/w:t>/);
});

test("build(): throws on a self-inflicted trailing-blank-pages defect, since the builder can't prevent that one by construction", () => {
  const doc = createDocument();
  doc.append(paragraph({ text: "Real content" }), paragraph(), paragraph(), paragraph());
  assert.throws(() => doc.build(), /trailing-blank-pages/);
});

test("build({ allowFindings: true }): the escape hatch actually works", () => {
  const doc = createDocument();
  doc.append(paragraph({ text: "Real content" }), paragraph(), paragraph(), paragraph());
  const buffer = doc.build({ allowFindings: true });
  assert.ok(buffer.length > 0);
});

test("numberedList(): rejects an unknown format rather than silently building something wrong", () => {
  const doc = createDocument();
  assert.throws(() => doc.numberedList({ format: "not-a-real-format" }), /unknown format/);
});

test("table(): builds a clean table with no findings, in the middle of a document", async () => {
  const doc = createDocument();
  doc.append(
    paragraph({ text: "Pricing" }),
    table({ rows: [["Item", "Price"], ["Widget", "$5"], ["Gadget", "$10"]] }),
    paragraph({ text: "Thanks for your business." }),
  );
  const buffer = doc.build();
  const opened = await openDocx(buffer);
  const { findings } = scan(opened);
  assert.deepEqual(findings, []);

  const xml = opened.text("word/document.xml");
  assert.match(xml, /<w:tbl>/);
  assert.equal((xml.match(/<w:tr>/g) || []).length, 3);
});

test("table(): rejects an empty rows array rather than building something meaningless", () => {
  assert.throws(() => table({ rows: [] }), /non-empty/);
});

test("table(): a single empty trailing cell does NOT trip trailing-blank-pages (its non-empty neighbor breaks the run)", async () => {
  const doc = createDocument();
  doc.append(paragraph({ text: "Notes" }), table({ rows: [["Item", "Price"], ["Widget", ""]] }));
  assert.doesNotThrow(() => doc.build());
});

test("table(): KNOWN LIMITATION — 2+ consecutive empty trailing cells (e.g. an empty last row) DOES trip the trailing-blank-pages self-scan", async () => {
  // This is the interaction documented in table()'s docstring and
  // ROADMAP.md Weekend 8: the scan rules find <w:p> elements with a
  // whole-document text scan that doesn't know about table nesting, so a
  // table's own cell paragraphs can be mistaken for stray trailing body
  // paragraphs — but only once there are 2+ consecutive empty ones at the
  // very end (see the single-empty-cell test above, which is the case
  // that turned out NOT to trigger it, discovered by actually running
  // this rather than assuming). Asserting the real observed behavior here
  // so a future fix to the underlying rule has a regression test to
  // satisfy.
  const doc = createDocument();
  doc.append(
    paragraph({ text: "Notes" }),
    table({ rows: [["Item", "Price"], ["Widget", "$5"], ["", ""]] }), // whole last row empty
  );
  assert.throws(() => doc.build(), /trailing-blank-pages/);

  // The documented workaround (end with a real paragraph after the table)
  // avoids it entirely.
  const doc2 = createDocument();
  doc2.append(
    paragraph({ text: "Notes" }),
    table({ rows: [["Item", "Price"], ["Widget", "$5"], ["", ""]] }),
    paragraph({ text: "End of table." }),
  );
  assert.doesNotThrow(() => doc2.build());
});

test("pageBreak(): a standalone page-break paragraph is never mistaken for a trailing blank", async () => {
  const doc = createDocument();
  doc.append(paragraph({ text: "Page one." }), pageBreak(), paragraph({ text: "Page two." }));
  const buffer = doc.build(); // must not throw
  const opened = await openDocx(buffer);
  assert.match(opened.text("word/document.xml"), /<w:br w:type="page"\/>/);
});

test("paragraph({ pageBreakBefore }): sets pageBreakBefore correctly ordered ahead of numPr/spacing", async () => {
  const doc = createDocument();
  const scope = doc.numberedList({ format: "decimal" });
  doc.append(paragraph({ text: "New section", pageBreakBefore: true, spacing: { before: 240 }, list: scope }));
  const buffer = doc.build();
  const opened = await openDocx(buffer);
  const { findings } = scan(opened, { rules: ["heading-pPr-order"] });
  assert.deepEqual(findings, []);

  const [para] = findParagraphs(opened.text("word/document.xml"));
  const pPr = topLevelElements(elementInner(para.xml)).find((el) => el.name === "w:pPr");
  const children = topLevelElements(elementInner(pPr.xml));
  assert.equal(children[0].name, "w:pageBreakBefore");
});

test("doc.hyperlink(): produces a real relationship and a paragraph docx-doctor finds no fault with", async () => {
  const doc = createDocument();
  const link = doc.hyperlink("our website", "https://example.com");
  doc.append(paragraph({ children: ["Visit ", link, " for more."] }));
  const buffer = doc.build();

  const opened = await openDocx(buffer);
  const { findings } = scan(opened);
  assert.deepEqual(findings, []);

  const relsXml = opened.text("word/_rels/document.xml.rels");
  assert.match(relsXml, new RegExp(`Id="${link.rId}"[^>]*Target="https://example\\.com"`));
  const docXml = opened.text("word/document.xml");
  assert.match(docXml, new RegExp(`<w:hyperlink r:id="${link.rId}">`));
  assert.match(docXml, /our website/);
});
