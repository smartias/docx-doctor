// A minimal, spec-correct, undamaged .docx: heading, body paragraph, a
// 2-item numbered list. Built in code rather than committed as a binary so
// it's readable/diffable in git — the same reasoning behind setty-docx's
// own test fixture. LibreOffice headless conversion isn't available in
// this sandbox to auto-verify it opens in a real office suite, so do that
// manually once, in real Word or LibreOffice — see ROADMAP.md Weekend 1.
import { buildDocxFromDocument } from "./parts.mjs";

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>This is a heading</w:t></w:r></w:p>
<w:p><w:r><w:t>This is a body paragraph with some ordinary text to round-trip.</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First item</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Second item</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body>
</w:document>`;

const { parts, buffer } = buildDocxFromDocument(DOCUMENT);

export const MINIMAL_DOCX_PARTS = parts;

/** @returns {Uint8Array} a fresh, valid .docx buffer (same bytes every call) */
export function buildMinimalDocx() {
  return buffer.slice();
}
