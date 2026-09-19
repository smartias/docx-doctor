// Five paragraphs exercising the heading-pPr-order defect and its
// neighbors, so the rule has to actually discriminate rather than just
// matching "has an rPr":
//   0: BROKEN — rPr followed by one element (spacing). The classic
//      append-after-existing-rPr bug.
//   1: correct — spacing before rPr (rPr last). Must NOT be flagged.
//   2: pPr with no rPr at all (just numPr). Must NOT be flagged.
//   3: no pPr at all. Must NOT be flagged.
//   4: BROKEN — rPr followed by TWO elements (ind, spacing), to prove
//      repair() preserves their relative order and doesn't just swap two.
import { buildDocxFromDocument } from "./parts.mjs";

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:rPr><w:b/></w:rPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:r><w:t>Broken heading spacing</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:before="240" w:after="120"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Correctly ordered heading</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>List item, no rPr in pPr</w:t></w:r></w:p>
<w:p><w:r><w:t>Plain paragraph, no pPr at all</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:b/></w:rPr><w:ind w:left="720"/><w:spacing w:before="240" w:after="120"/></w:pPr><w:r><w:t>Broken with two trailing elements</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body>
</w:document>`;

const { parts, buffer } = buildDocxFromDocument(DOCUMENT);

export const HEADING_PPR_ORDER_DOCX_PARTS = parts;

/** @returns {Uint8Array} */
export function buildHeadingPPrOrderDocx() {
  return buffer.slice();
}
