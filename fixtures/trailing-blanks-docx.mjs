// Same shape as the minimal fixture, but damaged the way real templates
// get damaged: 3 empty paragraphs accumulated at the end of the body
// before sectPr (a self-closing <w:p/>, a whitespace-only run, and
// another self-closing one) — the defect trailing-blank-pages.detect()
// should find and repair() should remove, leaving the real content and
// exactly the conventional single trailing paragraph intact.
import { buildDocxFromDocument } from "./parts.mjs";

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>This is a heading</w:t></w:r></w:p>
<w:p><w:r><w:t>This is a body paragraph with real content.</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:t xml:space="preserve">   </w:t></w:r></w:p>
<w:p/>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body>
</w:document>`;

const { parts, buffer } = buildDocxFromDocument(DOCUMENT);

export const TRAILING_BLANKS_DOCX_PARTS = parts;
export const TRAILING_BLANKS_REAL_CONTENT = [
  "This is a heading",
  "This is a body paragraph with real content.",
];

/** @returns {Uint8Array} */
export function buildTrailingBlanksDocx() {
  return buffer.slice();
}
