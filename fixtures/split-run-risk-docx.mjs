// Five paragraphs exercising the split-run-risk defect and its neighbors:
//   0: heading, no tokens at all
//   1: BROKEN — "{{PROJECT_NAME}}" split across 2 runs ("{{PROJECT_" | "NAME}}")
//   2: correct — the same-looking token, but intact within a single run
//   3: BROKEN — "${CLIENT_NAME}" split across 3 runs
//   4: plain text, no tokens
import { buildDocxFromDocument } from "./parts.mjs";

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Heading</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Hello {{PROJECT_</w:t></w:r><w:r><w:t>NAME}}</w:t></w:r></w:p>
<w:p><w:r><w:t>This is the {{PROJECT_NAME}} paragraph, intact.</w:t></w:r></w:p>
<w:p><w:r><w:t>\${</w:t></w:r><w:r><w:t>CLIENT</w:t></w:r><w:r><w:t>_NAME}</w:t></w:r></w:p>
<w:p><w:r><w:t>Just a normal paragraph with no tokens.</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body>
</w:document>`;

const { parts, buffer } = buildDocxFromDocument(DOCUMENT);

export const SPLIT_RUN_RISK_DOCX_PARTS = parts;

/** @returns {Uint8Array} */
export function buildSplitRunRiskDocx() {
  return buffer.slice();
}
