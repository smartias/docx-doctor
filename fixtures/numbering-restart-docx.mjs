// A numbering.xml with two abstract lists (0: decimal, 1: lower-letter)
// and three <w:num> entries: numId 1 and 2 both point at abstractNumId 0
// (the restart bug — a second num pointing at the same list), numId 3
// points at abstractNumId 1 (a genuinely different list type).
//
// document.xml paragraphs:
//   0: heading, not a list item
//   1: body paragraph, not a list item
//   2-3: list items using numId 1 ("First", "Second")
//   4-5: list items using numId 2 — SAME abstractNumId as numId 1's list,
//        so Word restarts the count here even though it reads as one
//        continuous list ("Third", "Fourth")
//   6:   list item using numId 3 — a genuinely different list
//        (abstractNumId 1), must NOT be flagged ("A")
//   7:   plain paragraph, breaks the run
//   8:   a lone list item reusing numId 1 — a run of length 1, nothing to
//        compare against, must NOT be flagged ("Sixth")
import { buildDocx } from "./parts.mjs";

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="1">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1)"/><w:lvlJc w:val="left"/>
<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const numPara = (numId, text) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Heading</w:t></w:r></w:p>
<w:p><w:r><w:t>Body paragraph, not a list item.</w:t></w:r></w:p>
${numPara(1, "First")}
${numPara(1, "Second")}
${numPara(2, "Third")}
${numPara(2, "Fourth")}
${numPara(3, "A")}
<w:p><w:r><w:t>A plain paragraph, breaking the run.</w:t></w:r></w:p>
${numPara(1, "Sixth")}
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body>
</w:document>`;

const { parts, buffer } = buildDocx({ documentXml: DOCUMENT, numberingXml: NUMBERING });

export const NUMBERING_RESTART_DOCX_PARTS = parts;

/** @returns {Uint8Array} */
export function buildNumberingRestartDocx() {
  return buffer.slice();
}
