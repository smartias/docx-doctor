// Shared OOXML paragraph-walking utilities. Regex-based on purpose — a
// document.xml body is well-formed XML we control the shape of, and a full
// DOM parser buys nothing here but a dependency (same approach setty-docx.js
// takes for the same reason).

const PARA_RE = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>|<w:p(?:\s[^>]*)?\/>/g;
const TEXT_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;

// Content that makes a paragraph visually non-empty even with no <w:t>:
// images, embedded objects, explicit page breaks.
const NON_TEXT_CONTENT_RE = /<w:(drawing|pict|object)\b|<w:br\s+[^>]*w:type="page"/;

/**
 * @param {string} xml - a whole part's XML (e.g. word/document.xml)
 * @returns {{ xml: string, start: number, end: number }[]} paragraphs in
 *   document order, with byte offsets into the original string.
 */
export function findParagraphs(xml) {
  const out = [];
  PARA_RE.lastIndex = 0;
  let m;
  while ((m = PARA_RE.exec(xml))) {
    out.push({ xml: m[0], start: m.index, end: m.index + m[0].length });
  }
  return out;
}

/** Concatenated visible text of a single <w:p>...</w:p> XML string. */
export function paragraphText(paraXml) {
  let text = "";
  TEXT_RE.lastIndex = 0;
  let m;
  while ((m = TEXT_RE.exec(paraXml))) text += m[1];
  return text;
}

/** A paragraph is "empty" if it has no visible text and no non-text content. */
export function isEmptyParagraph(paraXml) {
  if (NON_TEXT_CONTENT_RE.test(paraXml)) return false;
  return paragraphText(paraXml).trim().length === 0;
}
