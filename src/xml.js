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

const TAG_RE = /<\/?[\w:-]+(?:\s[^>]*?)?\/?>/g;

/**
 * Walk the DIRECT children of an XML fragment (e.g. the contents of a
 * <w:pPr>...</w:pPr>) by tracking generic open/close depth — works for any
 * element names, not just <w:p>, as long as the XML is well-formed.
 * @param {string} xml
 * @returns {{ name: string, xml: string, start: number, end: number }[]}
 */
export function topLevelElements(xml) {
  const out = [];
  let depth = 0;
  let curStart = -1;
  let curName = null;
  TAG_RE.lastIndex = 0;
  let m;
  while ((m = TAG_RE.exec(xml))) {
    const tag = m[0];
    const isClose = tag.startsWith("</");
    const isSelfClose = !isClose && tag.endsWith("/>");
    const name = tag.match(/^<\/?([\w:-]+)/)[1];

    if (isSelfClose) {
      if (depth === 0) out.push({ name, xml: tag, start: m.index, end: m.index + tag.length });
      continue;
    }
    if (!isClose) {
      if (depth === 0) { curStart = m.index; curName = name; }
      depth++;
    } else {
      depth--;
      if (depth === 0) {
        const end = m.index + tag.length;
        out.push({ name: curName, xml: xml.slice(curStart, end), start: curStart, end });
      }
    }
  }
  return out;
}

/**
 * The content between an element's opening and closing tag, or null if the
 * element is self-closing (no children possible).
 * @param {string} elementXml - a single element, as returned by
 *   topLevelElements/findParagraphs (e.g. "<w:pPr>...</w:pPr>")
 */
export function elementInner(elementXml) {
  const openEnd = elementXml.indexOf(">") + 1;
  if (elementXml.slice(0, openEnd).endsWith("/>")) return null;
  const closeStart = elementXml.lastIndexOf("<");
  return elementXml.slice(openEnd, closeStart);
}

/** An attribute value from an element's opening tag, e.g. attrValue(el, "w:val"). */
export function attrValue(elementXml, name) {
  const m = elementXml.match(new RegExp(`${name}="([^"]*)"`));
  return m ? m[1] : undefined;
}

/** Escape text for safe placement inside a <w:t> element. */
export function xmlEscape(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
