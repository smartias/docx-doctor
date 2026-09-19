import { findParagraphs, isEmptyParagraph } from "../xml.js";

const DOC_PART = "word/document.xml";

/**
 * trailing-blank-pages
 *
 * Word templates accumulate empty paragraphs at the end of the body over
 * repeated save/edit cycles. This rule flags a *run* of 2+ empty
 * paragraphs at the very end of the body — not a single trailing empty
 * paragraph, since documents conventionally end with one and flagging
 * that would be a false positive on every clean template.
 *
 * v0 scope: only handles the body-level <w:sectPr> case (a single- or
 * final-section document, where sectPr is a direct child of <w:body>
 * after the last paragraph). A sectPr nested inside a paragraph's <w:pPr>
 * (mid-document section breaks) isn't detected — out of scope until a
 * real multi-section template motivates it.
 */
export const trailingBlankPagesRule = {
  id: "trailing-blank-pages",
  severity: "warn",

  /** @param {import("../document.js").Document} doc @returns {import("../index.js").Finding[]} */
  detect(doc) {
    const xml = doc.text(DOC_PART);
    if (!xml) return [];
    const paragraphs = findParagraphs(xml);
    if (paragraphs.length === 0) return [];

    let firstTrailingEmpty = -1;
    for (let i = paragraphs.length - 1; i >= 0; i--) {
      if (!isEmptyParagraph(paragraphs[i].xml)) break;
      firstTrailingEmpty = i;
    }
    if (firstTrailingEmpty === -1) return [];

    const count = paragraphs.length - firstTrailingEmpty;
    if (count < 2) return []; // one trailing empty paragraph is normal, not a defect

    return [{
      rule: "trailing-blank-pages",
      severity: "warn",
      part: DOC_PART,
      location: { paraIndexStart: firstTrailingEmpty, paraIndexEnd: paragraphs.length - 1 },
      message: `${count} empty paragraphs at the end of the document body ` +
        `(indices ${firstTrailingEmpty}-${paragraphs.length - 1}) — likely a stray trailing blank page.`,
      autofixable: true,
    }];
  },

  /** @param {import("../document.js").Document} doc @param {import("../index.js").Finding} finding */
  repair(doc, finding) {
    const xml = doc.text(DOC_PART);
    if (!xml) return;
    const paragraphs = findParagraphs(xml);
    const { paraIndexStart, paraIndexEnd } = finding.location;

    // Safety: never remove every paragraph in the body — a section needs
    // at least one block element before its sectPr. If the whole run of
    // empties starts at paragraph 0, leave that first one behind.
    const removeStart = paraIndexStart === 0 ? 1 : paraIndexStart;
    if (removeStart > paraIndexEnd) return;

    let next = xml;
    for (let i = paraIndexEnd; i >= removeStart; i--) {
      const p = paragraphs[i];
      next = next.slice(0, p.start) + next.slice(p.end);
    }
    doc.setText(DOC_PART, next);
  },
};
