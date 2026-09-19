import { findParagraphs, topLevelElements, elementInner } from "../xml.js";

const DOC_PART = "word/document.xml";

/**
 * heading-pPr-order
 *
 * Word paragraph properties (<w:pPr>) are order-sensitive. Per ECMA-376
 * §17.3.1.29, the paragraph-mark run properties (<w:rPr>, a direct child
 * of pPr — distinct from a run's own rPr) must come last inside pPr. Code
 * that "adds a pPr setting" by appending right before </w:pPr> silently
 * breaks the moment an rPr is already present: Word ignores everything
 * placed after it, with no error or warning.
 *
 * v0 scope: only checks "is anything after rPr" — the actual defect this
 * pattern produces in practice (naive append-before-close-tag code) —
 * not the full CT_PPr schema ordering.
 */
export const headingPPrOrderRule = {
  id: "heading-ppr-order",
  severity: "error",

  /** @param {import("../document.js").Document} doc @returns {import("../index.js").Finding[]} */
  detect(doc) {
    const xml = doc.text(DOC_PART);
    if (!xml) return [];
    const paragraphs = findParagraphs(xml);
    const findings = [];

    paragraphs.forEach((para, paraIndex) => {
      const paraInner = elementInner(para.xml);
      if (!paraInner) return; // self-closing <w:p/> has no pPr
      const pPr = topLevelElements(paraInner).find((el) => el.name === "w:pPr");
      if (!pPr) return;
      const inner = elementInner(pPr.xml);
      if (!inner) return;

      const children = topLevelElements(inner);
      const rPrIndex = children.findIndex((el) => el.name === "w:rPr");
      if (rPrIndex === -1) return;
      const after = children.slice(rPrIndex + 1);
      if (after.length === 0) return;

      findings.push({
        rule: "heading-ppr-order",
        severity: "error",
        part: DOC_PART,
        location: { paraIndex, afterRPr: after.map((el) => el.name) },
        message: `Paragraph ${paraIndex}: <w:pPr> has ${after.length} element(s) ` +
          `(${after.map((el) => el.name).join(", ")}) after the paragraph-mark <w:rPr> — ` +
          `Word silently ignores anything placed there.`,
        autofixable: true,
      });
    });

    return findings;
  },

  /** @param {import("../document.js").Document} doc @param {import("../index.js").Finding} finding */
  repair(doc, finding) {
    const xml = doc.text(DOC_PART);
    if (!xml) return;
    const paragraphs = findParagraphs(xml);
    const para = paragraphs[finding.location.paraIndex];
    if (!para) return;

    const paraInner = elementInner(para.xml);
    if (!paraInner) return;
    const innerOffset = para.xml.indexOf(">") + 1; // where paraInner starts within para.xml

    const pPr = topLevelElements(paraInner).find((el) => el.name === "w:pPr");
    if (!pPr) return;
    const inner = elementInner(pPr.xml);
    if (!inner) return;

    const children = topLevelElements(inner);
    const rPrIndex = children.findIndex((el) => el.name === "w:rPr");
    if (rPrIndex === -1) return;

    // Move rPr to the end; every other child keeps its relative order.
    const reordered = [
      ...children.slice(0, rPrIndex),
      ...children.slice(rPrIndex + 1),
      children[rPrIndex],
    ].map((el) => el.xml).join("");

    const openTag = pPr.xml.slice(0, pPr.xml.indexOf(">") + 1);
    const newPPr = openTag + reordered + "</w:pPr>";
    const pPrStart = innerOffset + pPr.start;
    const pPrEnd = innerOffset + pPr.end;
    const newParaXml = para.xml.slice(0, pPrStart) + newPPr + para.xml.slice(pPrEnd);

    doc.setText(DOC_PART, xml.slice(0, para.start) + newParaXml + xml.slice(para.end));
  },
};
