/**
 * heading-pPr-order
 *
 * Word paragraph properties (<w:pPr>) are order-sensitive. Per ECMA-376
 * §17.3.1.29, the paragraph-mark run properties (<w:rPr>) must come last
 * inside <w:pPr> — any element inserted after it (spacing, numbering, etc.)
 * is silently ignored by Word. Naive code that appends new pPr children
 * right before </w:pPr> hits this constantly.
 *
 * detect(): flag any pPr where a spacing/numbering element appears after
 * the paragraph-mark rPr.
 * repair(): reorder elements into valid position (insert-before-rPr, not
 * append-before-close-tag).
 *
 * Build order: Weekend 3 (see ROADMAP.md).
 */
export const headingPPrOrderRule = {
  id: "heading-ppr-order",
  severity: "error",

  /** @param {import("../index.js").Document} doc @returns {import("../index.js").Finding[]} */
  detect(doc) {
    throw new Error("heading-ppr-order.detect: not implemented — see ROADMAP.md Weekend 3");
  },

  /** @param {import("../index.js").Document} doc @param {import("../index.js").Finding} finding */
  repair(doc, finding) {
    throw new Error("heading-ppr-order.repair: not implemented — see ROADMAP.md Weekend 3");
  },
};
