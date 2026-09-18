/**
 * trailing-blank-pages
 *
 * Word templates accumulate empty paragraphs at the end of the body over
 * repeated save/edit cycles. Detect a run of empty <w:p> elements (no runs,
 * or runs with no visible text) immediately before the final sectPr, and
 * offer to remove them.
 *
 * Build order: Weekend 2 (see ROADMAP.md). Needs paragraph-walking over
 * word/document.xml first.
 */
export const trailingBlankPagesRule = {
  id: "trailing-blank-pages",
  severity: "warn",

  /** @param {import("../index.js").Document} doc @returns {import("../index.js").Finding[]} */
  detect(doc) {
    throw new Error("trailing-blank-pages.detect: not implemented — see ROADMAP.md Weekend 2");
  },

  /** @param {import("../index.js").Document} doc @param {import("../index.js").Finding} finding */
  repair(doc, finding) {
    throw new Error("trailing-blank-pages.repair: not implemented — see ROADMAP.md Weekend 2");
  },
};
