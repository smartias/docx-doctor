/**
 * numbering-restart
 *
 * A list that should be continuous (e.g. a numbered clause list spanning
 * several pages) sometimes restarts numbering mid-section, either because
 * a paragraph got pointed at a new numId in word/numbering.xml, or because
 * a startOverride was left behind from an earlier edit.
 *
 * detect(): parse the abstractNum/num mapping in word/numbering.xml, walk
 * paragraphs in word/document.xml, and flag numbering breaks between
 * paragraphs that are otherwise the same list (same style/indent).
 * repair(): point the restarted paragraphs back at the original numId.
 *
 * This is the hardest rule — real templates vary a lot in how they do
 * this. Aim for "flags the obvious cases, never false-positives on a
 * clean list," not perfect coverage. Build order: Weekend 4 (see
 * ROADMAP.md); depends on the round-trip guarantee from Weekend 1.
 */
export const numberingRestartRule = {
  id: "numbering-restart",
  severity: "error",

  /** @param {import("../index.js").Document} doc @returns {import("../index.js").Finding[]} */
  detect(doc) {
    throw new Error("numbering-restart.detect: not implemented — see ROADMAP.md Weekend 4");
  },

  /** @param {import("../index.js").Document} doc @param {import("../index.js").Finding} finding */
  repair(doc, finding) {
    throw new Error("numbering-restart.repair: not implemented — see ROADMAP.md Weekend 4");
  },
};
