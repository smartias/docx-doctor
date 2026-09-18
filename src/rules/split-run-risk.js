/**
 * split-run-risk
 *
 * Word splits a paragraph's visible text across multiple <w:r> runs for
 * reasons that have nothing to do with formatting (rsid tracking, spell-
 * check markers, prior edits) — so a token like "{{PROJECT_NAME}}" can be
 * split across two or three runs in the underlying XML even though it
 * looks like one word on screen. Naive string replacement on the raw XML
 * then silently fails to match.
 *
 * This rule is detect-only by design: fixing a split token is a
 * templating decision (how do you want the merged run's formatting to
 * look?), not a structural repair — hand it to a templating tool
 * (docxtemplater or a character-stream-aware replacer) instead.
 *
 * Build order: Weekend 5 (see ROADMAP.md).
 */
export const splitRunRiskRule = {
  id: "split-run-risk",
  severity: "warn",

  /** @param {import("../index.js").Document} doc @returns {import("../index.js").Finding[]} */
  detect(doc) {
    throw new Error("split-run-risk.detect: not implemented — see ROADMAP.md Weekend 5");
  },

  // No repair() — this rule is a linter, not a fixer. See module docstring.
};
