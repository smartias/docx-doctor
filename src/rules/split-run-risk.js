import { findParagraphs, topLevelElements, elementInner, paragraphText } from "../xml.js";

const DOC_PART = "word/document.xml";

// The two most common template-placeholder conventions: {{TOKEN}} and
// ${TOKEN}. Not configurable yet — good enough to prove the detection
// mechanism; a real second template will tell us what else to add.
const TOKEN_RE = /\{\{[^{}]+\}\}|\$\{[^{}]+\}/g;

/** Per-run text with its [start, end) offset into the paragraph's concatenated text. */
function runTextSpans(paraInner) {
  const runs = topLevelElements(paraInner).filter((el) => el.name === "w:r");
  let cursor = 0;
  return runs.map((run) => {
    const text = paragraphText(run.xml);
    const span = { start: cursor, end: cursor + text.length, text };
    cursor += text.length;
    return span;
  });
}

/**
 * split-run-risk
 *
 * Word splits a paragraph's visible text across multiple <w:r> runs for
 * reasons unrelated to formatting (rsid tracking, spell-check markers,
 * prior edits) — so a token like "{{PROJECT_NAME}}" can be split across
 * two or three runs in the underlying XML even though it reads as one
 * word on screen. A naive string replacement on the raw XML then
 * silently fails to match.
 *
 * Detect-only by design: fixing a split token is a templating decision
 * (how should the merged run's formatting look?), not a structural
 * repair — hand it to a templating tool (docxtemplater, or a character-
 * stream-aware replacer) instead.
 *
 * v0 scope: only <w:r> elements that are direct children of the
 * paragraph are considered — a run wrapped in <w:hyperlink>, <w:ins>, or
 * <w:del> isn't walked. Deferred, not silently missed.
 */
export const splitRunRiskRule = {
  id: "split-run-risk",
  severity: "warn",

  /** @param {import("../document.js").Document} doc @returns {import("../index.js").Finding[]} */
  detect(doc) {
    const xml = doc.text(DOC_PART);
    if (!xml) return [];
    const paragraphs = findParagraphs(xml);
    const findings = [];

    paragraphs.forEach((para, paraIndex) => {
      const paraInner = elementInner(para.xml);
      if (!paraInner) return;
      const spans = runTextSpans(paraInner);
      if (spans.length === 0) return;
      const text = spans.map((s) => s.text).join("");

      TOKEN_RE.lastIndex = 0;
      let m;
      while ((m = TOKEN_RE.exec(text))) {
        const start = m.index;
        const end = start + m[0].length;
        const touched = spans.filter((s) => s.start < end && s.end > start);
        const isIntact = touched.length === 1 && touched[0].start <= start && end <= touched[0].end;
        if (isIntact) continue;

        findings.push({
          rule: "split-run-risk",
          severity: "warn",
          part: DOC_PART,
          location: { paraIndex, token: m[0], runsTouched: touched.length },
          message: `Paragraph ${paraIndex}: token "${m[0]}" is split across ${touched.length} runs — ` +
            `a find-replace on the raw XML will silently miss it. Merge the runs first, or use a ` +
            `run-boundary-aware templating tool instead of string-replacing directly.`,
          autofixable: false,
        });
      }
    });

    return findings;
  },

  // No repair() — detect-only by design, see module docstring above.
};
