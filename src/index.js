/**
 * docx-doctor public API.
 *
 * openDocx/save are real (Weekend 1). scan/fix are real as of Weekend 2,
 * but only trailing-blank-pages has a working detect()/repair() so far —
 * the other RULES entries still throw. Pass `opts.rules` explicitly until
 * every built-in rule is implemented (ROADMAP.md), or scan()/fix() with no
 * options will throw on the first unimplemented rule.
 */

import { Document } from "./document.js";
import { RULES } from "./rules/index.js";

export { Document, RULES };

/**
 * Parse a .docx (zip of OOXML parts) into an in-memory Document.
 * @param {Uint8Array|Buffer} buffer
 * @returns {Promise<Document>}
 */
export async function openDocx(buffer) {
  return Document.open(buffer);
}

/**
 * Run rules against a Document and collect findings without changing it.
 * @param {Document} doc
 * @param {{ rules?: string[] }} [opts] - rule ids to run; defaults to all.
 * @returns {{ findings: Finding[] }}
 */
export function scan(doc, opts = {}) {
  const rules = opts.rules ? RULES.filter((r) => opts.rules.includes(r.id)) : RULES;
  const findings = [];
  for (const rule of rules) findings.push(...rule.detect(doc));
  return { findings };
}

/**
 * Apply autofixable rule repairs to a Document. Mutates doc in place and
 * returns it (repair() mutates doc.parts directly), plus which findings
 * were applied vs. skipped (not autofixable, or the rule has no repair()).
 * @param {Document} doc
 * @param {{ rules?: string[] }} [opts]
 * @returns {{ doc: Document, applied: Finding[], skipped: Finding[] }}
 */
export function fix(doc, opts = {}) {
  const { findings } = scan(doc, opts);
  const applied = [];
  const skipped = [];
  for (const finding of findings) {
    const rule = RULES.find((r) => r.id === finding.rule);
    if (!finding.autofixable || !rule?.repair) {
      skipped.push(finding);
      continue;
    }
    rule.repair(doc, finding);
    applied.push(finding);
  }
  return { doc, applied, skipped };
}

/**
 * @typedef {Object} Finding
 * @property {string} rule
 * @property {"error"|"warn"} severity
 * @property {string} part - e.g. "word/document.xml"
 * @property {Object} location - rule-specific (e.g. { paraIndex })
 * @property {string} message
 * @property {boolean} autofixable
 */
