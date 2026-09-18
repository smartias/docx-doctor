/**
 * docx-doctor public API.
 *
 * Nothing below is implemented yet — these are the planned signatures.
 * See ROADMAP.md for the build order (openDocx/save first, then rules).
 */

/**
 * Parse a .docx (zip of OOXML parts) into an in-memory Document.
 * @param {Uint8Array|Buffer} buffer
 * @returns {Promise<Document>}
 */
export async function openDocx(buffer) {
  throw new Error("openDocx: not implemented yet — see ROADMAP.md Weekend 1");
}

/**
 * Run rules against a Document and collect findings without changing it.
 * @param {Document} doc
 * @param {{ rules?: string[] }} [opts] - rule ids to run; defaults to all.
 * @returns {{ findings: Finding[] }}
 */
export function scan(doc, opts = {}) {
  throw new Error("scan: not implemented yet — see ROADMAP.md Weekend 2");
}

/**
 * Apply autofixable rule repairs to a Document.
 * @param {Document} doc
 * @param {{ rules?: string[] }} [opts]
 * @returns {{ doc: Document, applied: Finding[], skipped: Finding[] }}
 */
export function fix(doc, opts = {}) {
  throw new Error("fix: not implemented yet — see ROADMAP.md Weekend 2");
}

export { RULES } from "./rules/index.js";

/**
 * @typedef {Object} Finding
 * @property {string} rule
 * @property {"error"|"warn"} severity
 * @property {string} part - e.g. "word/document.xml"
 * @property {Object} location - rule-specific (e.g. { paraIndex })
 * @property {string} message
 * @property {boolean} autofixable
 */

/**
 * @typedef {Object} Document
 * @property {Map<string, string|Uint8Array>} parts
 * @property {() => Promise<Uint8Array>} toBuffer
 * @property {(path: string) => Promise<void>} save
 */
