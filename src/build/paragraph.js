import { xmlEscape } from "../xml.js";

function runPropsInner({ bold, italic }) {
  let inner = "";
  if (bold) inner += "<w:b/>";
  if (italic) inner += "<w:i/>";
  return inner;
}

function runXml(text, { bold, italic } = {}) {
  const rPrInner = runPropsInner({ bold, italic });
  const rPr = rPrInner ? `<w:rPr>${rPrInner}</w:rPr>` : "";
  return `<w:r>${rPr}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`;
}

function attrsXml(obj) {
  return Object.entries(obj).map(([k, v]) => `w:${k}="${v}"`).join(" ");
}

/**
 * Build a <w:pPr>, ALWAYS placing the paragraph-mark run properties
 * (<w:rPr>) last — the exact ordering heading-pPr-order exists to check.
 * Callers pass options in whatever order is convenient; output order is
 * fixed here, not derived from call-site argument order, so the Weekend 3
 * bug (something appended after an already-present rPr) can't occur
 * through this function no matter what the caller does.
 */
function pPrXml({ list, spacing, ind, jc, bold, italic }) {
  const parts = [];
  if (list) parts.push(`<w:numPr><w:ilvl w:val="${list.ilvl ?? 0}"/><w:numId w:val="${list.numId}"/></w:numPr>`);
  if (spacing) parts.push(`<w:spacing ${attrsXml(spacing)}/>`);
  if (ind) parts.push(`<w:ind ${attrsXml(ind)}/>`);
  if (jc) parts.push(`<w:jc w:val="${jc}"/>`);

  const rPrInner = runPropsInner({ bold, italic });
  if (rPrInner) parts.push(`<w:rPr>${rPrInner}</w:rPr>`); // always last, see docstring above

  return parts.length ? `<w:pPr>${parts.join("")}</w:pPr>` : "";
}

/**
 * A placeholder token, guaranteed to be emitted as ONE complete run — the
 * split-run-risk defect (Weekend 5) requires a token's text to be
 * fragmented across multiple <w:r> elements, which can't happen to
 * something built this way, since this always becomes exactly one <w:r>
 * with the whole "{{NAME}}" in a single <w:t>.
 * @param {string} name
 */
export function token(name) {
  return { __docxDoctorToken: true, name };
}

function isToken(x) {
  return x !== null && typeof x === "object" && x.__docxDoctorToken === true;
}

/**
 * Build one <w:p>. Either pass `text` (plus optional `bold`/`italic`) for
 * a single-run paragraph, or `children` — an array of strings, `token()`
 * results, and/or `{ text, bold?, italic? }` objects — for a paragraph
 * made of multiple runs (e.g. a token() embedded mid-sentence).
 *
 * @param {{
 *   text?: string, children?: Array<string|{text:string,bold?:boolean,italic?:boolean}|ReturnType<typeof token>>,
 *   bold?: boolean, italic?: boolean,
 *   spacing?: Record<string,string|number>, ind?: Record<string,string|number>, jc?: string,
 *   list?: { numId: string, ilvl?: number },
 * }} [opts]
 * @returns {string} a complete "<w:p>...</w:p>" element
 */
export function paragraph({ text, children, bold, italic, spacing, ind, jc, list } = {}) {
  const pPr = pPrXml({ list, spacing, ind, jc, bold, italic });

  let runs;
  if (children) {
    runs = children.map((child) => {
      if (isToken(child)) return runXml(`{{${child.name}}}`);
      if (typeof child === "string") return runXml(child);
      return runXml(child.text, { bold: child.bold, italic: child.italic });
    }).join("");
  } else {
    runs = text !== undefined ? runXml(text, { bold, italic }) : "";
  }

  return `<w:p>${pPr}${runs}</w:p>`;
}
