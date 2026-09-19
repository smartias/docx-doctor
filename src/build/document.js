import { Document } from "../document.js";
import { scan } from "../index.js";
import { emptyDocumentParts } from "./skeleton.js";
import { allocateList } from "./numbering.js";

const DOC_PART = "word/document.xml";
const NUMBERING_PART = "word/numbering.xml";

function insertBeforeSectPr(bodyXml, insertXml) {
  const idx = bodyXml.indexOf("<w:sectPr");
  if (idx === -1) return bodyXml + insertXml; // defensive — every skeleton/valid body has one
  return bodyXml.slice(0, idx) + insertXml + bodyXml.slice(idx);
}

export class BuilderDocument {
  constructor() {
    this.doc = new Document(new Map());
    for (const [path, content] of Object.entries(emptyDocumentParts())) {
      this.doc.setText(path, content);
    }
  }

  /** Append one or more paragraph() results to the end of the body. */
  append(...paragraphXmlStrings) {
    const xml = this.doc.text(DOC_PART);
    this.doc.setText(DOC_PART, insertBeforeSectPr(xml, paragraphXmlStrings.join("")));
    return this;
  }

  /**
   * Allocate a NEW list (one abstractNum + one num). Returns a handle —
   * pass it as paragraph({ list: handle }) on every paragraph that should
   * be part of THIS list; reusing the same handle is what keeps the
   * numbering continuous. Call numberedList() again only when you
   * actually want a second, independent list.
   * @param {{ format?: string, start?: number }} [opts]
   * @returns {{ numId: string, ilvl?: number }}
   */
  numberedList(opts = {}) {
    const { xml, numId } = allocateList(this.doc.text(NUMBERING_PART), opts);
    this.doc.setText(NUMBERING_PART, xml);
    return { numId };
  }

  /**
   * Finish the document. By default, runs docx-doctor's own scan() against
   * what you built and throws if it finds anything — a guarantee that a
   * document built with this API never leaves with an unnoticed defect
   * (the one rule the builder can't prevent by construction alone is
   * trailing-blank-pages, since nothing stops you from appending empty
   * paragraphs on purpose; this is the backstop for that).
   * @param {{ allowFindings?: boolean }} [opts]
   * @returns {Uint8Array}
   */
  build({ allowFindings = false } = {}) {
    if (!allowFindings) {
      const { findings } = scan(this.doc);
      if (findings.length > 0) {
        const details = findings.map((f) => `  - [${f.severity}] ${f.rule}: ${f.message}`).join("\n");
        throw new Error(
          `docx-doctor build(): ${findings.length} finding(s) in the document you constructed. ` +
          `Fix the construction, or pass { allowFindings: true } to build it anyway:\n${details}`
        );
      }
    }
    return this.doc.toBuffer();
  }
}

/** @returns {BuilderDocument} a valid, empty .docx ready for append()/numberedList() */
export function createDocument() {
  return new BuilderDocument();
}
