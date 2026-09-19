import { findParagraphs, topLevelElements, elementInner, attrValue } from "../xml.js";
import { parseNumbering } from "../numbering.js";

const DOC_PART = "word/document.xml";
const NUMBERING_PART = "word/numbering.xml";

/** {ilvl, numId} for a paragraph's <w:numPr>, or null if it isn't a list item. */
function paragraphNumPr(paraXml) {
  const paraInner = elementInner(paraXml);
  if (!paraInner) return null;
  const pPr = topLevelElements(paraInner).find((el) => el.name === "w:pPr");
  if (!pPr) return null;
  const pPrInner = elementInner(pPr.xml);
  if (!pPrInner) return null;
  const numPr = topLevelElements(pPrInner).find((el) => el.name === "w:numPr");
  if (!numPr) return null;
  const numPrInner = elementInner(numPr.xml) || "";
  const children = topLevelElements(numPrInner);
  const numIdEl = children.find((c) => c.name === "w:numId");
  if (!numIdEl) return null;
  const ilvlEl = children.find((c) => c.name === "w:ilvl");
  return {
    ilvl: ilvlEl ? Number(attrValue(ilvlEl.xml, "w:val")) : 0,
    numId: attrValue(numIdEl.xml, "w:val"),
  };
}

/**
 * Within one run of consecutive list paragraphs [runStart, runEnd], flag
 * any contiguous block whose numId differs from the run's first (the
 * "canonical" numId) but resolves to the SAME abstractNumId — that's the
 * restart: a second <w:num> pointing at the same list definition, so Word
 * starts a fresh count instead of continuing.
 *
 * A numId change to a genuinely DIFFERENT abstractNumId is never flagged
 * — that's switching list types (numbers to letters, etc.), not a defect.
 */
function findRestartsInRun(listInfo, runStart, runEnd, nums, findings) {
  const canonicalNumId = listInfo[runStart].numId;
  const canonicalAbstract = nums.get(canonicalNumId)?.abstractNumId;
  if (canonicalAbstract === undefined) return;

  let blockStart = runStart;
  let blockNumId = canonicalNumId;

  const flush = (blockEnd) => {
    if (blockNumId === canonicalNumId) return;
    const abstractNumId = nums.get(blockNumId)?.abstractNumId;
    if (abstractNumId !== canonicalAbstract) return;
    findings.push({
      rule: "numbering-restart",
      severity: "error",
      part: DOC_PART,
      location: { paraIndexStart: blockStart, paraIndexEnd: blockEnd, canonicalNumId, restartedNumId: blockNumId },
      message: `Paragraphs ${blockStart}-${blockEnd} restart numbering: numId ${blockNumId} points at ` +
        `the same list definition (abstractNumId ${abstractNumId}) as numId ${canonicalNumId}, used ` +
        `earlier in this same list, so Word restarts the count instead of continuing it.`,
      autofixable: true,
    });
  };

  for (let idx = runStart + 1; idx <= runEnd; idx++) {
    const numId = listInfo[idx].numId;
    if (numId !== blockNumId) {
      flush(idx - 1);
      blockStart = idx;
      blockNumId = numId;
    }
  }
  flush(runEnd);
}

export const numberingRestartRule = {
  id: "numbering-restart",
  severity: "error",

  /** @param {import("../document.js").Document} doc @returns {import("../index.js").Finding[]} */
  detect(doc) {
    const docXml = doc.text(DOC_PART);
    const numberingXml = doc.text(NUMBERING_PART);
    if (!docXml || !numberingXml) return [];

    const nums = parseNumbering(numberingXml);
    const paragraphs = findParagraphs(docXml);
    const listInfo = paragraphs.map((p) => paragraphNumPr(p.xml));

    const findings = [];
    let i = 0;
    while (i < paragraphs.length) {
      if (!listInfo[i]) { i++; continue; }
      let j = i;
      while (j + 1 < paragraphs.length && listInfo[j + 1]) j++;
      if (j > i) findRestartsInRun(listInfo, i, j, nums, findings);
      i = j + 1;
    }
    return findings;
  },

  /** @param {import("../document.js").Document} doc @param {import("../index.js").Finding} finding */
  repair(doc, finding) {
    const xml = doc.text(DOC_PART);
    if (!xml) return;
    const paragraphs = findParagraphs(xml);
    const { paraIndexStart, paraIndexEnd, canonicalNumId } = finding.location;

    let next = xml;
    // Right-to-left so edits to a later paragraph never invalidate the
    // still-original offsets of earlier ones (same reasoning as
    // trailing-blank-pages.repair()).
    for (let idx = paraIndexEnd; idx >= paraIndexStart; idx--) {
      const para = paragraphs[idx];
      const newParaXml = para.xml.replace(
        /(<w:numId\s+w:val=")[^"]*("\s*\/>)/,
        `$1${canonicalNumId}$2`
      );
      next = next.slice(0, para.start) + newParaXml + next.slice(para.end);
    }
    doc.setText(DOC_PART, next);
  },
};
