import { topLevelElements, elementInner, attrValue } from "../xml.js";

const FORMATS = {
  decimal: { numFmt: "decimal", lvlText: "%1." },
  lowerLetter: { numFmt: "lowerLetter", lvlText: "%1)" },
  upperLetter: { numFmt: "upperLetter", lvlText: "%1)" },
  lowerRoman: { numFmt: "lowerRoman", lvlText: "%1." },
  bullet: { numFmt: "bullet", lvlText: "•" },
};

function nextId(numberingXml, elementName, idAttr) {
  const root = topLevelElements(numberingXml).find((el) => el.name === "w:numbering");
  const inner = root ? elementInner(root.xml) || "" : "";
  const ids = topLevelElements(inner)
    .filter((el) => el.name === elementName)
    .map((el) => Number(attrValue(el.xml, idAttr)))
    .filter((n) => Number.isFinite(n));
  return String(ids.length ? Math.max(...ids) + 1 : 0);
}

/**
 * Add one new abstractNum + num pair to numbering.xml and return the
 * updated XML plus the new numId. Each call always allocates a genuinely
 * NEW list — this is deliberate: "start a new list" (call this again) and
 * "continue an existing list" (reuse the numId this already returned) are
 * different operations at the call site, so the numbering-restart bug
 * (Weekend 4 — a second <w:num> silently duplicating an existing list)
 * can't happen by accident here. It would take a deliberate second call.
 *
 * @param {string} numberingXml
 * @param {{ format?: keyof typeof FORMATS, start?: number }} [opts]
 * @returns {{ xml: string, numId: string, abstractNumId: string }}
 */
export function allocateList(numberingXml, { format = "decimal", start = 1 } = {}) {
  const spec = FORMATS[format];
  if (!spec) {
    throw new Error(`numberedList: unknown format "${format}" (known: ${Object.keys(FORMATS).join(", ")})`);
  }

  const abstractNumId = nextId(numberingXml, "w:abstractNum", "w:abstractNumId");
  const numId = nextId(numberingXml, "w:num", "w:numId");

  const abstractNumXml =
    `<w:abstractNum w:abstractNumId="${abstractNumId}">` +
    `<w:lvl w:ilvl="0"><w:start w:val="${start}"/><w:numFmt w:val="${spec.numFmt}"/>` +
    `<w:lvlText w:val="${spec.lvlText}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`;
  const numXml = `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractNumId}"/></w:num>`;

  // CT_Numbering requires every <w:abstractNum> before every <w:num>.
  const firstNumIdx = numberingXml.indexOf("<w:num ");
  const rootCloseIdx = numberingXml.lastIndexOf("</w:numbering>");
  let xml;
  if (firstNumIdx !== -1 && firstNumIdx < rootCloseIdx) {
    xml = numberingXml.slice(0, firstNumIdx) + abstractNumXml +
      numberingXml.slice(firstNumIdx, rootCloseIdx) + numXml + numberingXml.slice(rootCloseIdx);
  } else {
    xml = numberingXml.slice(0, rootCloseIdx) + abstractNumXml + numXml + numberingXml.slice(rootCloseIdx);
  }

  return { xml, numId, abstractNumId };
}
