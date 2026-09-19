import { topLevelElements, elementInner, attrValue } from "../xml.js";

function nextRid(relsXml) {
  const root = topLevelElements(relsXml).find((el) => el.name === "Relationships");
  const inner = root ? elementInner(root.xml) || "" : "";
  const ids = topLevelElements(inner)
    .filter((el) => el.name === "Relationship")
    .map((el) => attrValue(el.xml, "Id"))
    .map((id) => (id && /^rId(\d+)$/.test(id) ? Number(id.slice(3)) : null))
    .filter((n) => n !== null);
  return "rId" + String(ids.length ? Math.max(...ids) + 1 : 1);
}

/**
 * Add an External hyperlink relationship to word/_rels/document.xml.rels.
 * @param {string} relsXml
 * @param {string} url
 * @returns {{ xml: string, rId: string }}
 */
export function allocateHyperlinkRel(relsXml, url) {
  const rId = nextRid(relsXml);
  const rel = `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>`;
  const closeIdx = relsXml.lastIndexOf("</Relationships>");
  const xml = relsXml.slice(0, closeIdx) + rel + relsXml.slice(closeIdx);
  return { xml, rId };
}
