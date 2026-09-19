import { topLevelElements, elementInner, attrValue } from "./xml.js";

/**
 * Parse word/numbering.xml into numId -> { abstractNumId, overrides }.
 *
 * overrides is Map<ilvl:number, startVal:number>, from any <w:lvlOverride>
 * with a <w:startOverride> child — kept for informative messages, but
 * numbering-restart's core detection doesn't depend on it: a numId switch
 * to a *different* <w:num> that still points at the same abstractNumId is
 * itself the restart signal, override or not (each numId tracks its own
 * counter state in Word, starting fresh unless told otherwise).
 *
 * @param {string} xml - the whole word/numbering.xml part
 * @returns {Map<string, { abstractNumId: string, overrides: Map<number, number> }>}
 */
export function parseNumbering(xml) {
  const nums = new Map();
  if (!xml) return nums;

  const root = topLevelElements(xml).find((el) => el.name === "w:numbering");
  if (!root) return nums;
  const rootInner = elementInner(root.xml);
  if (!rootInner) return nums;

  for (const el of topLevelElements(rootInner)) {
    if (el.name !== "w:num") continue;
    const numId = attrValue(el.xml, "w:numId");
    if (!numId) continue;

    const inner = elementInner(el.xml) || "";
    const children = topLevelElements(inner);
    const abstractEl = children.find((c) => c.name === "w:abstractNumId");
    const abstractNumId = abstractEl ? attrValue(abstractEl.xml, "w:val") : undefined;

    const overrides = new Map();
    for (const lvlOverride of children.filter((c) => c.name === "w:lvlOverride")) {
      const ilvl = attrValue(lvlOverride.xml, "w:ilvl");
      const lvlInner = elementInner(lvlOverride.xml) || "";
      const startOverrideEl = topLevelElements(lvlInner).find((c) => c.name === "w:startOverride");
      if (startOverrideEl && ilvl !== undefined) {
        overrides.set(Number(ilvl), Number(attrValue(startOverrideEl.xml, "w:val")));
      }
    }

    nums.set(numId, { abstractNumId, overrides });
  }
  return nums;
}
