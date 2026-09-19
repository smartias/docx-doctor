import { paragraph } from "./paragraph.js";

const DEFAULT_TABLE_WIDTH_TWIPS = 9000; // ~6.25in — fits a standard letter page with 1" margins

const BORDERS =
  "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join("") +
  "</w:tblBorders>";

function toParagraphXml(item) {
  if (typeof item === "string") return item.startsWith("<w:p") ? item : paragraph({ text: item });
  return item;
}

function cellXml(cell, width) {
  const body = Array.isArray(cell) ? cell.map(toParagraphXml).join("") : toParagraphXml(cell);
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${body || paragraph({ text: "" })}</w:tc>`;
}

/**
 * Build a <w:tbl>. `rows` is an array of arrays; each cell is plain text
 * (wrapped in a simple paragraph), a pre-built paragraph() string (for
 * bold/spacing/etc. inside a cell), or an array of either for a
 * multi-paragraph cell.
 *
 * KNOWN LIMITATION (see ROADMAP.md Weekend 8, verified with a test, not
 * just asserted): the scan rules find paragraphs with a whole-document
 * text scan that doesn't distinguish a body-level paragraph from one
 * nested inside a table cell. trailing-blank-pages only flags a run of 2+
 * consecutive empty paragraphs, so in practice a SINGLE empty trailing
 * cell is fine (its non-empty neighbor breaks the run) — it takes 2+
 * consecutive empty cells at the very end of the table (e.g. an entirely
 * empty last row) to trip the check in build(), even though nothing is
 * actually wrong. Workaround, and also just normal Word convention
 * regardless of docx-doctor: append a real paragraph after a table if
 * it's the last thing in the document.
 *
 * @param {{ rows: Array<Array<string|string[]>>, columnWidths?: number[] }} opts
 * @returns {string} a complete "<w:tbl>...</w:tbl>" element
 */
export function table({ rows, columnWidths } = {}) {
  if (!rows || rows.length === 0) throw new Error("table(): rows must be a non-empty array");
  const colCount = Math.max(...rows.map((r) => r.length));
  const widths = columnWidths && columnWidths.length === colCount
    ? columnWidths
    : Array(colCount).fill(Math.floor(DEFAULT_TABLE_WIDTH_TWIPS / colCount));

  const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const trs = rows.map((row) => {
    const tcs = row.map((cell, i) => cellXml(cell, widths[i] ?? widths[widths.length - 1])).join("");
    return `<w:tr>${tcs}</w:tr>`;
  }).join("");

  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${trs}</w:tbl>`;
}
