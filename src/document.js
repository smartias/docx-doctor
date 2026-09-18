import { unzipParts, zipParts, partText, setPartText } from "./zip.js";

/** In-memory .docx: a Map of OOXML part paths to raw bytes, plus zip I/O. */
export class Document {
  /** @param {Map<string, Uint8Array>} parts */
  constructor(parts) {
    this.parts = parts;
  }

  /** @param {Uint8Array|Buffer} buffer */
  static open(buffer) {
    return new Document(unzipParts(buffer));
  }

  text(path) {
    return partText(this.parts, path);
  }

  setText(path, text) {
    setPartText(this.parts, path, text);
  }

  toBuffer() {
    return zipParts(this.parts);
  }

  async save(filePath) {
    const fs = await import("node:fs/promises");
    await fs.writeFile(filePath, this.toBuffer());
  }
}
