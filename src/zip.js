import { unzipSync, zipSync, strToU8, strFromU8 } from "fflate";

/** Unzip a .docx buffer into a Map<partPath, Uint8Array>. */
export function unzipParts(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const raw = unzipSync(bytes);
  return new Map(Object.entries(raw));
}

/** Zip a Map<partPath, Uint8Array|string> back into a .docx buffer. */
export function zipParts(parts) {
  const files = {};
  for (const [path, value] of parts) {
    files[path] = value instanceof Uint8Array ? value : strToU8(value);
  }
  return zipSync(files, { level: 6 });
}

export function partText(parts, path) {
  const bytes = parts.get(path);
  return bytes === undefined ? undefined : strFromU8(bytes);
}

export function setPartText(parts, path, text) {
  parts.set(path, strToU8(text));
}
