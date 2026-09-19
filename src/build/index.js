// docx-doctor/build — construct a NEW valid .docx instead of repairing an
// existing one. Each of the three repairable rules (trailing-blank-pages
// aside — see BuilderDocument#build) is made structurally impossible to
// produce through this API, rather than merely detectable afterward.
export { createDocument, BuilderDocument } from "./document.js";
export { paragraph, token } from "./paragraph.js";
