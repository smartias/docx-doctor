# Roadmap

Each "weekend" below is a self-contained, shippable chunk — finish it,
commit, and the project is strictly further along even if you stop there.
Don't skip ahead: Weekend 4 depends on the round-trip guarantee from
Weekend 1 actually holding.

## Weekend 1 — read and write a .docx losslessly

- [ ] Pick a zip library (`fflate` is small and dependency-free; `jszip` is
      more common and higher-level). Implement `openDocx(buffer)`: unzip,
      expose parts as a `Map<path, string|Buffer>`.
- [ ] Implement `doc.save()` / `doc.toBuffer()`: re-zip parts unchanged.
- [ ] **Round-trip test**: open a real (or synthetic) `.docx`, save it
      immediately, verify Word/LibreOffice still opens the result and
      nothing changed. This is the most important test in the whole
      project — get it green before writing a single rule.
- [ ] Build 1-2 synthetic, undamaged fixture `.docx` files by hand (a
      minimal doc saved via Word or LibreOffice) to round-trip against.

## Weekend 2 — first real rule: trailing-blank-pages

- [ ] Implement paragraph-walking over `word/document.xml` (find `<w:p>`
      elements, detect "empty" — no runs with visible text).
- [ ] `trailing-blank-pages.detect()`: find a run of empty paragraphs
      immediately before the final `sectPr`.
- [ ] `trailing-blank-pages.repair()`: remove them.
- [ ] Build a synthetic fixture with 2-3 trailing blank paragraphs; test
      that `detect()` finds them and `repair()` removes them without
      touching anything else.
- [ ] Wire up `scan()` and `fix()` in `index.js` to actually run rules
      against a doc (just this one rule for now).

## Weekend 3 — heading-pPr-order

- [ ] Implement `pPr` child-element parsing aware of the paragraph-mark
      `<w:rPr>` position (ECMA-376 §17.3.1.29 ordering rules).
- [ ] `detect()`: flag any `pPr` where a spacing/numbering element appears
      *after* the paragraph-mark `rPr` — Word silently ignores it, which is
      the concrete, testable failure mode.
- [ ] `repair()`: reorder elements into valid position (an insert-before-rPr
      helper, not append-before-close-tag).
- [ ] Fixture + test proving the Word-visible behavior: a paragraph whose
      spacing Word would ignore before the fix, correctly ordered after.

## Weekend 4 — numbering-restart (the hard one)

- [ ] Parse `word/numbering.xml`: the `abstractNum` / `num` mapping.
- [ ] `detect()`: find lists that restart numbering (`numId` change or
      `startOverride`) mid-section where the surrounding paragraphs are
      stylistically the same list.
- [ ] `repair()`: point the restarted paragraphs back at the original
      `numId`.
- [ ] Fixture covering at least 2 restart shapes (new `numId` vs.
      `startOverride`).
- [ ] This rule will likely need a second pass later — real templates vary
      a lot here. Don't aim for perfect; aim for "flags the obvious cases,
      never false-positives on a clean list."

## Weekend 5 — split-run-risk (linter, no fix) + CLI

- [ ] `detect()` only: find tokens/text likely split across multiple
      `<w:r>` runs (the thing that breaks naive find-replace). No
      `repair()` — recommend handing this to a templating tool instead.
- [ ] Build `src/cli.js` for real: `docx-doctor scan <file> [--ci]`,
      `docx-doctor fix <file> -o <out> [--rules a,b,c]`.
- [ ] `--ci` flag: nonzero exit code if any `severity: "error"` finding is
      present — this is what makes it adoptable in someone else's CI
      pipeline, which is the actual distribution wedge.

## Weekend 6 — polish and ship v0.1.0

- [ ] README examples using the CLI, not just the API.
- [ ] `npm publish` as `docx-doctor` (check name availability first) or a
      scoped alternative.
- [ ] One real write-up: "the four ways your Word template silently
      breaks." This is the actual distribution mechanism — more than the
      npm listing itself.

## Explicitly deferred (not v0.1)

- Templating/fill logic (token replacement, list/block expansion) — stays
  out of scope; hand off to docxtemplater.
- Multi-format support (`.dotx`, legacy `.doc`) — OOXML `.docx` only for
  now.
- A hosted web UI — CLI + library first; see if anyone actually uses it
  before building that.
