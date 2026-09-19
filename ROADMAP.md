# Roadmap

Each "weekend" below is a self-contained, shippable chunk — finish it,
commit, and the project is strictly further along even if you stop there.
Don't skip ahead: Weekend 4 depends on the round-trip guarantee from
Weekend 1 actually holding.

## Weekend 1 — read and write a .docx losslessly

- [x] Pick a zip library — went with `fflate` (small, no transitive deps).
      `openDocx(buffer)` unzips into `Document.parts: Map<path, Uint8Array>`
      (`src/zip.js`, `src/document.js`).
- [x] `doc.save()` / `doc.toBuffer()`: re-zip parts unchanged.
- [x] **Round-trip test** (`test/round-trip.test.mjs`): open → save →
      re-open reproduces every part byte-for-byte. Green.
- [x] Fixture: a hand-authored, spec-correct minimal `.docx`
      (`fixtures/minimal-docx.mjs`) — heading, body paragraph, a 2-item
      numbered list, built in code so it's readable/diffable in git.
- [ ] **Do this manually before Weekend 2**: LibreOffice headless
      conversion isn't working in the dev sandbox this was scaffolded in
      (`soffice --convert-to` fails on even a plain .txt — looks like a
      sandbox restriction, not a docx-doctor problem). Open
      `fixtures/minimal-docx.mjs`'s output in real Word or LibreOffice on
      your own machine once, by hand, to confirm it's genuinely valid
      before trusting it as the base for damaged-template fixtures later.
      (Quick way to get the bytes: add a one-off script that calls
      `buildMinimalDocx()` and writes the result to a `.docx` file.)

## Weekend 2 — first real rule: trailing-blank-pages ✅

- [x] Paragraph-walking over `word/document.xml` (`src/xml.js`:
      `findParagraphs`, `paragraphText`, `isEmptyParagraph`).
- [x] `trailing-blank-pages.detect()`: finds a run of 2+ empty paragraphs
      at the end of the body. Deliberately does *not* flag a single
      trailing empty paragraph — that's normal, not a defect — so it
      doesn't false-positive on every clean template.
- [x] `trailing-blank-pages.repair()`: removes them, with a safety rule
      that never empties the body entirely (leaves the first paragraph
      behind if the whole body turns out to be empty paragraphs).
- [x] Fixture (`fixtures/trailing-blanks-docx.mjs`) with 3 trailing empty
      paragraphs (a self-closing one, a whitespace-only run, another
      self-closing one) — mirrors real damage, not a toy case.
      `fixtures/parts.mjs` now holds the shared boilerplate so clean vs.
      damaged fixtures only differ in `document.xml`.
- [x] `scan()`/`fix()` wired up in `src/index.js` for real. Note: they
      still throw if you don't pass `opts.rules` — only
      trailing-blank-pages is implemented so far; see Weekends 3-5.
- [x] 8/8 tests green, including a "clean fixture produces zero findings"
      test (the false-positive check that matters most for this kind of
      tool) and a degenerate all-empty-body case for the repair safety rule.

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
