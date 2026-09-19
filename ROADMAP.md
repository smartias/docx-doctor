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

## Weekend 3 — heading-pPr-order ✅

- [x] `pPr` child-element parsing: `topLevelElements()`/`elementInner()`
      added to `src/xml.js` — generic depth-tracked "walk direct children
      of any XML fragment" helpers, not paragraph-specific, so Weekend 4
      (numbering.xml) can reuse them too.
- [x] `detect()`: flags any `pPr` with elements after the paragraph-mark
      `<w:rPr>` (ECMA-376 §17.3.1.29) — v0 scope is "is anything after
      rPr," not full CT_PPr schema ordering, since that's the actual
      defect naive append-before-close-tag code produces.
- [x] `repair()`: moves `rPr` to the end, preserves every other child's
      relative order (tested with 2 trailing elements, not just 1, to
      prove it doesn't accidentally just swap a pair).
- [x] Fixture (`fixtures/heading-ppr-order-docx.mjs`) with 5 paragraphs:
      2 broken (1 and 2 trailing elements), a correctly-ordered one, one
      with no `rPr` at all, and one with no `pPr` at all — the last three
      exist specifically to catch false positives.
- [x] Bug caught by the fixture during development, worth knowing about:
      `topLevelElements(para.xml)` on a *whole* `<w:p>...</w:p>` string
      returns `w:p` itself as the only top-level element — you have to
      unwrap with `elementInner()` first to see `pPr` as a child. Easy
      mistake to repeat when Weekend 4 walks `numbering.xml`.
- [x] 11/11 tests green.

## Weekend 4 — numbering-restart (the hard one) ✅

- [x] `src/numbering.js`: `parseNumbering()` maps `numId -> { abstractNumId,
      overrides }` from `word/numbering.xml`. Reused `topLevelElements()`/
      `elementInner()` from Weekend 3 without hitting the same unwrap
      mistake this time — the note in the Weekend 3 section paid off.
- [x] `detect()`: groups consecutive list paragraphs into runs, and within
      each run flags any block whose `numId` differs from the run's first
      ("canonical") `numId` but resolves to the *same* `abstractNumId` —
      a second `<w:num>` pointing at the same list definition, which is
      what actually causes Word to restart the count in practice. **Scope
      decision**: this fires whether or not there's an explicit
      `startOverride` — a numId switch alone is enough, since each numId
      tracks its own counter regardless. A `numId` switch to a
      *different* `abstractNumId` (a real list-type change) is never
      flagged — that's the false-positive guard.
- [x] `repair()`: retargets every paragraph in the flagged block back to
      the run's canonical `numId`, right-to-left so length-changing edits
      (e.g. numId "10" → "2") never invalidate not-yet-processed offsets.
- [x] Fixture (`fixtures/numbering-restart-docx.mjs`) with 2 abstractNums
      and 3 `num` entries: a canonical list, a restarted duplicate
      (same abstractNumId as canonical — this is the shape that actually
      matters), a genuinely different list type (must NOT be flagged),
      and a lone single-item "run" after a paragraph break (too short to
      compare against anything — must NOT be flagged). `fixtures/parts.mjs`
      gained `buildDocx({ documentXml, numberingXml })` so a fixture can
      override numbering.xml instead of only document.xml.
- [x] Explicit `startOverride`-only restarts (no numId change, just an
      override resetting an otherwise-identical numId) are NOT handled —
      genuinely deferred, not silently missed: real templates seen so far
      only exhibited the numId-duplication shape. Revisit if a real
      template surfaces the other one.
- [x] 14/14 tests green — first run, no debugging needed this time
      (Weekend 3's documented mistake was worth writing down).

## Weekend 5 — split-run-risk (linter, no fix) + CLI ✅

- [x] `detect()` only, no `repair()`: finds `{{TOKEN}}` / `${TOKEN}`-style
      placeholders whose text is split across multiple `<w:r>` runs, by
      concatenating each paragraph's run texts with offset tracking and
      checking whether a token match lands entirely within one run's span.
      v0 only recognizes those two bracket conventions and only walks
      `<w:r>` elements that are DIRECT children of the paragraph (a run
      wrapped in `<w:hyperlink>`/`<w:ins>`/`<w:del>` isn't seen) — both
      documented as deferred scope, not oversights.
- [x] `src/cli.js` is real now: `docx-doctor scan <file> [--rules a,b,c]
      [--ci]` and `docx-doctor fix <file> -o <out> [--rules a,b,c]`.
- [x] `--ci`: nonzero exit only when an **error**-severity finding is
      present — confirmed manually that a warn-only result (e.g.
      trailing-blank-pages) exits 0 even with `--ci`, so this is
      genuinely CI-adoptable without blocking on cosmetic findings.
- [x] All 4 built-in rules are implemented as of this weekend, so
      `scan()`/`fix()` with no `opts.rules` now runs cleanly against
      everything by default — the stale "pass opts.rules explicitly or
      it'll throw" caveat from Weekends 2-4 is gone (removed from
      `src/index.js`'s docstring).
- [x] `test/cli.test.mjs` drives the CLI as a real subprocess (not just
      calling its internals) against the existing fixtures — exit codes,
      `--ci` gating, and a fix-then-rescan-clean round trip.
- [x] 24/24 tests green.

## Weekend 6 — polish and ship v0.1.0

- [x] README rewritten with real CLI examples (`scan`/`fix`/`--ci`), a
      table of what each of the 4 rules catches, and an accurate status
      section — no more "nothing here works yet."
- [x] `docx-doctor` is unclaimed on npm (confirmed via `npm view` — 404).
      `package.json` filled in for publishing: version bumped to `0.1.0`,
      `files` whitelist (just `src/`, README, LICENSE — verified with
      `npm publish --dry-run`: 14 files, no `test/`/`fixtures/`/`.github/`
      leaking into the tarball), `repository`/`bugs`/`homepage`,
      `keywords`. `npm pkg fix` cleaned up a trivial `bin` path warning.
- [x] **Not done from here, deliberately**: the actual `npm publish` needs
      your own npm account (`npm login` first if you haven't). Once
      you're ready:
      ```
      npm login
      npm publish
      ```
      `--dry-run` output is already verified clean, so this should just
      work — but it's a real, public, hard-to-undo action, so it's yours
      to run, not something to script from a session that isn't logged in
      as you.
- [x] Write-up drafted: `WRITEUP.md` — "The four ways your Word template
      silently breaks," one section per rule, grounded in the real
      defects (not generic copy). Written in first person as a starting
      draft for you to edit and post wherever, not something to publish
      on your behalf.

## Explicitly deferred (not v0.1)

- Templating/fill logic (token replacement, list/block expansion) — stays
  out of scope; hand off to docxtemplater.
- Multi-format support (`.dotx`, legacy `.doc`) — OOXML `.docx` only for
  now.
- A hosted web UI — CLI + library first; see if anyone actually uses it
  before building that.
