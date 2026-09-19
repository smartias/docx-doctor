# docx-doctor

Detect and repair structural defects in real-world Word (`.docx`) templates —
the numbering restarts, silently-dropped heading spacing, and stray trailing
blank pages that accumulate in a template that's been hand-edited in Word for
years.

## Why

Most docx tooling (docxtemplater, Aspose.Words, the OpenXML SDK) assumes you
already have a clean template and just want to fill it with data. In
practice, a firm's actual "standard template" file has a decade of
hand-edits, copy-pasted sections, and Word's own UI silently corrupting
structure along the way. docx-doctor's job is that different, earlier
problem: **diagnose and fix the template itself**, before you templatize or
mail-merge it with something else.

## Non-goals

- Not a templating/mail-merge engine. Use docxtemplater (or your own
  token-replacement) for filling — docx-doctor is a pre-processing step, not
  a replacement.
- Not a general-purpose docx editor. It only understands a specific, growing
  set of known defect patterns.
- No AI/LLM involved — this is deterministic OOXML structure analysis.

## Install

```
npm install -g docx-doctor
```

or run it without installing via `npx docx-doctor ...`.

## CLI

```
docx-doctor scan template.docx
docx-doctor scan template.docx --rules trailing-blank-pages,numbering-restart
docx-doctor scan template.docx --ci        # nonzero exit only on error-severity findings
docx-doctor fix template.docx -o fixed.docx
docx-doctor fix template.docx -o fixed.docx --rules heading-ppr-order
```

`--ci` is the intended CI-pipeline hook: it fails the build on a
error-severity finding (a defect docx-doctor is confident about) but not on a
warning (like split-run-risk, which is informational — see below), so it
doesn't block a build over something cosmetic.

## Library

```js
import { openDocx, scan, fix } from "docx-doctor";

const doc = await openDocx(buffer);
const { findings } = scan(doc);                          // read-only
const { doc: fixed, applied, skipped } = fix(doc, { rules: ["numbering-restart"] });
await fixed.save("out.docx");
```

## What it catches (v0.1.0)

| Rule | Severity | Fixes it? | What it catches |
|---|---|---|---|
| `trailing-blank-pages` | warn | yes | A run of 2+ empty paragraphs left at the end of the document body from repeated save/edit cycles. |
| `heading-pPr-order` | error | yes | Paragraph properties inserted *after* the paragraph-mark `<w:rPr>` inside `<w:pPr>` — Word silently ignores them (ECMA-376 §17.3.1.29), a common result of naive "append before `</w:pPr>`" template code. |
| `numbering-restart` | error | yes | A list that reads as one continuous list but silently restarts its count, because a paragraph got pointed at a second `<w:num>` entry that points at the same underlying list definition. |
| `split-run-risk` | warn | no (linter only) | A `{{TOKEN}}` or `${TOKEN}` placeholder whose text is split across multiple `<w:r>` runs — invisible on screen, but it'll silently defeat a naive find-replace on the raw XML. Detect-only: fixing it is a templating decision, not a structural repair. |

Every rule is built to never false-positive on a clean template — each has a
test proving zero findings on an undamaged fixture, not just that it catches
the defect it's looking for.

## Building a document from scratch (`docx-doctor/build`)

Everything above repairs an existing file. `docx-doctor/build` is the other
direction: a small API for constructing a *new* `.docx` where three of the
four defects above are structurally impossible to produce, not just
detectable afterward.

```js
import { createDocument, paragraph, token } from "docx-doctor/build";

const doc = createDocument();                          // a valid, empty .docx
const scope = doc.numberedList({ format: "decimal" });  // one list definition

doc.append(
  paragraph({ text: "Scope of Work", bold: true, spacing: { before: 240, after: 120 } }),
  paragraph({ text: "First deliverable", list: scope }),
  paragraph({ text: "Second deliverable", list: scope }),   // same handle -> continuous numbering, guaranteed
  paragraph({ children: ["The project is ", token("PROJECT_NAME"), ", due soon."] }),
);

const buffer = doc.build();   // runs docx-doctor's own scan() first; throws if it finds anything
```

- `paragraph()` always places the paragraph-mark `<w:rPr>` last inside
  `<w:pPr>`, regardless of what order you pass options in — the
  `heading-pPr-order` bug can't happen through this function.
- `numberedList()` returns a handle; reusing it keeps a list continuous, and
  getting a *new* list is a separate, deliberate call — the
  `numbering-restart` bug (an accidental second `<w:num>` duplicating a
  list) requires calling `numberedList()` again on purpose, not tripping
  into it by accident.
- `token()` always becomes one complete, unsplit run — the only way to
  actually guarantee `split-run-risk` can't happen, since that rule has no
  automatic fix.
- `trailing-blank-pages` isn't preventable by construction (nothing stops
  you from appending empty paragraphs on purpose) — `build()`'s self-scan
  is the backstop for that one. Pass `{ allowFindings: true }` to opt out.

## Status

v0.1.0. All four scan/fix rules are real (not stubs), with tests including a
CLI test that drives the actual subprocess, plus `docx-doctor/build` for
constructing new documents defect-free. See `ROADMAP.md` for what's
explicitly deferred (multi-section `sectPr`, `startOverride`-only restarts,
runs wrapped in `<w:hyperlink>`/`<w:ins>`/`<w:del>`) — each one is a
documented scope decision, not an unknown gap.

## License

MIT — see `LICENSE`.
