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

## Status

Pre-alpha. See `ROADMAP.md` — this is a from-scratch build; nothing here
works yet, it's stubbed to the planned shape.

## Planned API

```js
import { openDocx, scan, fix } from "docx-doctor";

const doc = await openDocx(buffer);
const report = scan(doc);                 // report.findings: Finding[]
const { doc: fixed } = fix(doc, { rules: ["numbering-restart"] });
await fixed.save("out.docx");
```

```
npx docx-doctor scan template.docx --ci
npx docx-doctor fix template.docx -o fixed.docx --rules numbering-restart,trailing-blank-pages
```

## License

MIT — see `LICENSE`.
