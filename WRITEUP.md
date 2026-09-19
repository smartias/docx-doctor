<!--
Draft for a blog post / dev.to / wherever. Written in your voice as the
author — edit freely, this is a starting point, not a finished piece.
-->

# The four ways your Word template silently breaks

If you've ever generated a Word document from a template — a proposal, a
contract, a report — you've probably hit a moment where the output looks
subtly wrong and you can't figure out why. The numbering restarts partway
through a list that should be continuous. A heading's spacing just... isn't
applied, even though you can see the setting in the XML. A stray blank page
shows up at the end of every document, no matter how many times someone
"fixes" it in Word.

None of these are bugs in Word, and none of them are bugs in whatever
templating tool you're using to fill the document. They're structural
defects that have been silently baked into the template file itself, often
years before anyone noticed — and because they're silent, they survive
every round of manual "fix it in Word" edits, because a human editing in
Word can't see the XML that's actually broken.

Here are the four I've run into most.

## 1. Trailing blank pages that keep coming back

Every time a document gets saved, edited, and saved again, there's a good
chance an empty paragraph or two gets left behind at the end — someone
pressed enter once too many, or a section got deleted but its terminal
paragraph mark didn't. Individually harmless. Accumulated over a few years
of template maintenance, you get two or three blank pages tacked onto the
end of every document generated from it, and "just delete the blank
paragraphs" doesn't stick because the next person who edits the template
adds a couple more without noticing.

The fix is mechanical once you know what to look for: a run of two or more
empty `<w:p>` elements immediately before the body's closing `<w:sectPr>`.
(One trailing empty paragraph is normal — every section conventionally ends
with one. It's a *run* of them that's the defect.)

## 2. Paragraph properties that Word quietly ignores

This one's the most surprising once you see it. Inside a paragraph's
`<w:pPr>` element, the paragraph-mark run properties (`<w:rPr>`) have to be
the *last* child, per the OOXML spec (ECMA-376 §17.3.1.29). If anything —
spacing, indentation, a numbering reference — comes after it, Word doesn't
error. It doesn't warn. It just ignores that element entirely, as if it were
never there.

This happens constantly in practice because the natural way to "add a pPr
setting" programmatically is to insert it right before `</w:pPr>` — and that
silently breaks the moment an `<w:rPr>` is already present, which it usually
is. You end up with a heading style that visibly has the right spacing
value in the XML, and visibly doesn't apply it in Word, and there's no error
message anywhere pointing at why.

## 3. Lists that restart for no visible reason

A numbered list in a Word document doesn't carry its own counter — it
references a `<w:num>` entry in `numbering.xml`, and that `<w:num>` entry
points at an abstract list definition (`<w:abstractNum>`) that actually
defines the numbering format and start value. Word tracks the running count
*per `<w:num>` entry*, not per abstract definition.

So here's the failure mode: someone copies a numbered paragraph from one
part of the template to another — maybe via copy-paste, maybe Word's own
"repair" behavior on document merge — and Word (or the person's editing
tool) creates a *second* `<w:num>` entry pointing at the same abstract list.
Visually, nothing looks different. The list still looks like one continuous
numbered list in the editor. But because it's now split across two `<w:num>`
references, the count resets in the middle, and nobody can figure out why
item 8 is suddenly item 1 again.

## 4. Merge tokens that vanish for no reason

If you're generating documents by find-and-replacing tokens like
`{{PROJECT_NAME}}` directly against the raw XML, you'll eventually hit a
token that just doesn't get replaced — even though it's clearly sitting
right there in the document when you open it in Word.

The reason: Word splits a paragraph's visible text across multiple `<w:r>`
run elements for reasons that have nothing to do with what's on screen —
spell-check markers, revision tracking IDs, just the accumulated history of
edits. So `{{PROJECT_NAME}}` might actually be stored as `{{PROJECT_` in one
run and `NAME}}` in the very next one. On screen, one word. In the XML, two
separate `<w:t>` elements that a naive string search across a single run
will never match.

This one isn't really fixable as a "repair" — merging runs back together is
a formatting decision, not a structural correction — but it's absolutely
worth *detecting* before it silently eats a field in production.

---

All four of these share the same shape: invisible in the Word UI, silent
(no error, no warning) when they cause a problem, and they accumulate
because nobody can see them to fix them permanently. That's the actual gap
— not a shortage of tools that can *fill* a Word template, but a shortage of
tools that can tell you a template is already damaged before you build
anything on top of it.

That's what [docx-doctor](https://github.com/smartias/docx-doctor) does —
`docx-doctor scan your-template.docx` finds these four patterns (and is
built to grow more), with `--ci` support so a broken template fails a build
instead of shipping. It's early — four rules, MIT licensed, looking for
more real damaged templates to test against. If you've got one, that's the
most useful thing you could send.
