---
name: docx-review
description: "Read, edit, create, and diff Word documents (.docx) with tracked changes and comments using the docx-review CLI v1.4.2 — a .NET 8 tool built on Microsoft's Open XML SDK. Ships as a single native binary (no runtime). Use when: (1) Adding tracked changes (replace, delete, insert) to a .docx, (2) Adding or updating anchored comments to a .docx, (3) Reading/extracting text, tracked changes, comments, and metadata from a .docx, (4) Diffing two .docx files semantically, (5) Creating new documents from templates, (6) Responding to peer reviewer comments with tracked revisions, (7) Proofreading or revising manuscripts with reviewable output, (8) Any task requiring valid tracked-change .docx output with proper w:del/w:ins markup that renders natively in Word."
---

# docx-review v1.4.2

CLI tool for Word document review: tracked changes, comments, read, create, diff, and git integration. Built on Microsoft's Open XML SDK — 100% compatible tracked changes and comments.

## Install

```bash
brew install drpedapati/tools/docx-review
```

Binary: `/opt/homebrew/bin/docx-review` (self-contained, no runtime)

Verify: `docx-review --version`

## Agentic Edit Workflow

Before building any manifest, **always read the document first** to build a mental model of its structure and style. Never blindly insert raw text from the user — adapt it to match the document.

### Step 1: Read and study the document

```bash
docx-review input.docx --read --json
```

Examine the output to understand:
- **Section headers**: Are they ALL CAPS? Numbered? What style (Heading1, Normal)?
- **Entry formatting**: How are items separated — periods, commas, em-dashes, tabs?
- **Paragraph styles in use**: Normal, ListParagraph, BodyTextIndent, etc.
- **Structural patterns**: Do entries use sub-lines? Date-prefixed? Grouped by year?

### Step 2: Adapt user content to match document conventions

When the user provides raw text to insert, transform it to match the document's voice:
- If the document uses periods between label and description, don't use em-dashes
- If headers are ALL CAPS without numbering, don't add numbers
- If entries are comma-separated phrases, don't use semicolons
- Match the level of detail and phrasing style of similar sections

### Step 3: Build manifest, validate, apply

```bash
# Always dry-run first
docx-review input.docx edits.json --dry-run --json

# Then apply
docx-review input.docx edits.json -o reviewed.docx --json
# or in-place:
docx-review input.docx edits.json --in-place --json
```

### Step 4: Verify

```bash
docx-review reviewed.docx --read --json
```

Confirm each inserted item is its own paragraph (not crammed into one) and the text reads naturally alongside existing content.

## Workflow Decision Tree

- **Reading/extracting content?** → `docx-review input.docx --read --json`
- **Adding tracked changes or comments?** → Read first (Step 1 above) → Build JSON manifest → validate → apply
- **Creating new document?** → `docx-review --create -o output.docx`
- **Comparing two versions?** → `docx-review --diff old.docx new.docx`
- **New clean manuscript from markdown?** → `pandoc manuscript.md -o manuscript.docx` (not docx-review)

## Modes

### Edit: Apply tracked changes and comments

Takes a `.docx` + JSON manifest, produces a reviewed `.docx` with proper OOXML markup.

```bash
docx-review input.docx edits.json -o reviewed.docx
docx-review input.docx edits.json -o reviewed.docx --json    # structured output
docx-review input.docx edits.json --dry-run --json           # validate without modifying
cat edits.json | docx-review input.docx -o reviewed.docx     # stdin pipe
docx-review input.docx edits.json -o reviewed.docx --author "Dr. Smith"
docx-review input.docx edits.json -o reviewed.docx --no-accept-existing  # preserve prior tracked changes
docx-review input.docx edits.json --in-place --json              # edit file in place (no separate output)
```

`--in-place` (`-i`) edits the input file directly. Mutually exclusive with `-o`. Safe: uses a temp file internally and only overwrites the original on success.

By default, existing tracked changes are accepted before applying new edits (`--accept-existing`). This aligns the text view with the reader's "accepted" view and dramatically improves match rates on previously-edited documents. Use `--no-accept-existing` for multi-reviewer workflows where prior edits must be preserved.

### Read: Extract document content as JSON

```bash
docx-review input.docx --read --json
```

Returns: paragraphs (with styles), tracked changes (type/text/author/date), comments (anchor text/content/author), metadata (title/author/word count/revision), and summary statistics. For output schema, see `references/read-schema.md`.

### Create: Generate new documents from template

```bash
docx-review --create -o new.docx                        # blank NIH template
docx-review --create -o new.docx populate.json          # populate template sections
docx-review --create -o new.docx --template custom.docx # use custom template
```

### Diff: Semantic comparison of two documents

```bash
docx-review --diff old.docx new.docx
docx-review --diff old.docx new.docx --json
```

Detects: text changes (word-level), formatting (bold/italic/font/color), comment modifications, tracked change differences, metadata changes, structural additions/removals.

### Git: Textconv driver for meaningful Word diffs

```bash
docx-review --textconv document.docx    # normalized text output
docx-review --git-setup                 # print .gitattributes/.gitconfig instructions
```

Textconv inline markers: `[B]`/`[I]`/`[U]` (formatting), `[-deleted-]`/`[+inserted+]` (tracked changes), `/* [Author] text */` (comments), `[SUP]`/`[SUB]` (super/subscript), `[SC]`/`[CAPS]` (small caps/all caps), `[HIDDEN]` (hidden text), `[^N]` (footnote/endnote refs), `[IMG: name (hash)]` (images), `\t` (tabs).

## JSON Manifest Format

Build this JSON, pass it to `docx-review`.

```json
{
  "author": "Reviewer Name",
  "changes": [
    { "type": "replace", "find": "exact text in document", "replace": "new text" },
    { "type": "delete", "find": "exact text to delete" },
    { "type": "insert_after", "anchor": "exact anchor text", "text": "text to insert after" },
    { "type": "insert_before", "anchor": "exact anchor text", "text": "text to insert before" }
  ],
  "comments": [
    { "anchor": "exact text to attach comment to", "text": "Comment content" },
    { "op": "update", "id": 12, "text": "Updated comment text" }
  ]
}
```

### Change types

| Type | Fields | Result in Word |
|------|--------|---------------|
| `replace` | `find`, `replace` | Red strikethrough old + blue new text |
| `delete` | `find` | Red strikethrough |
| `insert_after` | `anchor`, `text` | Blue inserted text after anchor |
| `insert_before` | `anchor`, `text` | Blue inserted text before anchor |

### Critical rules for `find` and `anchor` text

1. **Must be exact copy-paste from the document.** The tool tries exact ordinal match first, then falls back to whitespace-flexible matching (treats any whitespace run including NBSP as equivalent).
2. **Include enough context for uniqueness** — 15+ words when the phrase is common.
3. **First occurrence wins.** The tool replaces/anchors at the first match only.
4. Always validate with `--dry-run --json` before applying.

### Multi-paragraph insertions

Use `\n` in `text` or `replace` fields to create separate Word paragraphs. Each line becomes its own `<w:p>` element with proper tracked-change markup.

```json
{
  "type": "insert_after",
  "anchor": "Research (75%)",
  "text": "\nPROJECTS\nClinCognition. Gamified CME platform\nAutoCleanEEG. Python/MNE preprocessing pipeline"
}
```

Result: three new paragraphs after "Research (75%)" — the header and two entries, each as a separate paragraph inheriting the source paragraph's formatting.

### Concrete example

Given a document containing: *"The study enrolled 30 patients with moderate to severe symptoms over a 12-month period at three clinical sites."*

```json
{
  "author": "Dr. Smith",
  "changes": [
    {
      "type": "replace",
      "find": "The study enrolled 30 patients with moderate to severe symptoms over a 12-month period",
      "replace": "The study enrolled 45 patients with moderate to severe symptoms over an 18-month period"
    }
  ],
  "comments": [
    {
      "anchor": "three clinical sites",
      "text": "Please confirm the number of sites — the methods section mentions four."
    }
  ]
}
```

Note: `find` includes enough surrounding text (15+ words) for a unique match. The `anchor` for the comment uses the specific phrase being questioned.

## Helper Scripts

### `scripts/validate_manifest.sh`

Dry-run validation with human-readable pass/fail summary. Run before applying edits.

```bash
scripts/validate_manifest.sh manuscript.docx edits.json
# Output: 8/8 edits matched
```

### `scripts/review_pipeline.sh`

Full pipeline: validate → apply → report output path. Aborts on validation failure.

```bash
scripts/review_pipeline.sh manuscript.docx edits.json reviewed.docx
```

## JSON Output (--json)

```json
{
  "input": "paper.docx",
  "output": "paper_reviewed.docx",
  "author": "Dr. Smith",
  "changes_attempted": 5,
  "changes_succeeded": 5,
  "comments_attempted": 3,
  "comments_succeeded": 3,
  "success": true,
  "results": [
    { "index": 0, "type": "comment", "success": true, "message": "Comment added" },
    { "index": 0, "type": "replace", "success": true, "message": "Replaced" }
  ]
}
```

Exit code 0 = all succeeded. Exit code 1 = at least one failed (partial success possible).

## Key Behaviors

- **Existing tracked changes accepted by default.** Before applying new edits, all prior tracked changes are accepted so the text view matches the reader's view. Override with `--no-accept-existing`.
- **Whitespace-flexible matching.** Exact ordinal match tried first; if that fails, falls back to a regex that normalizes whitespace runs (spaces, NBSP, tabs). Compiled regexes are cached.
- **Comments applied first**, then tracked changes. Ensures anchors resolve before XML is modified.
- **Formatting preserved.** RunProperties cloned from source runs onto both deleted and inserted text.
- **Multi-run text matching.** Text spanning multiple XML `<w:r>` elements (common in previously edited documents) is found and handled correctly.
- **Multi-paragraph insertions.** `\n` in inserted/replaced text creates proper `<w:p>` paragraph elements (not inline line breaks). Each new paragraph gets its own `<w:ins>` wrapper and inherits paragraph properties from the source (excluding section properties).
- **In-place editing.** `--in-place` safely edits via temp file + atomic move. Original untouched on failure.
- **Everything untouched is preserved.** Images, charts, bibliographies, footnotes, cross-references, styles, headers/footers survive intact.

## Companion Tools

| Tool | Install | Purpose |
|------|---------|---------|
| `pptx-review` | `brew install drpedapati/tools/pptx-review` | PowerPoint read/edit |
| `xlsx-review` | `brew install drpedapati/tools/xlsx-review` | Excel read/edit |

Same architecture: .NET 8, Open XML SDK, single binary, JSON in/out.
