# Changelog

All notable changes to docx-review are documented here.

## [1.4.2] - 2026-03-09

### Added
- `--in-place` / `-i` mode for editing a document via a temp copy and replacing the source only after a fully successful run
- Multi-paragraph `replace`, `insert_after`, and `insert_before` operations when `replace` or `text` contains `\n`
- Optional `format` and `style` fields on changes:
  - `format: "markdown"` parses `**bold**` and `*italic*` into separate inserted runs
  - `style: "Heading1"` and similar overrides paragraph style on inserted paragraphs

### Fixed
- In `--in-place` mode, any failed edit now preserves the original source document instead of overwriting it with a partial result
- Manifests can now chain later edits against paragraphs created earlier in the same run
- Release builds now ship without trimming because trim-enabled single-file binaries were unstable at runtime

## [1.4.1] - 2026-03-05

### Fixed
- Superscript/subscript text now emits `[SUP]`/`[SUB]` markers in textconv (fixes false "missing space" edits on affiliations like `1Division`)
- Hyperlink display text no longer silently dropped during extraction
- Footnote/endnote references emit `[^N]` inline markers; content extracted to `FOOTNOTES`/`ENDNOTES` sections
- Tab characters emit `\t` instead of being silently dropped (fixes word-merging)
- `SimpleField`, `SdtRun`, and `SdtBlock` content now extracted (fixes missing text from fields and content controls)
- Hidden text (`w:vanish`) marked as `[HIDDEN]` so LLM avoids editing invisible content
- Small caps and all caps emit `[SC]`/`[CAPS]` markers
- `CarriageReturn` and `SymbolChar` elements handled in run text extraction

## [1.4.0] - 2026-03-05

### Added
- Whitespace-flexible text matching (`FlexIndexOf`): tries exact ordinal match first, falls back to regex that treats any whitespace run (including NBSP) as equivalent. Compiled regexes are cached for performance.
- `AcceptAllTrackedChanges`: accepts all existing tracked changes (insertions, deletions, moves) before applying new edits, aligning the editor's text view with the reader's "accepted" view. Raises change-apply success rate from ~31% to ~98% on documents with prior tracked changes.
- `--accept-existing` / `--no-accept-existing` CLI flag: controls whether existing tracked changes are accepted before applying new ones (default: accept). Use `--no-accept-existing` for multi-reviewer workflows where prior edits should be preserved.

### Changed
- `DeletedText` elements now use the actual matched text from the paragraph (preserving original whitespace) instead of the normalized find pattern.
- `FindTextInParagraphs` and `FindAnchorInParagraphs` use `FlexIndexOf` instead of ordinal `Contains`.

## [1.3.5] - 2026-03-04

### Added
- Release automation: `make release`, `make update-taps` targets for GitHub releases and Homebrew tap updates

## [1.3.4] - 2026-03-04

### Fixed
- Runs with multiple `<w:t>` elements (e.g. titles with `<w:br/>`) now extract all text correctly

## [1.3.3] - 2026-03-04

### Fixed
- Emit `\n` for `<w:br/>` line-break elements in text extraction so position calculations account for breaks within paragraphs
- Fix cascading tracked change failures by using accepted-text view for run mapping

## [1.3.2] - 2026-02-21

### Fixed
- Comment update op no longer injects automatic formatting — text is stored exactly as provided

## [1.3.1] - 2026-02-20

### Changed
- Aligned documentation with pandoc-first manuscript workflow (prefer `pandoc` for new drafts, `docx-review` for revision markup)

## [1.3.0] - 2026-02-18

### Added
- `--create` mode: generate new documents from bundled NIH template with optional populate manifest
- `--template <path>` flag for custom templates
- Embedded `templates/nih-standard.docx` as .NET resource (Arial 11pt, 0.75" margins, NIH section structure)
- AI agent skill definition (`skill/SKILL.md`) with workflow guidance and reference schemas
- GitHub Action for automatic `.docx` diff comments on PRs (`.github/workflows/docx-diff.yml`)

## [1.2.0] - 2026-02-14

### Added
- `--diff old.docx new.docx`: semantic document comparison across all layers
  - Body text: LCS-based paragraph alignment with word-level diff
  - Comments: added/removed/modified with author and anchor text
  - Tracked changes: new/resolved revisions between versions
  - Formatting: bold, italic, underline, font, size, color changes
  - Styles: paragraph style changes (e.g. Normal → Heading2)
  - Metadata: title, author, revision count, word count
- `--textconv`: git textconv driver producing normalized, diffable text
  - Inline formatting markers (`[B]`/`[I]`/`[U]`)
  - Tracked changes as `[-deleted-]`/`[+inserted+]`
  - Comments as `/* [Author] text */`
  - Tables as pipe-delimited rows
  - Images as `[IMG: filename (hash)]`
- `--git-setup`: prints `.gitattributes` and `.gitconfig` instructions
- New source files: `DocumentExtractor.cs`, `DocumentDiffer.cs`, `DiffModels.cs`, `TextConv.cs`

## [1.1.0] - 2026-02-14

### Added
- `--read` mode: extract full review state from `.docx` files
  - Paragraphs with styles and tracked changes (insertions, deletions)
  - Comments with anchor text, author, date
  - Document metadata (title, author, word count, revision)
  - Aggregated summary statistics
- JSON output (`--read --json`) and human-readable summary
- New models: `ReadResult`, `ParagraphInfo`, `TrackedChangeInfo`, `CommentInfo`, `DocumentMetadata`, `ReadSummary`

## [1.0.0] - 2026-02-14

### Added
- Initial release: CLI tool for adding tracked changes and comments to Word documents
- .NET 8 + Open XML SDK 3.3.0, ships as single ~12MB self-contained binary
- Tracked change types: `replace` (w:del + w:ins), `delete` (w:del), `insert_after` (w:ins), `insert_before` (w:ins)
- Comment anchoring with `CommentRangeStart`/`CommentRangeEnd` markers
- Multi-run text matching across split XML runs
- Formatting preservation via `RunProperties` cloning
- JSON manifest input (file or stdin pipe)
- CLI flags: `--author`, `--json`, `--dry-run`, `-o`/`--output`, `--version`
- Makefile: `build`, `install`, `all` (cross-compile), `docker`, `test`, `clean`
- GitHub Actions release workflow (macOS arm64/x64, Linux x64/arm64)
- Cross-platform: osx-arm64, osx-x64, linux-x64, linux-arm64
