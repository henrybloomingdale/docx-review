# docx-review

A CLI tool that adds **tracked changes** and **comments** to Word (.docx) documents using Microsoft's [Open XML SDK](https://github.com/dotnet/Open-XML-SDK). Takes a `.docx` file and a JSON edit manifest, produces a reviewed document with proper `w:del`/`w:ins` markup and comment anchors that render natively in Microsoft Word — no macros, no compatibility issues.

**Ships as a single 12MB native binary.** No runtime, no Docker required.

## Why Open XML SDK?

We evaluated three approaches for programmatic document review:

| Approach | Tracked Changes | Comments | Formatting |
|----------|:-:|:-:|:-:|
| **Open XML SDK (.NET)** | ✅ 100% | ✅ 100% | ✅ Preserved |
| python-docx / docx-editor | ✅ 100% | ⚠️ ~80% | ✅ Preserved |
| pandoc + Lua filters | ❌ Lossy | ❌ Limited | ⚠️ Degraded |

Open XML SDK is the gold standard — it's Microsoft's own library for manipulating Office documents. Comments anchor correctly 100% of the time, tracked changes use proper revision markup, and formatting is always preserved.

## Quick Start

### Option 1: Homebrew (recommended)

```bash
brew install henrybloomingdale/tools/docx-review
```

### Option 2: Native Binary

```bash
git clone https://github.com/henrybloomingdale/docx-review.git
cd docx-review
make install    # Builds + installs to /usr/local/bin
```

Requires [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) for building (`brew install dotnet@8`). The resulting binary is self-contained — no .NET runtime needed to run it.

### Option 3: Docker

```bash
make docker     # Builds Docker image
docker run --rm -v "$(pwd):/work" -w /work docx-review input.docx edits.json -o reviewed.docx
```

### Usage

```bash
# Basic usage
docx-review input.docx edits.json -o reviewed.docx

# Pipe JSON from stdin
cat edits.json | docx-review input.docx -o reviewed.docx

# Custom author name
docx-review input.docx edits.json -o reviewed.docx --author "Dr. Smith"

# Dry run (validate without modifying)
docx-review input.docx edits.json --dry-run

# JSON output for pipelines
docx-review input.docx edits.json -o reviewed.docx --json
```

## JSON Manifest Format

```json
{
  "author": "Reviewer Name",
  "changes": [
    {
      "type": "replace",
      "find": "original text in the document",
      "replace": "replacement text with tracked change"
    },
    {
      "type": "delete",
      "find": "text to mark as deleted"
    },
    {
      "type": "insert_after",
      "anchor": "text to find",
      "text": "new text inserted after the anchor"
    },
    {
      "type": "insert_before",
      "anchor": "text to find",
      "text": "new text inserted before the anchor"
    }
  ],
  "comments": [
    {
      "anchor": "text to anchor the comment to",
      "text": "Comment content displayed in the margin"
    }
  ]
}
```

### Change Types

| Type | Required Fields | Description |
|------|----------------|-------------|
| `replace` | `find`, `replace` | Finds text and creates a tracked replacement (w:del + w:ins) |
| `delete` | `find` | Finds text and marks as deleted (w:del only) |
| `insert_after` | `anchor`, `text` | Finds anchor text, inserts new text after it (w:ins) |
| `insert_before` | `anchor`, `text` | Finds anchor text, inserts new text before it (w:ins) |

### Comment Format

Each comment needs:
- `anchor` — text in the document to attach the comment to (CommentRangeStart/End markers)
- `text` — the comment content shown in Word's review pane

## Semantic Diff & Git Integration

Compare two `.docx` files semantically — detects text changes, formatting differences, comment and tracked change modifications, and metadata changes.

### Quick Start

```bash
# Compare two documents
docx-review --diff old.docx new.docx

# JSON output for automation
docx-review --diff old.docx new.docx --json

# Use as a git textconv driver
docx-review --textconv document.docx
```

### Git Integration

Track `.docx` changes in git with human-readable diffs:

```bash
# Print setup instructions
docx-review --git-setup
```

This tells you to add to `.gitattributes`:
```
*.docx diff=docx
```

And to `.gitconfig` (global or per-repo):
```ini
[diff "docx"]
    textconv = docx-review --textconv
```

Now `git diff` and `git log -p` show readable text diffs for Word documents instead of binary gibberish.

### What the Diff Detects

| Category | Details |
|----------|---------|
| **Text changes** | Word-level insertions, deletions, and modifications within matched paragraphs |
| **Formatting** | Bold, italic, underline, strikethrough, font family, font size, color changes |
| **Comments** | Added, removed, or modified comments (matched by author + anchor text) |
| **Tracked changes** | Differences in existing tracked changes between versions |
| **Metadata** | Title, author, description, last modified, revision count |
| **Structure** | Paragraph additions and removals (LCS-based alignment, Jaccard ≥ 0.5) |

### TextConv Output Format

The `--textconv` driver normalizes documents to readable text:

```
[B]bold text[/B]  [I]italic[/I]  [U]underlined[/U]
[-deleted text-]  [+inserted text+]
/* [Author] comment text */
| Cell 1 | Cell 2 | Cell 3 |
[IMG: figure1.png (sha256: abc123...)]
```

## CLI Flags

| Flag | Description |
|------|-------------|
| `-o`, `--output <path>` | Output file path (default: `<input>_reviewed.docx`) |
| `--author <name>` | Reviewer name for tracked changes (overrides manifest `author`) |
| `--json` | Output results as JSON (for scripting/pipelines) |
| `--dry-run` | Validate the manifest without modifying the document |
| `--read` | Extract document content (tracked changes, comments, metadata) |
| `--diff` | Compare two documents semantically |
| `--textconv` | Git textconv driver (normalized text output) |
| `--git-setup` | Print `.gitattributes` and `.gitconfig` setup instructions |
| `-v`, `--version` | Show version |
| `-h`, `--help` | Show help |

## Build Targets

```
make              # Build native binary for current platform (12MB, self-contained)
make install      # Build + install to /usr/local/bin
make all          # Cross-compile for macOS ARM64, macOS x64, Linux x64
make docker       # Build Docker image
make test         # Run test (requires TEST_DOC=path/to/doc.docx)
make clean        # Remove build artifacts
make help         # Show all targets
```

## Exit Codes

- `0` — All changes and comments applied successfully
- `1` — One or more edits failed (partial success possible)

## JSON Output Mode

With `--json`, the tool outputs structured results:

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

## How It Works

1. Copies the input `.docx` to the output path
2. Opens the document using Open XML SDK
3. Adds **comments first** (before tracked changes modify the XML tree)
4. Applies tracked changes (replace → w:del + w:ins, delete → w:del, insert → w:ins)
5. Handles multi-run text matching (text spanning multiple XML runs)
6. Preserves original run formatting (RunProperties cloned from source)
7. Saves and reports results

## Development

```bash
# Build native binary (requires .NET 8 SDK)
make build

# Build and run locally
dotnet run -- input.docx edits.json -o reviewed.docx

# Cross-compile all platforms
make all
# → build/osx-arm64/docx-review  (macOS Apple Silicon)
# → build/osx-x64/docx-review    (macOS Intel)
# → build/linux-x64/docx-review  (Linux)
```

## License

MIT — see [LICENSE](LICENSE).

---

*Built by [CinciNeuro](https://github.com/henrybloomingdale) for AI-assisted manuscript review workflows.*
