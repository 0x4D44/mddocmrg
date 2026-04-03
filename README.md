# mddocmrg

A Rust CLI tool that extracts plain text from Microsoft Word (`.docx`) files and merges them into a single output file.

Useful for consolidating documentation or extracting text content from batches of Word documents for further processing or analysis.

## Features

- **Text extraction** from `.docx` files by parsing the internal `word/document.xml`
- **Batch processing** via glob patterns (e.g., `*.docx`, `docs/**/*.docx`)
- **Hyperlink cleaning** with `--strip-hyperlinks` to remove `HYPERLINK` field codes while preserving visible link text
- Merged output separated by double newlines

## Installation

Requires [Rust](https://www.rust-lang.org/tools/install).

```bash
cargo build --release
```

The binary is produced at `target/release/mddocmrg`.

## Usage

```
mddocmrg [options] <file_pattern1> <file_pattern2> ...
```

### Options

| Flag | Description |
|------|-------------|
| `-h`, `-?` | Display help and exit |
| `-s`, `--strip-hyperlinks` | Remove hyperlink field instructions from output |

### Examples

```bash
# Merge all DOCX files in the current directory
mddocmrg "*.docx"

# Merge specific files with hyperlink stripping
mddocmrg -s Chapter1.docx Chapter2.docx

# Merge files from a subdirectory
mddocmrg "docs/*.docx"
```

Output is written to `merged.txt` in the current working directory.

## How It Works

`.docx` files are ZIP archives containing Office Open XML. The tool opens each file as a ZIP, reads `word/document.xml`, and walks the XML tree with `quick-xml` to extract text nodes (`<w:t>` elements). When `--strip-hyperlinks` is enabled, text inside `<w:instrText>` elements (which contain field codes like `HYPERLINK "https://..."`) is skipped, preserving only the visible link text.

## Dependencies

| Crate | Purpose |
|-------|---------|
| [zip](https://crates.io/crates/zip) 0.6 | Read `.docx` ZIP archives |
| [quick-xml](https://crates.io/crates/quick-xml) 0.27 | Parse Office Open XML |
| [glob](https://crates.io/crates/glob) 0.3 | Expand file path patterns |

## Testing

```bash
cargo test
```

Tests create temporary `.docx` files in-memory (using `tempfile`) covering extraction, merging, hyperlink stripping, error handling, and CLI argument parsing.
