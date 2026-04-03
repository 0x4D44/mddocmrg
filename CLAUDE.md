# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

`mddocmrg` is a Rust CLI tool that extracts plain text from `.docx` files (ZIP archives containing `word/document.xml`) and merges them into a single output file. It supports glob patterns for batch processing and optional hyperlink field-code stripping.

## Build & Test Commands

```bash
cargo build --release          # Build release binary
cargo test                     # Run all tests (unit + integration)
cargo test <test_name>         # Run a single test
cargo llvm-cov                 # Code coverage
```

## Architecture

Single-file implementation in `src/main.rs` with three public functions forming the pipeline:

1. **`extract_text_from_docx(path, strip_hyperlinks)`** - Opens a `.docx` as a ZIP, reads `word/document.xml`, parses XML with `quick-xml`, extracts text nodes (optionally skipping `w:instrText` elements containing HYPERLINK field codes).
2. **`merge_docx_files(paths, strip_hyperlinks)`** - Calls `extract_text_from_docx` for each path, concatenates results separated by `\n\n`.
3. **`run(args, output_path)`** - CLI entry point: parses args, expands glob patterns, calls `merge_docx_files`, writes result to `output_path`.

`main()` calls `run()` with `"merged.txt"` as the hardcoded output path. The `run()` function accepts `output_path` as a parameter to enable testing without filesystem side effects.

Integration tests in `tests/cli_integration.rs` exercise the binary via `cargo run`.

## Dependencies

- `zip` (0.6) - Read `.docx` ZIP archives
- `quick-xml` (0.27) - Parse Office Open XML
- `glob` (0.3) - Expand file patterns
- `tempfile` (dev) - Temporary dirs for tests
