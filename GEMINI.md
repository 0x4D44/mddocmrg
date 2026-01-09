# mddocmrg

## Project Overview
`mddocmrg` (Markdown/Document Merge) is a Rust-based Command Line Interface (CLI) utility designed to extract plain text from multiple Microsoft Word (`.docx`) files and merge them into a single output file (`merged.txt`).

It is particularly useful for consolidating documentation or extracting text content from batches of Word documents for further processing or analysis.

## Key Features
*   **Text Extraction:** parsing `word/document.xml` within `.docx` archives to extract visible text.
*   **Batch Processing:** Supports glob patterns (e.g., `*.docx`, `docs/**/*.docx`) to process multiple files in one go.
*   **Merging:** Concatenates extracted text from all matched files, separating entries with double newlines.
*   **Hyperlink Cleaning:** Optional flag `--strip-hyperlinks` (or `-s`) to remove underlying `HYPERLINK` field codes/instructions while preserving the visible link text.

## Tech Stack
*   **Language:** Rust (Edition 2021)
*   **Core Dependencies:**
    *   `zip`: For reading `.docx` files (which are ZIP archives).
    *   `quick-xml`: For efficient parsing of the internal XML structure.
    *   `glob`: For shell-style pattern matching of file paths.

## Installation & Build
Ensure you have Rust and Cargo installed.

```bash
# Build the project
cargo build --release

# Run the binary directly
./target/release/mddocmrg --help
```

## Usage
The general syntax is:
```bash
cargo run -- [options] <file_pattern1> <file_pattern2> ...
```

### Options
*   `-h`, `-?`: Display help message.
*   `-s`, `--strip-hyperlinks`: Remove hyperlink field instructions (e.g., `HYPERLINK "http://..."`) from the output, keeping only the visible text.

### Examples
```bash
# Merge all DOCX files in the current directory
cargo run -- "*.docx"

# Merge specific files and strip hyperlink codes
cargo run -- -s "Chapter1.docx" "Chapter2.docx"

# Merge files from a subdirectory
cargo run -- "docs/*.docx"
```

**Note:** The output is always written to a file named `merged.txt` in the current working directory.

## Development
### Testing
The project achieves >98% logic coverage with a comprehensive test suite.
```bash
cargo test
```

### Code Structure
*   `src/main.rs`: Contains the entire implementation, refactored for testability:
    *   `run`: The main entry point for logic, taking arguments and an output path.
    *   `extract_text_from_docx`: Logic to unzip and parse XML.
    *   `merge_docx_files`: Orchestrates processing of multiple paths.
    *   `tests` module: Comprehensive unit and integration tests covering CLI parsing, extraction workflows, and error handling.
*   `wrk_docs/`: Contains detailed code coverage reports and improvement plans.