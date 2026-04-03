use colored::Colorize;
use glob::glob;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::fs::{self, File};
use std::io::Read;
use std::path::Path;
use walkdir::WalkDir;
use zip::read::ZipArchive;

const DEFAULT_TYPES: &[&str] = &["md", "docx"];

/// Prints the startup banner with name, version, and build date.
fn print_banner() {
    println!(
        "{} {} (built {})",
        env!("CARGO_PKG_NAME").bright_cyan().bold(),
        format!("v{}", env!("CARGO_PKG_VERSION")).bright_yellow(),
        env!("BUILD_DATE")
    );
}

/// Prints usage instructions.
fn print_usage(program: &str) {
    let prog_name = Path::new(program)
        .file_name()
        .map(|s| s.to_string_lossy())
        .unwrap_or_else(|| "mddocmrg".into());
    println!("Usage: {} [options] [<pattern> ...]", prog_name);
    println!();
    println!("With no patterns, scans the current directory tree for .md and .docx files.");
    println!("With patterns, merges files matching those glob patterns.");
    println!();
    println!("Options:");
    println!("  -h, -?                  Display this help message and exit");
    println!(
        "  -s, --strip-hyperlinks  Remove hyperlink field instructions from DOCX output"
    );
    println!("  -o <file>               Output file (default: merged.txt)");
    println!(
        "  --no-recurse            Only scan the current directory, not subdirectories"
    );
    println!(
        "  -t, --types <types>     Comma-separated file types to include (default: md,docx)"
    );
}

/// Extracts text from a markdown file (plain text pass-through).
pub fn extract_text_from_md(path: &str) -> Result<String, Box<dyn std::error::Error>> {
    let content = fs::read_to_string(path)?;
    Ok(content.trim().to_string())
}

/// Extracts the text content from the provided DOCX file.
/// If `strip_hyperlinks` is true, any field instruction text (inside <w:instrText>)
/// that starts with "HYPERLINK" is skipped. This generally removes the hyperlink's
/// underlying field code while keeping the visible text.
pub fn extract_text_from_docx(
    path: &str,
    strip_hyperlinks: bool,
) -> Result<String, Box<dyn std::error::Error>> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file)?;
    let mut document_xml = archive.by_name("word/document.xml")?;
    let mut xml_content = String::new();
    document_xml.read_to_string(&mut xml_content)?;

    let mut reader = Reader::from_str(&xml_content);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut text = String::new();

    // Track whether we're inside a hyperlink field instruction.
    let mut in_instr_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.name().as_ref() == b"w:instrText" {
                    in_instr_text = true;
                }
            }
            Ok(Event::End(ref e)) => {
                if e.name().as_ref() == b"w:instrText" {
                    in_instr_text = false;
                }
            }
            Ok(Event::Text(e)) => {
                // If stripping hyperlinks and we're in an instruction text element,
                // skip appending this text.
                if strip_hyperlinks && in_instr_text {
                    // Skip this text.
                } else {
                    text.push_str(&e.unescape()?);
                    text.push(' ');
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Box::new(e)),
            _ => {} // Ignore other events.
        }
        buf.clear();
    }
    Ok(text.trim().to_string())
}

/// Extracts text from a file based on its extension.
pub fn extract_text(
    path: &str,
    strip_hyperlinks: bool,
) -> Result<String, Box<dyn std::error::Error>> {
    let ext = Path::new(path)
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    match ext.as_str() {
        "md" => extract_text_from_md(path),
        "docx" => extract_text_from_docx(path, strip_hyperlinks),
        _ => Err(format!("Unsupported file type: .{}", ext).into()),
    }
}

/// Merges the text extracted from multiple DOCX files into one string.
/// Each file's text is separated by two newline characters.
pub fn merge_docx_files(
    paths: &[&str],
    strip_hyperlinks: bool,
) -> Result<String, Box<dyn std::error::Error>> {
    let mut merged_text = String::new();
    for path in paths {
        let text = extract_text_from_docx(path, strip_hyperlinks)?;
        merged_text.push_str(&text);
        merged_text.push_str("\n\n");
    }
    Ok(merged_text.trim().to_string())
}

/// Merges text extracted from multiple files (any supported type) into one string.
/// Prints progress for each file processed.
pub fn merge_files(
    paths: &[&str],
    strip_hyperlinks: bool,
) -> Result<String, Box<dyn std::error::Error>> {
    let mut merged_text = String::new();
    for path in paths {
        println!("  Processing: {}", path);
        let text = extract_text(path, strip_hyperlinks)?;
        merged_text.push_str(&text);
        merged_text.push_str("\n\n");
    }
    Ok(merged_text.trim().to_string())
}

/// Scans a directory for files matching the given type extensions.
/// Skips hidden directories (starting with `.`), `target/`, and `node_modules/`.
pub fn scan_directory(start_dir: &str, types: &[&str], recurse: bool) -> Vec<String> {
    let walker = if recurse {
        WalkDir::new(start_dir)
    } else {
        WalkDir::new(start_dir).max_depth(1)
    };

    let mut files: Vec<String> = walker
        .into_iter()
        .filter_entry(|e| {
            if e.file_type().is_dir() && e.depth() > 0 {
                let name = e.file_name().to_string_lossy();
                !name.starts_with('.') && name != "target" && name != "node_modules"
            } else {
                true
            }
        })
        .filter_map(|e| e.ok())
        .filter(|e| e.path().is_file())
        .filter(|e| {
            e.path()
                .extension()
                .and_then(|ext| ext.to_str())
                .map(|ext| types.iter().any(|t| t.eq_ignore_ascii_case(ext)))
                .unwrap_or(false)
        })
        .map(|e| e.path().to_string_lossy().into_owned())
        .collect();

    files.sort();
    files
}

/// Runs the application logic.
pub fn run(args: Vec<String>, output_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let program = args
        .first()
        .cloned()
        .unwrap_or_else(|| "mddocmrg".to_string());

    print_banner();

    // Parse arguments.
    let mut patterns = Vec::new();
    let mut strip_hyperlinks = false;
    let mut recurse = true;
    let mut types: Vec<String> = Vec::new();
    let mut custom_output: Option<String> = None;
    let mut i = 1;

    while i < args.len() {
        match args[i].as_str() {
            "-h" | "-?" | "--help" => {
                print_usage(&program);
                return Ok(());
            }
            "--strip-hyperlinks" | "-s" => {
                strip_hyperlinks = true;
            }
            "--no-recurse" => {
                recurse = false;
            }
            "--types" | "-t" => {
                i += 1;
                if i < args.len() {
                    types = args[i].split(',').map(|s| s.trim().to_string()).collect();
                } else {
                    return Err("--types requires an argument".into());
                }
            }
            "-o" => {
                i += 1;
                if i < args.len() {
                    custom_output = Some(args[i].clone());
                } else {
                    return Err("-o requires an argument".into());
                }
            }
            _ => {
                patterns.push(args[i].clone());
            }
        }
        i += 1;
    }

    let output = custom_output.as_deref().unwrap_or(output_path);
    let type_refs: Vec<&str> = if types.is_empty() {
        DEFAULT_TYPES.to_vec()
    } else {
        types.iter().map(|s| s.as_str()).collect()
    };

    // Collect files.
    let file_paths = if patterns.is_empty() {
        let mode = if recurse {
            "directory tree"
        } else {
            "current directory"
        };
        println!("Scanning {} from: .", mode);
        scan_directory(".", &type_refs, recurse)
    } else {
        println!("Expanding {} pattern(s)...", patterns.len());
        let mut paths = Vec::new();
        for pattern in &patterns {
            for entry in glob(pattern)? {
                match entry {
                    Ok(path) => paths.push(path.to_string_lossy().into_owned()),
                    Err(e) => eprintln!("Error processing pattern {}: {}", pattern, e),
                }
            }
        }
        paths.sort();
        paths
    };

    if file_paths.is_empty() {
        return Err("No matching files found".into());
    }

    // Count by type for summary.
    let mut counts: HashMap<String, usize> = HashMap::new();
    for p in &file_paths {
        let ext = Path::new(p)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("unknown")
            .to_lowercase();
        *counts.entry(ext).or_insert(0) += 1;
    }
    let mut type_summary: Vec<String> = counts
        .iter()
        .map(|(ext, count)| format!("{} .{}", count, ext))
        .collect();
    type_summary.sort();
    println!(
        "Found {} file(s) ({})",
        file_paths.len(),
        type_summary.join(", ")
    );

    // Merge.
    let paths_ref: Vec<&str> = file_paths.iter().map(|s| s.as_str()).collect();
    let merged_text = merge_files(&paths_ref, strip_hyperlinks)?;
    fs::write(output, &merged_text)?;

    // Final summary.
    println!(
        "Merged {} file(s) -> {} ({} bytes)",
        file_paths.len(),
        output,
        merged_text.len()
    );

    Ok(())
}

/// Main function.
#[cfg(not(test))]
fn main() {
    real_main();
}

#[allow(dead_code)]
fn real_main() {
    let args: Vec<String> = std::env::args().collect();
    real_main_inner(args);
}

fn real_main_inner(args: Vec<String>) {
    if let Err(e) = run(args, "merged.txt") {
        eprintln!("Error: {}", e);
        #[cfg(not(test))]
        std::process::exit(1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use tempfile::tempdir;
    use zip::write::FileOptions;
    use zip::CompressionMethod;

    /// Creates a temporary DOCX file with minimal content (a single paragraph).
    fn create_test_docx(text: &str) -> (tempfile::TempDir, String) {
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("test.docx");
        let file_path_str = file_path.to_str().unwrap().to_string();

        let file = File::create(&file_path).unwrap();
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // Create a minimal document.xml with a single paragraph.
        let xml_content = format!(
            r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>{}</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>"###,
            text
        );

        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(xml_content.as_bytes()).unwrap();
        zip.finish().unwrap();

        (temp_dir, file_path_str)
    }

    /// Creates a temporary DOCX file with custom XML content.
    fn create_test_docx_with_xml(xml_content: &str) -> (tempfile::TempDir, String) {
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("test.docx");
        let file_path_str = file_path.to_str().unwrap().to_string();

        let file = File::create(&file_path).unwrap();
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(xml_content.as_bytes()).unwrap();
        zip.finish().unwrap();

        (temp_dir, file_path_str)
    }

    /// Creates a temporary markdown file.
    fn create_test_md(content: &str) -> (tempfile::TempDir, String) {
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("test.md");
        let file_path_str = file_path.to_str().unwrap().to_string();
        fs::write(&file_path, content).unwrap();
        (temp_dir, file_path_str)
    }

    // --- DOCX extraction tests (preserved) ---

    #[test]
    fn test_extract_text_from_docx_without_strip() {
        let test_text = "Hello, world!";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let extracted = extract_text_from_docx(&docx_path, false).unwrap();
        assert!(extracted.contains(test_text));
    }

    #[test]
    fn test_merge_docx_files_without_strip() {
        let test_text1 = "First document text.";
        let test_text2 = "Second document text.";
        let (_temp_dir1, docx_path1) = create_test_docx(test_text1);
        let (_temp_dir2, docx_path2) = create_test_docx(test_text2);

        let merged = merge_docx_files(&[&docx_path1, &docx_path2], false).unwrap();
        assert!(merged.contains(test_text1));
        assert!(merged.contains(test_text2));
        assert!(merged.contains("\n\n"));
    }

    #[test]
    fn test_invalid_docx_file() {
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("invalid.docx");
        let file_path_str = file_path.to_str().unwrap().to_string();
        fs::write(&file_path, "Not a valid docx file").unwrap();
        let result = extract_text_from_docx(&file_path_str, false);
        assert!(result.is_err());
    }

    #[test]
    fn test_strip_hyperlink_instr_text() {
        let xml_content = r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText>HYPERLINK "https://example.com" 	 "_blank"</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Visible Link Text</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>"###;
        let (_temp_dir, docx_path) = create_test_docx_with_xml(xml_content);

        let extracted_with_strip = extract_text_from_docx(&docx_path, true).unwrap();
        let extracted_without_strip = extract_text_from_docx(&docx_path, false).unwrap();

        assert!(!extracted_with_strip.contains("HYPERLINK"));
        assert!(extracted_with_strip.contains("Visible Link Text"));
        assert!(extracted_without_strip.contains("HYPERLINK"));
        assert!(extracted_without_strip.contains("Visible Link Text"));
    }

    #[test]
    fn test_corrupt_xml_content() {
        let invalid_xml = r###"<w:document><w:body>Mismatched</w:document>"###;
        let (_temp_dir, docx_path) = create_test_docx_with_xml(invalid_xml);
        let result = extract_text_from_docx(&docx_path, false);
        assert!(result.is_err());
    }

    #[test]
    fn test_extract_nonexistent_file() {
        let result = extract_text_from_docx("nonexistent_file_at_all.docx", false);
        assert!(result.is_err());
    }

    #[test]
    fn test_extract_missing_document_xml() {
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("no_doc.docx");
        let file_path_str = file_path.to_str().unwrap().to_string();

        let file = File::create(&file_path).unwrap();
        let mut zip = zip::ZipWriter::new(file);
        zip.start_file("not_word/document.xml", Default::default())
            .unwrap();
        zip.write_all(b"content").unwrap();
        zip.finish().unwrap();

        let result = extract_text_from_docx(&file_path_str, false);
        assert!(result.is_err());
    }

    #[test]
    fn test_invalid_escape_sequence() {
        let invalid_xml = r###"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>&invalid;</w:t></w:r></w:p></w:body></w:document>"###;
        let (_temp_dir, docx_path) = create_test_docx_with_xml(invalid_xml);
        let result = extract_text_from_docx(&docx_path, false);
        assert!(result.is_err());
    }

    // --- Markdown extraction tests ---

    #[test]
    fn test_extract_text_from_md() {
        let content = "# Hello\n\nThis is a test.";
        let (_temp_dir, md_path) = create_test_md(content);
        let extracted = extract_text_from_md(&md_path).unwrap();
        assert_eq!(extracted, content);
    }

    #[test]
    fn test_extract_text_from_md_nonexistent() {
        let result = extract_text_from_md("nonexistent_file.md");
        assert!(result.is_err());
    }

    // --- Dispatch tests ---

    #[test]
    fn test_extract_text_dispatch_md() {
        let content = "# Markdown content";
        let (_temp_dir, md_path) = create_test_md(content);
        let extracted = extract_text(&md_path, false).unwrap();
        assert!(extracted.contains("Markdown content"));
    }

    #[test]
    fn test_extract_text_dispatch_docx() {
        let test_text = "DOCX dispatch test";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let extracted = extract_text(&docx_path, false).unwrap();
        assert!(extracted.contains(test_text));
    }

    #[test]
    fn test_extract_text_unsupported_type() {
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("test.xyz");
        fs::write(&file_path, "content").unwrap();
        let result = extract_text(file_path.to_str().unwrap(), false);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("Unsupported"));
    }

    // --- Directory scanning tests ---

    #[test]
    fn test_scan_directory_recurse() {
        let temp_dir = tempdir().unwrap();
        let sub_dir = temp_dir.path().join("sub");
        fs::create_dir(&sub_dir).unwrap();

        fs::write(temp_dir.path().join("root.md"), "# Root").unwrap();
        fs::write(sub_dir.join("child.md"), "# Child").unwrap();
        fs::write(temp_dir.path().join("ignore.txt"), "ignored").unwrap();

        let files = scan_directory(temp_dir.path().to_str().unwrap(), &["md"], true);
        assert_eq!(files.len(), 2);
    }

    #[test]
    fn test_scan_directory_no_recurse() {
        let temp_dir = tempdir().unwrap();
        let sub_dir = temp_dir.path().join("sub");
        fs::create_dir(&sub_dir).unwrap();

        fs::write(temp_dir.path().join("root.md"), "# Root").unwrap();
        fs::write(sub_dir.join("child.md"), "# Child").unwrap();

        let files = scan_directory(temp_dir.path().to_str().unwrap(), &["md"], false);
        assert_eq!(files.len(), 1);
    }

    #[test]
    fn test_scan_directory_skips_hidden() {
        let temp_dir = tempdir().unwrap();
        let hidden_dir = temp_dir.path().join(".hidden");
        fs::create_dir(&hidden_dir).unwrap();

        fs::write(temp_dir.path().join("visible.md"), "visible").unwrap();
        fs::write(hidden_dir.join("hidden.md"), "hidden").unwrap();

        let files = scan_directory(temp_dir.path().to_str().unwrap(), &["md"], true);
        assert_eq!(files.len(), 1);
    }

    #[test]
    fn test_scan_directory_multiple_types() {
        let temp_dir = tempdir().unwrap();
        fs::write(temp_dir.path().join("test.md"), "md content").unwrap();

        let (_docx_dir, docx_path) = create_test_docx("docx content");
        let dest = temp_dir.path().join("test.docx");
        fs::copy(&docx_path, &dest).unwrap();

        let files = scan_directory(temp_dir.path().to_str().unwrap(), &["md", "docx"], true);
        assert_eq!(files.len(), 2);
    }

    #[test]
    fn test_scan_directory_empty() {
        let temp_dir = tempdir().unwrap();
        let files = scan_directory(temp_dir.path().to_str().unwrap(), &["md"], true);
        assert!(files.is_empty());
    }

    // --- Merge files tests ---

    #[test]
    fn test_merge_files_md() {
        let (_temp_dir1, md_path1) = create_test_md("First markdown");
        let (_temp_dir2, md_path2) = create_test_md("Second markdown");

        let merged = merge_files(&[&md_path1, &md_path2], false).unwrap();
        assert!(merged.contains("First markdown"));
        assert!(merged.contains("Second markdown"));
        assert!(merged.contains("\n\n"));
    }

    // --- CLI / run() tests ---

    #[test]
    fn test_run_help() {
        let args = vec!["prog".to_string(), "-h".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_help_question_mark() {
        let args = vec!["prog".to_string(), "-?".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_no_args_scan_mode() {
        // No args triggers scan mode; project root has .md files
        let temp_dir = tempdir().unwrap();
        let output = temp_dir.path().join("scan_output.txt");
        let args = vec!["prog".to_string()];
        let result = run(args, output.to_str().unwrap());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_empty_args_list() {
        // Empty args also triggers scan mode
        let temp_dir = tempdir().unwrap();
        let output = temp_dir.path().join("scan_output.txt");
        let args: Vec<String> = vec![];
        let result = run(args, output.to_str().unwrap());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_invalid_pattern() {
        let temp_dir = tempdir().unwrap();
        let output = temp_dir.path().join("output.txt");
        let args = vec!["prog".to_string(), "nonexistent_file_*.docx".to_string()];
        let result = run(args, output.to_str().unwrap());
        assert!(result.is_err());
        assert_eq!(result.unwrap_err().to_string(), "No matching files found");
    }

    #[test]
    fn test_run_valid_workflow() {
        let test_text = "Integration Test Content";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let output_dir = tempdir().unwrap();
        let output_file = output_dir.path().join("merged_valid.txt");

        let args = vec!["prog".to_string(), docx_path];
        let result = run(args, output_file.to_str().unwrap());
        assert!(result.is_ok());

        let content = fs::read_to_string(&output_file).unwrap();
        assert!(content.contains(test_text));
    }

    #[test]
    fn test_run_with_strip_flag() {
        let test_text = "Content with flag";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let output_dir = tempdir().unwrap();
        let output_file = output_dir.path().join("merged_stripped.txt");

        let args = vec!["prog".to_string(), "-s".to_string(), docx_path];
        let result = run(args, output_file.to_str().unwrap());
        assert!(result.is_ok());

        let content = fs::read_to_string(&output_file).unwrap();
        assert!(content.contains(test_text));
    }

    #[test]
    fn test_run_output_error() {
        let test_text = "Content";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let temp_dir_out = tempdir().unwrap();
        // Passing a directory as output path causes a write error
        let output_path = temp_dir_out.path().to_str().unwrap();

        let args = vec!["prog".to_string(), docx_path];
        let result = run(args, output_path);
        assert!(result.is_err());
    }

    #[test]
    fn test_run_output_flag() {
        let test_text = "Output flag test";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let output_dir = tempdir().unwrap();
        let output_file = output_dir.path().join("custom_output.txt");

        let args = vec![
            "prog".to_string(),
            "-o".to_string(),
            output_file.to_str().unwrap().to_string(),
            docx_path,
        ];
        let result = run(args, "should_not_be_used.txt");
        assert!(result.is_ok());

        let content = fs::read_to_string(&output_file).unwrap();
        assert!(content.contains(test_text));
        assert!(!Path::new("should_not_be_used.txt").exists());
    }

    #[test]
    fn test_run_with_types_flag() {
        let temp_dir = tempdir().unwrap();
        let output = temp_dir.path().join("output.txt");
        // Scan for .xyz files — none exist
        let args = vec!["prog".to_string(), "-t".to_string(), "xyz".to_string()];
        let result = run(args, output.to_str().unwrap());
        assert!(result.is_err());
    }

    #[test]
    fn test_types_requires_arg() {
        let args = vec!["prog".to_string(), "-t".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_err());
        assert!(result
            .unwrap_err()
            .to_string()
            .contains("--types requires"));
    }

    #[test]
    fn test_output_requires_arg() {
        let args = vec!["prog".to_string(), "-o".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_err());
        assert!(result
            .unwrap_err()
            .to_string()
            .contains("-o requires"));
    }

    #[test]
    fn test_print_usage_no_filename() {
        print_usage("");
    }

    #[test]
    fn test_run_glob_error() {
        // Placeholder for GlobError testing
    }

    #[test]
    fn test_real_main_inner_error() {
        let args = vec![
            "prog".to_string(),
            "nonexistent_pattern_123.docx".to_string(),
        ];
        real_main_inner(args);
    }

    #[test]
    fn test_run_glob_pattern_error() {
        let args = vec!["prog".to_string(), "[".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_err());
    }
}
