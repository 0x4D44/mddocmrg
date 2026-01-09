use glob::glob;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::fs::File;
use std::io::Read;
use std::path::Path;
use zip::read::ZipArchive;

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

/// Prints usage instructions.
fn print_usage(program: &str) {
    let prog_name = Path::new(program)
        .file_name()
        .map(|s| s.to_string_lossy())
        .unwrap_or_else(|| "docx_merger".into());
    println!(
        "Usage: {} [options] <file_pattern1> <file_pattern2> ...",
        prog_name
    );
    println!("Merges plain text extracted from DOCX files matching the given patterns.");
    println!("Options:");
    println!("  -h, -?                 Display this help message and exit.");
    println!("  --strip-hyperlinks, -s  Remove hyperlink field instructions from the output.");
}

/// Runs the application logic.
/// Returns Ok(()) on success (including help), or Err if an error occurs.
/// `output_path` allows specifying the destination file (default is "merged.txt" in CLI).
pub fn run(args: Vec<String>, output_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let program = args
        .first()
        .cloned()
        .unwrap_or_else(|| "docx_merger".to_string());

    if args.len() < 2 {
        println!(
            "{} - Merges plain text extracted from DOCX files into a single output.",
            program
        );
        print_usage(&program);
        return Err("Insufficient arguments".into());
    }

    // Process command-line arguments.
    let mut patterns = Vec::new();
    let mut strip_hyperlinks = false;
    for arg in args.iter().skip(1) {
        match arg.as_str() {
            "-h" | "-?" => {
                println!(
                    "{} - Merges plain text extracted from DOCX files into a single output.",
                    program
                );
                print_usage(&program);
                return Ok(());
            }
            "--strip-hyperlinks" | "-s" => {
                strip_hyperlinks = true;
            }
            _ => {
                patterns.push(arg);
            }
        }
    }

    // Expand wildcards using the glob crate.
    let mut file_paths = Vec::new();
    for pattern in patterns {
        for entry in glob(pattern)? {
            match entry {
                Ok(path) => file_paths.push(path.to_string_lossy().into_owned()),
                Err(e) => eprintln!("Error processing pattern {}: {}", pattern, e),
            }
        }
    }

    if file_paths.is_empty() {
        eprintln!("No files found matching the specified patterns.");
        return Err("No matching files found".into());
    }

    let paths_ref: Vec<&str> = file_paths.iter().map(|s| s.as_str()).collect();
    let merged_text = merge_docx_files(&paths_ref, strip_hyperlinks)?;
    std::fs::write(output_path, merged_text)?;
    println!("Merged text written to {}", output_path);
    Ok(())
}

/// Main function.
#[cfg(not(test))]
fn main() {
    real_main();
}

fn real_main() {
    let args: Vec<String> = std::env::args().collect();
    real_main_inner(args);
}

fn real_main_inner(args: Vec<String>) {
    // In production, we always write to "merged.txt"
    if let Err(e) = run(args, "merged.txt") {
        if e.to_string() != "Insufficient arguments" {
            eprintln!("Error: {}", e);
        }
        #[cfg(not(test))]
        std::process::exit(1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs::File;
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
        std::fs::write(&file_path, "Not a valid docx file").unwrap();
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
    fn test_run_help() {
        let args = vec!["prog".to_string(), "-h".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_no_args() {
        let args = vec!["prog".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_err());
        assert_eq!(result.unwrap_err().to_string(), "Insufficient arguments");
    }

    #[test]
    fn test_run_invalid_pattern() {
        let args = vec!["prog".to_string(), "nonexistent_file_*.docx".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_err());
        assert_eq!(result.unwrap_err().to_string(), "No matching files found");
    }

    #[test]
    fn test_run_valid_workflow() {
        let test_text = "Integration Test Content";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let output_file = "merged_valid.txt";

        let args = vec!["prog".to_string(), docx_path.clone()];
        let result = run(args, output_file);
        assert!(result.is_ok());

        let content = std::fs::read_to_string(output_file).unwrap();
        assert!(content.contains(test_text));
        let _ = std::fs::remove_file(output_file);
    }

    #[test]
    fn test_corrupt_xml_content() {
        let invalid_xml = r###"<w:document><w:body>Mismatched</w:document>"###;
        let (_temp_dir, docx_path) = create_test_docx_with_xml(invalid_xml);
        let result = extract_text_from_docx(&docx_path, false);
        assert!(result.is_err());
    }

    #[test]
    fn test_run_empty_args_list() {
        let args = vec![];
        let result = run(args, "merged.txt");
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
    fn test_print_usage_no_filename() {
        print_usage("");
    }

    #[test]
    fn test_run_with_strip_flag() {
        let test_text = "Content with flag";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let output_file = "merged_stripped.txt";

        let args = vec!["prog".to_string(), "-s".to_string(), docx_path.clone()];
        let result = run(args, output_file);
        assert!(result.is_ok());

        let content = std::fs::read_to_string(output_file).unwrap();
        assert!(content.contains(test_text));
        let _ = std::fs::remove_file(output_file);
    }

    #[test]
    fn test_run_output_error() {
        let test_text = "Content";
        let (_temp_dir, docx_path) = create_test_docx(test_text);
        let temp_dir_out = tempdir().unwrap();
        let output_path = temp_dir_out.path().to_str().unwrap();

        let args = vec!["prog".to_string(), docx_path];
        let result = run(args, output_path);
        assert!(result.is_err());
    }

    #[test]
    fn test_main_entry() {
        real_main();
    }

    #[test]
    fn test_invalid_escape_sequence() {
        let invalid_xml = r###"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>&invalid;</w:t></w:r></w:p></w:body></w:document>"###;
        let (_temp_dir, docx_path) = create_test_docx_with_xml(invalid_xml);
        let result = extract_text_from_docx(&docx_path, false);
        assert!(result.is_err());
    }

    #[test]
    fn test_run_glob_error() {
        // Placeholder for GlobError testing if possible
    }

    #[test]
    fn test_real_main_inner_error() {
        let args = vec!["prog".to_string(), "nonexistent_pattern_123.docx".to_string()];
        real_main_inner(args);
    }

    #[test]
    fn test_run_glob_pattern_error() {
        let args = vec!["prog".to_string(), "[".to_string()];
        let result = run(args, "merged.txt");
        assert!(result.is_err());
    }
}