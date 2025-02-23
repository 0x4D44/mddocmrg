use std::fs::File;
use std::io::Read;
use std::path::Path;
use zip::read::ZipArchive;
use quick_xml::Reader;
use quick_xml::events::Event;
use glob::glob;

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
                match e.name().as_ref() {
                    b"w:instrText" => in_instr_text = true,
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"w:instrText" => in_instr_text = false,
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                // If stripping hyperlinks and we're in an instruction text element,
                // skip appending this text.
                if strip_hyperlinks && in_instr_text {
                    // Skip this text.
                } else {
                    text.push_str(&e.unescape()?.into_owned());
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
    println!("Usage: {} [options] <file_pattern1> <file_pattern2> ...", prog_name);
    println!("Merges plain text extracted from DOCX files matching the given patterns.");
    println!("Options:");
    println!("  -h, -?                 Display this help message and exit.");
    println!("  --strip-hyperlinks, -s  Remove hyperlink field instructions from the output.");
}

/// Main function.
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().collect();
    let program = args.get(0).cloned().unwrap_or_else(|| "docx_merger".to_string());

    println!("{} - Merges plain text extracted from DOCX files into a single output.", program);

    if args.len() < 2 {
        print_usage(&program);
        std::process::exit(1);
    }

    // Process command-line arguments.
    let mut patterns = Vec::new();
    let mut strip_hyperlinks = false;
    for arg in args.iter().skip(1) {
        match arg.as_str() {
            "-h" | "-?" => {
                print_usage(&program);
                std::process::exit(0);
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
        std::process::exit(1);
    }

    let paths_ref: Vec<&str> = file_paths.iter().map(|s| s.as_str()).collect();
    let merged_text = merge_docx_files(&paths_ref, strip_hyperlinks)?;
    std::fs::write("merged.txt", merged_text)?;
    println!("Merged text written to merged.txt");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;
    use std::fs::File;
    use std::io::Write;
    use zip::write::FileOptions;
    use zip::CompressionMethod;

    /// Creates a temporary DOCX file with minimal content (a single paragraph).
    fn create_test_docx(text: &str) -> Result<(tempfile::TempDir, String), Box<dyn std::error::Error>> {
        let temp_dir = tempdir()?;
        let file_path = temp_dir.path().join("test.docx");
        let file_path_str = file_path.to_str().unwrap().to_string();

        let file = File::create(&file_path)?;
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // Create a minimal document.xml with a single paragraph.
        let xml_content = format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>{}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#, text);

        zip.start_file("word/document.xml", options)?;
        zip.write_all(xml_content.as_bytes())?;
        zip.finish()?;

        Ok((temp_dir, file_path_str))
    }

    /// Creates a temporary DOCX file with custom XML content.
    fn create_test_docx_with_xml(xml_content: &str) -> Result<(tempfile::TempDir, String), Box<dyn std::error::Error>> {
        let temp_dir = tempdir()?;
        let file_path = temp_dir.path().join("test.docx");
        let file_path_str = file_path.to_str().unwrap().to_string();

        let file = File::create(&file_path)?;
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        zip.start_file("word/document.xml", options)?;
        zip.write_all(xml_content.as_bytes())?;
        zip.finish()?;

        Ok((temp_dir, file_path_str))
    }

    #[test]
    fn test_extract_text_from_docx_without_strip() {
        let test_text = "Hello, world!";
        let (_temp_dir, docx_path) = create_test_docx(test_text).unwrap();
        let extracted = extract_text_from_docx(&docx_path, false).unwrap();
        assert!(extracted.contains(test_text));
    }

    #[test]
    fn test_merge_docx_files_without_strip() {
        let test_text1 = "First document text.";
        let test_text2 = "Second document text.";
        let (_temp_dir1, docx_path1) = create_test_docx(test_text1).unwrap();
        let (_temp_dir2, docx_path2) = create_test_docx(test_text2).unwrap();

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
        // Create a DOCX file that contains a hyperlink field.
        // Typically, a hyperlink field is stored as an instruction text (w:instrText)
        // followed by the visible text.
        let xml_content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:instrText>HYPERLINK "https://example.com" \t "_blank"</w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <w:t>Visible Link Text</w:t>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;
        let (_temp_dir, docx_path) = create_test_docx_with_xml(xml_content).unwrap();

        // When strip_hyperlinks is true, the hyperlink field instruction should be omitted.
        let extracted_with_strip = extract_text_from_docx(&docx_path, true).unwrap();
        // When strip_hyperlinks is false, both the instruction and the visible text will appear.
        let extracted_without_strip = extract_text_from_docx(&docx_path, false).unwrap();

        assert!(!extracted_with_strip.contains("HYPERLINK"), "Instruction text should be stripped");
        assert!(extracted_with_strip.contains("Visible Link Text"), "Visible text should be kept");
        assert!(extracted_without_strip.contains("HYPERLINK"), "Instruction text is present when not stripping");
        assert!(extracted_without_strip.contains("Visible Link Text"), "Visible text is present");
    }
}
