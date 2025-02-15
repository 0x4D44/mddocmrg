use std::fs::File;
use std::io::Read;
use std::path::Path;
use zip::read::ZipArchive;
use quick_xml::Reader;
use quick_xml::events::Event;
use glob::glob;

/// Extracts the text content from the provided DOCX file.
/// This function ignores formatting and returns the extracted text.
pub fn extract_text_from_docx(path: &str) -> Result<String, Box<dyn std::error::Error>> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file)?;
    let mut document_xml = archive.by_name("word/document.xml")?;
    let mut xml_content = String::new();
    document_xml.read_to_string(&mut xml_content)?;

    let mut reader = Reader::from_str(&xml_content);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(e)) => {
                text.push_str(&e.unescape()?.into_owned());
                text.push(' ');
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Box::new(e)),
            _ => {} // Skip any non-text events.
        }
        buf.clear();
    }
    Ok(text.trim().to_string())
}

/// Merges the text extracted from multiple DOCX files into one string.
/// Each file's text is separated by two newline characters.
pub fn merge_docx_files(paths: &[&str]) -> Result<String, Box<dyn std::error::Error>> {
    let mut merged_text = String::new();
    for path in paths {
        let text = extract_text_from_docx(path)?;
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
    println!("  -h, -?       Display this help message and exit.");
}

/// Main function.
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().collect();
    let program = args.get(0).map(|s| s.clone()).unwrap_or_else(|| "docx_merger".to_string());

    // Output programme name and brief purpose.
    println!("{} - Merges plain text extracted from DOCX files into a single output.", program);

    if args.len() < 2 {
        print_usage(&program);
        std::process::exit(1);
    }

    // Process command-line arguments.
    let mut patterns = Vec::new();
    for arg in args.iter().skip(1) {
        if arg == "-h" || arg == "-?" {
            print_usage(&program);
            std::process::exit(0);
        } else {
            patterns.push(arg);
        }
    }

    // Expand wildcards using the glob crate.
    let mut file_paths = Vec::new();
    for pattern in patterns {
        for entry in glob(pattern)? {
            match entry {
                Ok(path) => {
                    file_paths.push(path.to_string_lossy().into_owned());
                }
                Err(e) => eprintln!("Error processing pattern {}: {}", pattern, e),
            }
        }
    }

    if file_paths.is_empty() {
        eprintln!("No files found matching the specified patterns.");
        std::process::exit(1);
    }

    let paths_ref: Vec<&str> = file_paths.iter().map(|s| s.as_str()).collect();
    let merged_text = merge_docx_files(&paths_ref)?;
    std::fs::write("merged.txt", merged_text)?;
    println!("Merged text written to merged.txt");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;
    use std::fs::File;
    use std::io::Write; // Required for the write_all method.
    use zip::write::FileOptions;
    use zip::CompressionMethod;

    /// Creates a temporary DOCX file with minimal content (a single paragraph).
    /// Returns a tuple of the temporary directory (to keep it alive) and the file path.
    fn create_test_docx(text: &str) -> Result<(tempfile::TempDir, String), Box<dyn std::error::Error>> {
        let temp_dir = tempdir()?;
        let file_path = temp_dir.path().join("test.docx");
        let file_path_str = file_path.to_str().unwrap().to_string();

        let file = File::create(&file_path)?;
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // Create a minimal document.xml with a single paragraph and run.
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

    #[test]
    fn test_extract_text_from_docx() {
        let test_text = "Hello, world!";
        let (_temp_dir, docx_path) = create_test_docx(test_text).unwrap();
        let extracted = extract_text_from_docx(&docx_path).unwrap();
        assert!(extracted.contains(test_text));
    }

    #[test]
    fn test_merge_docx_files() {
        let test_text1 = "First document text.";
        let test_text2 = "Second document text.";
        let (_temp_dir1, docx_path1) = create_test_docx(test_text1).unwrap();
        let (_temp_dir2, docx_path2) = create_test_docx(test_text2).unwrap();

        let merged = merge_docx_files(&[&docx_path1, &docx_path2]).unwrap();
        assert!(merged.contains(test_text1));
        assert!(merged.contains(test_text2));
        // Check that the texts are separated by two newlines.
        assert!(merged.contains("\n\n"));
    }

    #[test]
    fn test_invalid_docx_file() {
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("invalid.docx");
        let file_path_str = file_path.to_str().unwrap().to_string();
        // Write some invalid content.
        std::fs::write(&file_path, "Not a valid docx file").unwrap();
        let result = extract_text_from_docx(&file_path_str);
        assert!(result.is_err());
    }
}
