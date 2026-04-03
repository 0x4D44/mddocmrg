use std::process::Command;

#[test]
fn test_main_help() {
    let output = Command::new("cargo")
        .args(&["run", "--", "-h"])
        .output()
        .expect("Failed to execute cargo run");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Usage:"));
}

#[test]
fn test_main_no_args_scan_mode() {
    let output = Command::new("cargo")
        .args(&["run", "--"])
        .output()
        .expect("Failed to execute cargo run");

    // No args triggers scan mode, finds .md files in project root
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Scanning directory tree"));
    assert!(stdout.contains("Merged"));

    // Clean up
    let _ = std::fs::remove_file("merged.txt");
}

#[test]
fn test_main_invalid_pattern() {
    let output = Command::new("cargo")
        .args(&["run", "--", "nonexistent_xyz_*.docx"])
        .output()
        .expect("Failed to execute cargo run");

    assert!(!output.status.success());
}
