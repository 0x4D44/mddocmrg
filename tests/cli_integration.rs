use std::process::Command;

#[test]
fn test_main_execution() {
    // Build the binary if needed, but cargo test usually ensures it's available.
    // However, llvm-cov might need special handling.
    // Actually, llvm-cov runs the tests. 
    // We can try running the binary from target/debug.
    
    let output = Command::new("cargo")
        .args(&["run", "--", "-h"])
        .output()
        .expect("Failed to execute cargo run");
    
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Usage:"));
}

#[test]
fn test_main_no_args() {
    let output = Command::new("cargo")
        .args(&["run", "--"])
        .output()
        .expect("Failed to execute cargo run");
    
    assert!(!output.status.success());
}
