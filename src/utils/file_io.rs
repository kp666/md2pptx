use crate::utils::error::Error;
use crate::Result;
use std::fs;
use std::path::{Path, PathBuf};
use walkdir::WalkDir;

/// Recursively find all Markdown files in a directory
pub fn find_markdown_files(dir: &Path) -> Result<Vec<PathBuf>> {
    let mut markdown_files = Vec::new();

    if !dir.exists() {
        return Err(Error::file_not_found(dir.display().to_string()));
    }

    if !dir.is_dir() {
        return Err(Error::configuration(format!(
            "Path {} is not a directory",
            dir.display()
        )));
    }

    for entry in WalkDir::new(dir) {
        let entry = entry?;
        let path = entry.path();

        if path.is_file() {
            if let Some(extension) = path.extension() {
                if extension == "md" || extension == "markdown" {
                    markdown_files.push(path.to_path_buf());
                }
            }
        }
    }

    if markdown_files.is_empty() {
        return Err(Error::configuration(format!(
            "No Markdown files found in directory: {}",
            dir.display()
        )));
    }

    // Sort files for consistent processing order
    markdown_files.sort();
    Ok(markdown_files)
}

/// Read file contents as UTF-8 string
pub fn read_file_to_string(path: &Path) -> Result<String> {
    if !path.exists() {
        return Err(Error::file_not_found(path.display().to_string()));
    }

    let content = fs::read_to_string(path)?;
    Ok(content)
}

/// Write bytes to file
pub fn write_file(path: &Path, content: &[u8]) -> Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }

    fs::write(path, content)?;
    Ok(())
}

/// Validate that a path has the expected extension
pub fn validate_file_extension(path: &Path, expected_ext: &str) -> Result<()> {
    match path.extension() {
        Some(ext) if ext == expected_ext => Ok(()),
        Some(ext) => Err(Error::invalid_file_format(format!(
            "Expected .{} file, got .{}",
            expected_ext,
            ext.to_string_lossy()
        ))),
        None => Err(Error::invalid_file_format(format!(
            "File {} has no extension, expected .{}",
            path.display(),
            expected_ext
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use tempfile::tempdir;

    #[test]
    fn test_find_markdown_files() {
        let temp_dir = tempdir().unwrap();
        let temp_path = temp_dir.path();

        // Create test files
        fs::write(temp_path.join("test1.md"), "# Test 1").unwrap();
        fs::write(temp_path.join("test2.markdown"), "# Test 2").unwrap();
        fs::write(temp_path.join("test.txt"), "Not markdown").unwrap();

        let files = find_markdown_files(temp_path).unwrap();
        assert_eq!(files.len(), 2);

        let filenames: Vec<_> = files
            .iter()
            .map(|p| p.file_name().unwrap().to_str().unwrap())
            .collect();
        assert!(filenames.contains(&"test1.md"));
        assert!(filenames.contains(&"test2.markdown"));
    }

    #[test]
    fn test_validate_file_extension() {
        let path = Path::new("test.pptx");
        assert!(validate_file_extension(path, "pptx").is_ok());
        assert!(validate_file_extension(path, "pdf").is_err());
    }
}
