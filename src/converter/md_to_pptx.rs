use crate::parser::markdown::MarkdownDocument;
use crate::presentation::{builder::PresentationBuilder, templates::SlideTemplate};
use crate::utils::{error::Error, file_io};
use crate::LogLevel;
use crate::Result;
use std::path::{Path, PathBuf};

/// Convert Markdown files from a directory to a PowerPoint presentation
pub fn convert_markdown_to_pptx(
    input_dir: &Path,
    output_file: &Path,
    template_name: &str,
    recursive: bool,
    log_level: LogLevel,
) -> Result<()> {
    if log_level.should_print_info() {
        println!("Starting conversion process...");
    }

    // Find all Markdown files in the input directory
    let markdown_files = if recursive {
        file_io::find_markdown_files(input_dir)?
    } else {
        find_markdown_files_non_recursive(input_dir)?
    };

    if log_level.should_print_info() {
        println!("Found {} Markdown files to process", markdown_files.len());
    }

    // Parse all Markdown files and combine them into a single document
    let combined_document = parse_and_combine_markdown_files(&markdown_files, log_level)?;

    if log_level.should_print_info() {
        println!(
            "Parsed Markdown files into {} slides",
            combined_document.slides.len()
        );
    }

    // Get the template
    let template = SlideTemplate::from_name(template_name);

    if log_level.should_print_debug() {
        println!("Using template: {:?}", template);
    }

    // Build the PowerPoint presentation
    let presentation_builder = PresentationBuilder::from_markdown(&combined_document, template)?;

    if log_level.should_print_info() {
        println!("Building PowerPoint presentation...");
    }

    let pptx_data = presentation_builder.build()?;

    // Write the output file
    file_io::write_file(output_file, &pptx_data)?;

    if log_level.should_print_info() {
        println!(
            "Successfully created PowerPoint file: {}",
            output_file.display()
        );
        println!("File size: {} bytes", pptx_data.len());
    }

    Ok(())
}

/// Find Markdown files in a directory (non-recursive)
fn find_markdown_files_non_recursive(dir: &Path) -> Result<Vec<PathBuf>> {
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

    let entries = std::fs::read_dir(dir)?;

    for entry in entries {
        let entry = entry?;
        let path = entry.path();

        if path.is_file() {
            if let Some(extension) = path.extension() {
                if extension == "md" || extension == "markdown" {
                    markdown_files.push(path);
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

/// Parse multiple Markdown files and combine them into a single document
fn parse_and_combine_markdown_files(
    markdown_files: &[PathBuf],
    log_level: LogLevel,
) -> Result<MarkdownDocument> {
    let mut combined_slides = Vec::new();
    let mut combined_metadata = crate::parser::markdown::DocumentMetadata::default();

    // Track if we've set the main metadata yet
    let mut metadata_set = false;

    for (index, file_path) in markdown_files.iter().enumerate() {
        if log_level.should_print_debug() {
            println!(
                "Processing file {}/{}: {}",
                index + 1,
                markdown_files.len(),
                file_path.display()
            );
        }

        // Read the file content
        let content = file_io::read_file_to_string(file_path)?;

        // Parse the Markdown document
        let document = MarkdownDocument::parse(&content).map_err(|e| {
            Error::conversion(format!("Failed to parse {}: {}", file_path.display(), e))
        })?;

        // Use metadata from the first file that has it
        if !metadata_set && (document.metadata.title.is_some() || document.metadata.author.is_some()) {
            combined_metadata = document.metadata.clone();
            metadata_set = true;
        }

        // Add all slides from this document
        let slide_count = document.slides.len();
        for mut slide in document.slides {
            // If the slide doesn't have a title, create one from the filename
            if slide.title.is_none() && !slide.content.is_empty() {
                if let Some(filename) = file_path.file_stem() {
                    slide.title = Some(filename.to_string_lossy().to_string());
                }
            }

            combined_slides.push(slide);
        }

        if log_level.should_print_debug() {
            println!(
                "  Added {} slides from {}",
                slide_count,
                file_path.display()
            );
        }
    }

    // If no metadata was found, create default metadata from the directory name
    if !metadata_set {
        if let Some(dir_name) = markdown_files[0]
            .parent()
            .and_then(|p| p.file_name())
            .and_then(|n| n.to_str())
        {
            combined_metadata.title = Some(format!("{} Presentation", dir_name));
        } else {
            combined_metadata.title = Some("Markdown Presentation".to_string());
        }
    }

    if combined_slides.is_empty() {
        return Err(Error::conversion("No slides found in any Markdown files"));
    }

    Ok(MarkdownDocument {
        slides: combined_slides,
        metadata: combined_metadata,
    })
}

/// Convert each Markdown file in a directory to separate PowerPoint presentations
pub fn convert_separate_files(
    input_dir: &Path,
    output_dir: &Path,
    template_name: &str,
    recursive: bool,
    log_level: LogLevel,
) -> Result<usize> {
    if log_level.should_print_info() {
        println!("Starting separate file conversion process...");
    }

    // Find all Markdown files in the input directory
    let markdown_files = if recursive {
        file_io::find_markdown_files(input_dir)?
    } else {
        find_markdown_files_non_recursive(input_dir)?
    };

    if log_level.should_print_info() {
        println!(
            "Found {} Markdown files to process separately",
            markdown_files.len()
        );
    }

    let mut processed_count = 0;

    // Process each markdown file separately
    for file_path in &markdown_files {
        if log_level.should_print_debug() {
            println!("Processing: {}", file_path.display());
        }

        // Create output filename based on input filename
        let output_filename = if let Some(stem) = file_path.file_stem() {
            format!("{}.pptx", stem.to_string_lossy())
        } else {
            format!("presentation_{}.pptx", processed_count + 1)
        };

        let output_file = output_dir.join(output_filename);

        // Convert the single file
        match convert_single_markdown_file(file_path, &output_file, template_name, log_level) {
            Ok(_) => {
                processed_count += 1;
                if log_level.should_print_info() {
                    println!("  ✓ Created: {}", output_file.display());
                }
            }
            Err(e) => {
                eprintln!("  ✗ Failed to convert {}: {}", file_path.display(), e);
                // Continue processing other files instead of failing completely
            }
        }
    }

    if processed_count == 0 {
        return Err(Error::conversion("No files were successfully processed"));
    }

    if log_level.should_print_info() {
        println!(
            "Successfully processed {} out of {} files",
            processed_count,
            markdown_files.len()
        );
    }

    Ok(processed_count)
}

/// Convert a single Markdown file to PowerPoint
pub fn convert_single_markdown_file(
    input_file: &Path,
    output_file: &Path,
    template_name: &str,
    log_level: LogLevel,
) -> Result<()> {
    if log_level.should_print_info() {
        println!("Converting single file: {}", input_file.display());
    }

    // Validate input file
    if !input_file.exists() {
        return Err(Error::file_not_found(input_file.display().to_string()));
    }

    file_io::validate_file_extension(input_file, "md")
        .or_else(|_| file_io::validate_file_extension(input_file, "markdown"))?;

    // Read and parse the Markdown file
    let content = file_io::read_file_to_string(input_file)?;
    let document = MarkdownDocument::parse(&content)?;

    if log_level.should_print_info() {
        println!("Parsed {} slides from Markdown file", document.slides.len());
    }

    // Get the template and build presentation
    let template = SlideTemplate::from_name(template_name);
    let presentation_builder = PresentationBuilder::from_markdown(&document, template)?;
    let pptx_data = presentation_builder.build()?;

    // Write the output file
    file_io::write_file(output_file, &pptx_data)?;

    if log_level.should_print_info() {
        println!(
            "Successfully created PowerPoint file: {}",
            output_file.display()
        );
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use tempfile::tempdir;

    #[test]
    fn test_find_markdown_files_non_recursive() {
        let temp_dir = tempdir().unwrap();
        let temp_path = temp_dir.path();

        // Create test files
        fs::write(temp_path.join("test1.md"), "# Test 1").unwrap();
        fs::write(temp_path.join("test2.markdown"), "# Test 2").unwrap();
        fs::write(temp_path.join("test.txt"), "Not markdown").unwrap();

        // Create subdirectory with markdown file (should be ignored in non-recursive mode)
        let sub_dir = temp_path.join("subdir");
        fs::create_dir(&sub_dir).unwrap();
        fs::write(sub_dir.join("test3.md"), "# Test 3").unwrap();

        let files = find_markdown_files_non_recursive(temp_path).unwrap();
        assert_eq!(files.len(), 2);

        let filenames: Vec<_> = files
            .iter()
            .map(|p| p.file_name().unwrap().to_str().unwrap())
            .collect();
        assert!(filenames.contains(&"test1.md"));
        assert!(filenames.contains(&"test2.markdown"));
        assert!(!filenames.contains(&"test3.md")); // Should not be found in non-recursive mode
    }

    #[test]
    fn test_parse_and_combine_markdown_files() {
        let temp_dir = tempdir().unwrap();
        let temp_path = temp_dir.path();

        // Create test files with different content
        fs::write(
            temp_path.join("file1.md"),
            r#"# File 1 Title

Content from file 1.

## File 1 Section

More content."#,
        )
        .unwrap();

        fs::write(
            temp_path.join("file2.md"),
            r#"# File 2 Title

Content from file 2.

- Item 1
- Item 2"#,
        )
        .unwrap();

        let files = vec![temp_path.join("file1.md"), temp_path.join("file2.md")];

        let combined = parse_and_combine_markdown_files(&files, LogLevel::Quiet).unwrap();

        // Should have slides from both files
        assert!(combined.slides.len() >= 2);

        // Check that content from both files is present
        let all_titles: Vec<_> = combined
            .slides
            .iter()
            .filter_map(|s| s.title.as_ref())
            .collect();

        assert!(all_titles.iter().any(|&title| title.contains("File 1")));
        assert!(all_titles.iter().any(|&title| title.contains("File 2")));
    }

    #[test]
    fn test_convert_single_markdown_file() {
        let temp_dir = tempdir().unwrap();
        let temp_path = temp_dir.path();

        let input_file = temp_path.join("test.md");
        let output_file = temp_path.join("test.pptx");

        fs::write(
            &input_file,
            r#"# Test Presentation

This is a test slide.

## Second Slide

- Point 1
- Point 2
- Point 3"#,
        )
        .unwrap();

        let result =
            convert_single_markdown_file(&input_file, &output_file, "default", LogLevel::Quiet);

        assert!(result.is_ok());
        assert!(output_file.exists());

        // Check that the file has some content (PPTX files should be reasonably sized)
        let file_size = fs::metadata(&output_file).unwrap().len();
        assert!(file_size > 1000); // PPTX files should be at least 1KB
    }
}
