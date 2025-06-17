use clap::{Arg, Command};
use md2pptx::{convert_markdown_to_pptx, LogLevel, Result};
use std::path::PathBuf;
fn main() -> Result<()> {
    let matches = Command::new("md2pptx")
        .version("0.1.0")
        .author("Your Name <your.email@example.com>")
        .about("Convert Markdown files to PowerPoint presentations")
        .arg(
            Arg::new("input")
                .help("Input directory containing Markdown files")
                .required(true)
                .index(1)
                .value_parser(clap::value_parser!(PathBuf)),
        )
        .arg(
            Arg::new("output")
                .help("Output PowerPoint file (.pptx) or output directory for separate files")
                .required(true)
                .index(2)
                .value_parser(clap::value_parser!(PathBuf)),
        )
        .arg(
            Arg::new("template")
                .short('t')
                .long("template")
                .help("PowerPoint template to use")
                .value_name("TEMPLATE")
                .default_value("default"),
        )
        .arg(
            Arg::new("recursive")
                .short('r')
                .long("recursive")
                .help("Process subdirectories recursively")
                .action(clap::ArgAction::SetTrue),
        )
        .arg(
            Arg::new("verbose")
                .short('v')
                .long("verbose")
                .help("Enable verbose output")
                .action(clap::ArgAction::SetTrue),
        )
        .arg(
            Arg::new("quiet")
                .short('q')
                .long("quiet")
                .help("Suppress all output except errors")
                .action(clap::ArgAction::SetTrue)
                .conflicts_with("verbose"),
        )
        .arg(
            Arg::new("separate")
                .short('s')
                .long("separate")
                .help("Generate separate PPTX file for each MD file (output must be directory)")
                .action(clap::ArgAction::SetTrue),
        )
        .get_matches();

    let input_dir = matches.get_one::<PathBuf>("input").unwrap();
    let output_path = matches.get_one::<PathBuf>("output").unwrap();
    let template = matches.get_one::<String>("template").unwrap();
    let recursive = matches.get_flag("recursive");
    let verbose = matches.get_flag("verbose");
    let quiet = matches.get_flag("quiet");
    let separate = matches.get_flag("separate");

    // Set up logging level based on flags
    let log_level = if quiet {
        LogLevel::Quiet
    } else if verbose {
        LogLevel::Verbose
    } else {
        LogLevel::Normal
    };

    if !quiet {
        if separate {
            println!(
                "Converting Markdown files from {} to separate PPTX files in {}",
                input_dir.display(),
                output_path.display()
            );
        } else {
            println!(
                "Converting Markdown files from {} to {}",
                input_dir.display(),
                output_path.display()
            );
        }
    }

    // Validate input directory exists
    if !input_dir.exists() {
        eprintln!(
            "Error: Input directory '{}' does not exist",
            input_dir.display()
        );
        std::process::exit(1);
    }

    if !input_dir.is_dir() {
        eprintln!(
            "Error: Input path '{}' is not a directory",
            input_dir.display()
        );
        std::process::exit(1);
    }

    // Validate output path based on mode
    if separate {
        // In separate mode, output must be a directory
        if output_path.exists() && !output_path.is_dir() {
            eprintln!("Error: When using --separate, output path must be a directory");
            std::process::exit(1);
        }
        // Create output directory if it doesn't exist
        if !output_path.exists() {
            if let Err(e) = std::fs::create_dir_all(output_path) {
                eprintln!(
                    "Error: Failed to create output directory '{}': {}",
                    output_path.display(),
                    e
                );
                std::process::exit(1);
            }
        }
    } else {
        // In combined mode, output must be a .pptx file
        if let Some(ext) = output_path.extension() {
            if ext != "pptx" {
                eprintln!("Error: Output file must have .pptx extension");
                std::process::exit(1);
            }
        } else {
            eprintln!("Error: Output file must have .pptx extension");
            std::process::exit(1);
        }
    }

    // Perform the conversion
    if separate {
        match md2pptx::convert_separate_files(
            input_dir,
            output_path,
            template,
            recursive,
            log_level,
        ) {
            Ok(count) => {
                if !quiet {
                    println!(
                        "Conversion completed successfully! {} files processed.",
                        count
                    );
                }
            }
            Err(e) => {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
    } else {
        match convert_markdown_to_pptx(input_dir, output_path, template, recursive, log_level) {
            Ok(_) => {
                if !quiet {
                    println!("Conversion completed successfully!");
                }
            }
            Err(e) => {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
    }

    Ok(())
}
