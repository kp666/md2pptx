pub mod converter;
pub mod parser;
pub mod presentation;
pub mod utils;

pub use converter::md_to_pptx::{
    convert_markdown_to_pptx, convert_separate_files, convert_single_markdown_file,
};
pub use parser::markdown::MarkdownDocument;
pub use presentation::builder::PresentationBuilder;

#[derive(Debug, Clone, Copy)]
pub enum LogLevel {
    Quiet,
    Normal,
    Verbose,
}

impl LogLevel {
    pub fn should_print_info(&self) -> bool {
        matches!(self, LogLevel::Normal | LogLevel::Verbose)
    }

    pub fn should_print_debug(&self) -> bool {
        matches!(self, LogLevel::Verbose)
    }
}

/// Result type used throughout the application
pub type Result<T> = std::result::Result<T, crate::utils::error::Error>;
