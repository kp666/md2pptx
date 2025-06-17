use thiserror::Error;

#[derive(Error, Debug)]
pub enum Error {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Markdown parsing error: {message}")]
    MarkdownParsing { message: String },

    #[error("PowerPoint generation error: {message}")]
    PptxGeneration { message: String },

    #[error("File not found: {path}")]
    FileNotFound { path: String },

    #[error("Invalid file format: {path}")]
    InvalidFileFormat { path: String },

    #[error("Configuration error: {message}")]
    Configuration { message: String },

    #[error("Conversion error: {message}")]
    Conversion { message: String },

    #[error("XML processing error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Zip file error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("JSON serialization error: {0}")]
    Json(#[from] serde_json::Error),

    #[error("Walkdir error: {0}")]
    WalkDir(#[from] walkdir::Error),
}

impl Error {
    pub fn markdown_parsing(message: impl Into<String>) -> Self {
        Self::MarkdownParsing {
            message: message.into(),
        }
    }

    pub fn pptx_generation(message: impl Into<String>) -> Self {
        Self::PptxGeneration {
            message: message.into(),
        }
    }

    pub fn file_not_found(path: impl Into<String>) -> Self {
        Self::FileNotFound { path: path.into() }
    }

    pub fn invalid_file_format(path: impl Into<String>) -> Self {
        Self::InvalidFileFormat { path: path.into() }
    }

    pub fn configuration(message: impl Into<String>) -> Self {
        Self::Configuration {
            message: message.into(),
        }
    }

    pub fn conversion(message: impl Into<String>) -> Self {
        Self::Conversion {
            message: message.into(),
        }
    }
}
