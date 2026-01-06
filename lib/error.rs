use thiserror::Error;

#[derive(Error, Debug)]
pub enum ExcelifyError {
    #[error("Sheet not found at index {0}")]
    SheetNotFound(usize),

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Invalid cell reference: {0}")]
    InvalidCellReference(String),
}

pub type Result<T> = std::result::Result<T, ExcelifyError>;
