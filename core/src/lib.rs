// ssengine-core library

// Module declarations
pub mod model;
pub mod ast;
pub mod error;
pub mod evaluator;
pub mod functions;
pub mod parser;

// Re-export key types
pub use model::{Cell, Sheet, Workbook};
pub use error::EngineError;

// Create a new workbook
pub fn new_workbook() -> Workbook {
    Workbook::new()
}
