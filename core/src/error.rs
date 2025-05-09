// ssengine-core/src/error.rs
use thiserror::Error;

#[derive(Error, Debug, Clone, PartialEq)]
pub enum EngineError {
    #[error("Parse error: {0}")]
    ParseError(String),

    #[error("Evaluation error: {0}")]
    EvaluationError(String),

    #[error("Circular reference detected: {0}")]
    CircularReference(String),

    #[error("Invalid cell reference: {0}")]
    InvalidReference(String),

    #[error("Unknown function: {0}")]
    UnknownFunction(String),
    
    #[error("Internal error: {0}")]
    Internal(String),

    #[error("Not implemented: {0}")]
    NotImplemented(String),
}

#[derive(Error, Debug, Clone, PartialEq)]
pub enum CellError {
    #[error("#DIV/0!")]
    DivisionByZero,
    
    #[error("#VALUE!")]
    InvalidValue,
    
    #[error("#REF!")]
    InvalidReference,
    
    #[error("#NAME?")]
    NameNotFound,
    
    #[error("#NUM!")]
    InvalidNumber,
    
    #[error("#N/A")]
    NotAvailable,
}
