// ssengine-core/src/functions.rs
// Registry and implementation of spreadsheet functions

use std::collections::HashMap;
use crate::model::CellValue;
use crate::error::{EngineError, CellError};

// Function signature for spreadsheet functions
pub type FunctionImpl = fn(args: &[CellValue]) -> Result<CellValue, EngineError>;

// Registry of available functions
pub struct FunctionRegistry {
    functions: HashMap<String, FunctionImpl>,
}

impl FunctionRegistry {
    pub fn new() -> Self {
        let mut registry = FunctionRegistry {
            functions: HashMap::new(),
        };
        
        // Register built-in functions
        registry.register_defaults();
        
        registry
    }
    
    // Register a function with the registry
    pub fn register(&mut self, name: &str, implementation: FunctionImpl) {
        self.functions.insert(name.to_uppercase(), implementation);
    }
    
    // Look up a function by name
    pub fn get(&self, name: &str) -> Option<&FunctionImpl> {
        self.functions.get(&name.to_uppercase())
    }
    
    // Register all default functions
    fn register_defaults(&mut self) {
        // Math functions
        self.register("SUM", sum);
        self.register("AVERAGE", average);
        
        // Logical functions
        self.register("IF", if_func);
        
        // Other functions will be added as we implement them
    }
}

// Function implementations

// SUM function
fn sum(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let mut total = 0.0;
    
    for arg in args {
        match arg {
            CellValue::Number(n) => total += n,
            CellValue::Text(_) => return Err(EngineError::EvaluationError("Cannot sum text values".into())),
            CellValue::Boolean(b) => total += if *b { 1.0 } else { 0.0 },
            CellValue::Blank => {}, // Ignore blank cells
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Formulas should be evaluated before using in functions".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    Ok(CellValue::Number(total))
}

// AVERAGE function
fn average(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let mut total = 0.0;
    let mut count = 0;
    
    for arg in args {
        match arg {
            CellValue::Number(n) => {
                total += n;
                count += 1;
            },
            CellValue::Boolean(b) => {
                total += if *b { 1.0 } else { 0.0 };
                count += 1;
            },
            CellValue::Text(_) => return Err(EngineError::EvaluationError("Cannot average text values".into())),
            CellValue::Blank => {}, // Ignore blank cells
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Formulas should be evaluated before using in functions".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    if count == 0 {
        return Err(EngineError::EvaluationError("AVERAGE requires at least one numeric value".into()));
    }
    
    Ok(CellValue::Number(total / count as f64))
}

// IF function
fn if_func(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError("IF requires exactly 3 arguments".into()));
    }
    
    let condition = match &args[0] {
        CellValue::Boolean(b) => *b,
        CellValue::Number(n) => *n != 0.0,
        CellValue::Text(s) => !s.is_empty(),
        CellValue::Blank => false,
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Formulas should be evaluated before using in functions".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    if condition {
        Ok(args[1].clone())
    } else {
        Ok(args[2].clone())
    }
}
