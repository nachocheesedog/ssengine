// ssengine-core/src/evaluator.rs
// Formula evaluation engine

use crate::ast::{AstNode, Literal, Reference, BinaryOperator, UnaryOperator};
use crate::error::{EngineError, CellError};
use crate::model::{Workbook, Sheet, CellAddress, CellValue};
use crate::functions::FunctionRegistry;

// Evaluation context for resolving cell references and tracking state
pub struct EvaluationContext<'a> {
    pub workbook: &'a Workbook,
    pub current_sheet: &'a str,
    pub current_cell: CellAddress,
    
    // Track cells being evaluated to detect circular references
    evaluating_cells: Vec<(String, CellAddress)>,
}

impl<'a> EvaluationContext<'a> {
    pub fn new(workbook: &'a Workbook, sheet: &'a str, cell: CellAddress) -> Self {
        EvaluationContext {
            workbook,
            current_sheet: sheet,
            current_cell: cell,
            evaluating_cells: Vec::new(),
        }
    }
    
    // Check for circular references
    pub fn is_circular(&self, sheet: &str, cell: &CellAddress) -> bool {
        self.evaluating_cells.contains(&(sheet.to_string(), cell.clone()))
    }
    
    // Push a cell being evaluated
    pub fn push_cell(&mut self, sheet: &str, cell: CellAddress) {
        self.evaluating_cells.push((sheet.to_string(), cell));
    }
    
    // Pop a cell from the evaluation stack
    pub fn pop_cell(&mut self) {
        self.evaluating_cells.pop();
    }
}

// Main evaluator struct
pub struct Evaluator {
    function_registry: FunctionRegistry,
}

impl Evaluator {
    pub fn new() -> Self {
        Evaluator {
            function_registry: FunctionRegistry::new(),
        }
    }
    
    // Evaluate an AST node
    pub fn evaluate(&self, node: &AstNode, context: &mut EvaluationContext) -> Result<CellValue, EngineError> {
        match node {
            // Basic implementation - will be expanded upon
            AstNode::Literal(lit) => match lit {
                Literal::Number(n) => Ok(CellValue::Number(*n)),
                Literal::Text(s) => Ok(CellValue::Text(s.clone())),
                Literal::Boolean(b) => Ok(CellValue::Boolean(*b)),
                Literal::Error(e) => Err(EngineError::EvaluationError(format!("Error literal: {:?}", e))),
            },
            _ => Err(EngineError::NotImplemented("AST node evaluation".to_string())),
        }
    }
}
