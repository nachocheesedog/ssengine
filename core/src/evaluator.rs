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

    // Resolve a reference and get its value
    pub fn resolve_reference(&mut self, r: &Reference) -> Result<CellValue, EngineError> {
        match r {
            Reference::Cell(addr) => self.resolve_cell_value(self.current_sheet, addr),
            Reference::SheetCell { sheet, address } => self.resolve_cell_value(sheet, address),
            Reference::Range { start, end } => self.resolve_cell_value(self.current_sheet, start),
            Reference::SheetRange { sheet, start, end } => self.resolve_cell_value(sheet, start),
        }
    }

    fn resolve_cell_value(&mut self, sheet: &str, addr: &CellAddress) -> Result<CellValue, EngineError> {
        if self.is_circular(sheet, addr) {
            return Err(EngineError::CircularReference(format!("Circular reference detected at {}!{}", sheet, addr.to_a1())));
        }
        self.push_cell(sheet, addr.clone());
        let res = self.workbook.get_cell_value(sheet, addr.row, addr.col);
        self.pop_cell();
        res
    }
}

// Main evaluator struct
pub struct Evaluator {
    function_registry: FunctionRegistry,
}

impl Evaluator {
    /// Create a new Evaluator with default functions registered
    pub fn new() -> Self {
        let mut eval = Evaluator { function_registry: FunctionRegistry::new() };
        eval.function_registry.register_defaults();
        eval
    }

    /// Evaluate a formula string by parsing to AST and evaluating
    pub fn evaluate_formula(&self, workbook: &Workbook, sheet: &str, formula: &str) -> Result<CellValue, EngineError> {
        let ast = crate::parser::parse_formula(formula)?;
        let mut ctx = EvaluationContext::new(workbook, sheet, CellAddress::from_a1(formula)?);
        self.evaluate(&ast, &mut ctx)
    }

    /// Evaluate an AST node
    pub fn evaluate(&self, node: &AstNode, context: &mut EvaluationContext) -> Result<CellValue, EngineError> {
        match node {
            AstNode::Literal(lit) => self.evaluate_literal(lit),
            AstNode::Reference(r) => context.resolve_reference(r),
            AstNode::BinaryOp{op,left,right} => self.evaluate_binary_op(op, left, right, context),
            AstNode::UnaryOp{op,operand} => self.evaluate_unary_op(op, operand, context),
            AstNode::FunctionCall{name,args} => self.evaluate_function(name, args, context),
        }
    }

    fn evaluate_literal(&self, lit: &Literal) -> Result<CellValue, EngineError> {
        Ok(match lit {
            Literal::Number(n) => CellValue::Number(*n),
            Literal::Text(s) => CellValue::Text(s.clone()),
            Literal::Boolean(b) => CellValue::Boolean(*b),
            Literal::Error(e) => CellValue::Error(e.clone()),
        })
    }

    fn evaluate_binary_op(&self, op: &BinaryOperator, l: &AstNode, r: &AstNode, ctx: &mut EvaluationContext) -> Result<CellValue, EngineError> {
        let lv = self.evaluate(l, ctx)?;
        let rv = self.evaluate(r, ctx)?;
        match op {
            BinaryOperator::Add => self.add(&lv, &rv),
            BinaryOperator::Subtract => self.subtract(&lv, &rv),
            BinaryOperator::Multiply => self.multiply(&lv, &rv),
            BinaryOperator::Divide => self.divide(&lv, &rv),
            BinaryOperator::Power => self.power(&lv, &rv),
            BinaryOperator::Equal => self.equal(&lv, &rv),
            BinaryOperator::NotEqual => self.not_equal(&lv, &rv),
            BinaryOperator::LessThan => self.less_than(&lv, &rv),
            BinaryOperator::LessThanOrEqual => self.less_than_or_equal(&lv, &rv),
            BinaryOperator::GreaterThan => self.greater_than(&lv, &rv),
            BinaryOperator::GreaterThanOrEqual => self.greater_than_or_equal(&lv, &rv),
            BinaryOperator::Concat => self.concatenate(&lv, &rv),
        }
    }

    fn evaluate_unary_op(&self, op: &UnaryOperator, node: &AstNode, ctx: &mut EvaluationContext) -> Result<CellValue, EngineError> {
        let v = self.evaluate(node, ctx)?;
        match op {
            UnaryOperator::Positive => Ok(v),
            UnaryOperator::Negative => self.negate(&v),
            UnaryOperator::Percent => self.percent(&v),
        }
    }

    fn evaluate_function(&self, name: &str, args: &[AstNode], ctx: &mut EvaluationContext) -> Result<CellValue, EngineError> {
        let mut vals = Vec::new();
        for a in args { vals.push(self.evaluate(a, ctx)?); }
        self.function_registry.call(name, &vals)
    }

    // Operator helper methods
    fn add(&self, left: &CellValue, right: &CellValue) -> Result<CellValue, EngineError> {
        match (left, right) {
            (CellValue::Number(a), CellValue::Number(b)) => Ok(CellValue::Number(a + b)),
            (CellValue::Blank, CellValue::Number(b)) => Ok(CellValue::Number(*b)),
            (CellValue::Number(a), CellValue::Blank) => Ok(CellValue::Number(*a)),
            (CellValue::Error(e), _) | (_, CellValue::Error(e)) => Ok(CellValue::Error(e.clone())),
            _ => Ok(CellValue::Error(CellError::InvalidValue)),
        }
    }

    fn subtract(&self, left: &CellValue, right: &CellValue) -> Result<CellValue, EngineError> {
        match (left, right) {
            (CellValue::Number(a), CellValue::Number(b)) => Ok(CellValue::Number(a - b)),
            (CellValue::Blank, CellValue::Number(b)) => Ok(CellValue::Number(-b)),
            (CellValue::Number(a), CellValue::Blank) => Ok(CellValue::Number(*a)),
            (CellValue::Error(e), _) | (_, CellValue::Error(e)) => Ok(CellValue::Error(e.clone())),
            _ => Ok(CellValue::Error(CellError::InvalidValue)),
        }
    }

    fn multiply(&self, left: &CellValue, right: &CellValue) -> Result<CellValue, EngineError> {
        match (left, right) {
            (CellValue::Number(a), CellValue::Number(b)) => Ok(CellValue::Number(a * b)),
            (CellValue::Blank, _) | (_, CellValue::Blank) => Ok(CellValue::Number(0.0)),
            (CellValue::Error(e), _) | (_, CellValue::Error(e)) => Ok(CellValue::Error(e.clone())),
            _ => Ok(CellValue::Error(CellError::InvalidValue)),
        }
    }

    fn divide(&self, left: &CellValue, right: &CellValue) -> Result<CellValue, EngineError> {
        match (left, right) {
            (CellValue::Number(a), CellValue::Number(b)) => if *b == 0.0 {
                Ok(CellValue::Error(CellError::DivisionByZero))
            } else { Ok(CellValue::Number(a / b)) },
            (CellValue::Error(e), _) | (_, CellValue::Error(e)) => Ok(CellValue::Error(e.clone())),
            _ => Ok(CellValue::Error(CellError::InvalidValue)),
        }
    }

    fn power(&self, left: &CellValue, right: &CellValue) -> Result<CellValue, EngineError> {
        match (left, right) {
            (CellValue::Number(a), CellValue::Number(b)) => {
                let r = a.powf(*b);
                if r.is_nan() || r.is_infinite() {
                    Ok(CellValue::Error(CellError::InvalidNumber))
                } else {
                    Ok(CellValue::Number(r))
                }
            },
            (CellValue::Error(e), _) | (_, CellValue::Error(e)) => Ok(CellValue::Error(e.clone())),
            _ => Ok(CellValue::Error(CellError::InvalidValue)),
        }
    }

    fn negate(&self, v: &CellValue) -> Result<CellValue, EngineError> {
        match v {
            CellValue::Number(n) => Ok(CellValue::Number(-n)),
            CellValue::Blank => Ok(CellValue::Number(0.0)),
            CellValue::Error(e) => Ok(CellValue::Error(e.clone())),
            _ => Ok(CellValue::Error(CellError::InvalidValue)),
        }
    }

    fn percent(&self, v: &CellValue) -> Result<CellValue, EngineError> {
        match v {
            CellValue::Number(n) => Ok(CellValue::Number(n / 100.0)),
            CellValue::Blank => Ok(CellValue::Number(0.0)),
            CellValue::Error(e) => Ok(CellValue::Error(e.clone())),
            _ => Ok(CellValue::Error(CellError::InvalidValue)),
        }
    }

    fn equal(&self, left: &CellValue, right: &CellValue) -> Result<CellValue, EngineError> {
        use CellValue::*;
        let res = match (left, right) {
            (Number(a), Number(b)) => a == b,
            (Text(a), Text(b)) => a == b,
            (Boolean(a), Boolean(b)) => a == b,
            (Blank, Blank) => true,
            (Blank, Number(b)) => *b == 0.0,
            (Number(a), Blank) => *a == 0.0,
            (Blank, Text(b)) => b.is_empty(),
            (Text(a), Blank) => a.is_empty(),
            (Error(_), _) | (_, Error(_)) => return Ok(CellValue::Error(left.as_error().clone())),
            _ => false,
        };
        Ok(CellValue::Boolean(res))
    }

    fn not_equal(&self, l: &CellValue, r: &CellValue) -> Result<CellValue, EngineError> {
        let CellValue::Boolean(b) = self.equal(l, r)? else { return Err(EngineError::EvaluationError("Invalid comparison".into())) };
        Ok(CellValue::Boolean(!b))
    }

    fn less_than(&self, left: &CellValue, right: &CellValue) -> Result<CellValue, EngineError> {
        use CellValue::*;
        let res = match (left, right) {
            (Number(a), Number(b)) => a < b,
            (Text(a), Text(b)) => a.to_lowercase() < b.to_lowercase(),
            (Blank, Number(b)) => 0.0 < *b,
            (Number(a), Blank) => *a < 0.0,
            (Error(_), _) | (_, Error(_)) => return Ok(CellValue::Error(left.as_error().clone())),
            _ => return Ok(CellValue::Error(CellError::InvalidValue)),
        };
        Ok(CellValue::Boolean(res))
    }

    fn greater_than(&self, l: &CellValue, r: &CellValue) -> Result<CellValue, EngineError> {
        self.less_than(r, l)
    }

    fn less_than_or_equal(&self, l: &CellValue, r: &CellValue) -> Result<CellValue, EngineError> {
        let CellValue::Boolean(b) = self.less_than(l, r)? else { return Err(EngineError::EvaluationError("Invalid comparison".into())) };
        Ok(CellValue::Boolean(b || matches!(self.equal(l, r)?, CellValue::Boolean(true))))
    }

    fn greater_than_or_equal(&self, l: &CellValue, r: &CellValue) -> Result<CellValue, EngineError> {
        self.less_than_or_equal(r, l)
    }

    fn concatenate(&self, left: &CellValue, right: &CellValue) -> Result<CellValue, EngineError> {
        let to_str = |v: &CellValue| match v {
            CellValue::Text(s) => Ok(s.clone()),
            CellValue::Number(n) => Ok(n.to_string()),
            CellValue::Boolean(b) => Ok(b.to_string()),
            CellValue::Blank => Ok(String::new()),
            CellValue::Error(e) => Err(e.clone()),
            _ => Err(CellError::InvalidValue),
        };
        Ok(CellValue::Text(format!("{}{}", to_str(left)?, to_str(right)?)))
    }
}
