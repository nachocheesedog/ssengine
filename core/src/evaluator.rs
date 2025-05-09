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

    // -- Operator helpers omitted for brevity --
}
